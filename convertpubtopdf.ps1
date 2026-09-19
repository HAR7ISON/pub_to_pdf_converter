# Change 'C:\' to 'D:\', 'E:\', etc. to scan a different entire drive.
Set-Variable -Name DriveRoot -Value 'C:\' -Option Constant

# Requires Windows and an installed copy of Microsoft Publisher.
# Always scans the entire drive recursively, including hidden/system entries.
# PDFs are saved beside their source .pub files; existing PDFs are skipped.
# Logs are saved beside this script. Run in your normal Windows user session;
# protected folders that remain inaccessible are recorded in the log.
$ErrorActionPreference = 'Stop'
$logPath = Join-Path $PSScriptRoot ("conversion_log_{0}_{1}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $PID)
$log = $null
$app = $null
$locationPushed = $false
$foundCount = 0
$successCount = 0
$skipCount = 0
$failureCount = 0
$issueCount = 0
$exitCode = 0

function Write-Log {
    param([string]$Level, [string]$Message, [switch]$LogOnly)
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if (-not $LogOnly -and $Level -ne 'ERROR') { Write-Host $line }
    $script:log.WriteLine($line)
}

try {
    # Fail before converting anything if the log cannot be created.
    $log = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.Encoding]::UTF8)
    $log.AutoFlush = $true
    Write-Log 'START' "Checking conversion prerequisites. Log: $logPath"

    if ($DriveRoot -notmatch '^[A-Za-z]:\\$') {
        throw "DriveRoot must be a drive root such as C:\ or D:\."
    }
    # Interop assemblies alone do not install Publisher or register its COM server.
    # Check the application first so missing Publisher produces an actionable error.
    $publisherType = [System.Type]::GetTypeFromProgID('Publisher.Application')
    if ($null -eq $publisherType -or $publisherType.GUID -eq [Guid]::Empty) {
        throw 'Microsoft Publisher is not registered in this Windows session (Publisher.Application). Run the converter normally, without Run as administrator: Store-installed Publisher may be unavailable in an elevated session. If it still fails, open Publisher once to complete setup, or install/repair Publisher. Office/Publisher interop assemblies alone cannot convert .pub files.'
    }

    try {
        Add-Type -AssemblyName Office
        Add-Type -AssemblyName Microsoft.Office.Interop.Publisher
    } catch {
        throw "Publisher interop assemblies could not load. Repair the Publisher/Office installation with its .NET programmability support. Details: $($_.Exception.Message)"
    }
    try {
        $app = New-Object -ComObject Publisher.Application
    } catch {
        throw "Microsoft Publisher could not start. Run the converter without Run as administrator, especially with Store-installed Publisher. Open Publisher manually to complete setup, then retry; repair its Office installation if necessary. Details: $($_.Exception.Message)"
    }

    Push-Location -LiteralPath $DriveRoot
    $locationPushed = $true
    Write-Log 'INFO' "Scanning $DriveRoot recursively."

    # Walk one directory at a time so an inaccessible folder does not stop
    # other folders from being scanned or require storing the entire drive.
    $pending = [System.Collections.Generic.Stack[string]]::new()
    $pending.Push($DriveRoot)
    while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        $scanErrors = @()
        $entries = @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors)
        foreach ($scanError in $scanErrors) {
            $issueCount++
            Write-Log 'ERROR' "Cannot fully scan '$directory': $scanError"
        }

        foreach ($file in $entries) {
            if ($file.PSIsContainer) {
                if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    # Junctions/symbolic links can loop or lead outside this drive.
                    Write-Log 'SKIP' "Directory link not followed: $($file.FullName)"
                } else {
                    $pending.Push($file.FullName)
                }
                continue
            }
            if ($file.Extension -ine '.pub') { continue }

            $foundCount++
            $sourcePath = $file.FullName
            $pdfPath = [System.IO.Path]::ChangeExtension($sourcePath, '.pdf')
            $doc = $null
            try {
                if (Test-Path -LiteralPath $pdfPath) {
                    $skipCount++
                    Write-Log 'SKIP' "PDF already exists: '$sourcePath' -> '$pdfPath'"
                    continue
                }

                $doc = $app.Open($sourcePath)
                if ($null -eq $doc) { throw 'Publisher did not return a document.' }
                $doc.ExportAsFixedFormat([Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF, $pdfPath)
                $pdf = Get-Item -LiteralPath $pdfPath
                if ($pdf.PSIsContainer -or $pdf.Length -eq 0) {
                    throw 'Export did not produce a nonempty PDF file.'
                }
                $successCount++
                Write-Log 'SUCCESS' "Converted '$sourcePath' -> '$pdfPath'"
            } catch {
                $failureCount++
                Write-Log 'ERROR' "Conversion failed for '$sourcePath' -> '$pdfPath': $_"
                Write-Log 'INFO' 'If the export left a partial PDF, review and remove it before retrying; existing PDFs are skipped.' -LogOnly
            } finally {
                if ($null -ne $doc) {
                    try {
                        $doc.Close()
                    } catch {
                        $issueCount++
                        Write-Log 'ERROR' "Could not close '$sourcePath': $_"
                    } finally {
                        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
                        $doc = $null
                    }
                }
            }
        }
    }
} catch {
    $exitCode = 1
    if ($null -ne $log) {
        Write-Log 'FATAL' "Conversion stopped: $_"
    } else {
        Write-Host "Cannot create log '$logPath': $_" -ForegroundColor Red
    }
} finally {
    if ($null -ne $app) {
        try {
            $app.Quit()
        } catch {
            $issueCount++
            Write-Log 'ERROR' "Could not quit Publisher: $_"
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        }
    }
    if ($locationPushed) { Pop-Location }
    # A completed scan with individual issues is different from a fatal stop.
    if ($exitCode -eq 0 -and ($failureCount -gt 0 -or $issueCount -gt 0)) { $exitCode = 2 }
    if ($null -ne $log) {
        Write-Log 'SUMMARY' "Found: $foundCount; converted: $successCount; existing PDFs skipped: $skipCount; failed conversions: $failureCount; scan/cleanup issues: $issueCount; exit code: $exitCode." -LogOnly
        Write-Host "Converted: $successCount; existing PDFs skipped: $skipCount; not converted: $failureCount."
        $log.Dispose()
    }
}
exit $exitCode
