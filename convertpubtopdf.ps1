# Change 'C:\' to 'D:\', 'E:\', etc. to scan a different entire drive.
Set-Variable -Name DriveRoot -Value 'C:\' -Option Constant

# Requires Windows and an installed copy of Microsoft Publisher.
# Always scans the entire drive recursively, including hidden/system entries.
# PDFs are saved beside their source .pub files; existing PDFs are skipped.
# Logs are saved beside this script. Run as administrator for broader access;
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
    param([string]$Level, [string]$Message)
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    $script:log.WriteLine($line)
}

try {
    # Fail before converting anything if the log cannot be created.
    $log = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.Encoding]::UTF8)
    $log.AutoFlush = $true
    Write-Log 'START' "Scanning $DriveRoot recursively. Log: $logPath"

    if ($DriveRoot -notmatch '^[A-Za-z]:\\$') {
        throw "DriveRoot must be a drive root such as C:\ or D:\."
    }
    Push-Location -LiteralPath $DriveRoot
    $locationPushed = $true

    Add-Type -AssemblyName Office
    Add-Type -AssemblyName Microsoft.Office.Interop.Publisher
    $app = New-Object -ComObject Publisher.Application

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
                Write-Log 'INFO' 'If the export left a partial PDF, review and remove it before retrying; existing PDFs are skipped.'
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
        Write-Log 'FATAL' "Scan stopped: $_"
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
    if ($failureCount -gt 0 -or $issueCount -gt 0) { $exitCode = 1 }
    if ($null -ne $log) {
        Write-Log 'SUMMARY' "Found: $foundCount; converted: $successCount; existing PDFs skipped: $skipCount; failed conversions: $failureCount; scan/cleanup issues: $issueCount; exit code: $exitCode."
        $log.Dispose()
    }
}
exit $exitCode
