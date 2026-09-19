# Publisher to PDF Converter

Converts Microsoft Publisher (`.pub`) files to PDFs by recursively scanning an entire Windows drive. The default drive is `C:\`. Full-drive scanning is the only mode; there are no file, folder, or recursion options.

## Requirements

- Windows with Windows PowerShell 5.1 (`powershell.exe`).
- Microsoft Publisher installed and working, with the Office and Publisher interop assemblies available to PowerShell.
- Permission to read source files, save PDFs in their folders, and create logs in the script folder.

## Usage

1. Keep `convertpubtopdf.ps1` and `run_converter.bat` together in a writable folder.
2. To scan a different drive, edit the constant at the top of `convertpubtopdf.ps1`:

   ```powershell
   # Change 'C:\' to 'D:\', 'E:\', etc. to scan a different entire drive.
   Set-Variable -Name DriveRoot -Value 'C:\' -Option Constant
   ```

3. Double-click `run_converter.bat` in your normal Windows user session. The launcher does not request administrator privileges. Avoid **Run as administrator**, especially with Store-installed Publisher, which may not be available to an elevated process.
4. Wait for the completion message, then check the log beside the scripts. The launcher pauses so you can read its output.

The script switches its working directory to the configured drive root. You can launch the batch file from any folder. Scanning a whole drive can take a long time.

To run directly without the batch file's pause, open a normal command prompt and run:

```bat
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "C:\path\to\convertpubtopdf.ps1"
```

## Output and scan behavior

- Each PDF is saved beside its source: `D:\Documents\Brochure.pub` becomes `D:\Documents\Brochure.pdf`.
- Source Publisher files are retained. Existing PDFs are skipped, not overwritten or checked for freshness.
- Hidden and system entries are included. Only `.pub` files are converted.
- Inaccessible folders are logged, and the scan continues elsewhere. Administrator access does not guarantee access to every file.
- Directory junctions, symbolic links, and other directory reparse points are logged and skipped to avoid loops or leaving the selected drive.
- Individual conversion failures are logged, and processing continues with other files.

## Logs and troubleshooting

Each run creates `conversion_log_<timestamp>_<process-id>.txt` beside the PowerShell script. It records successful conversions with source and destination paths, existing-PDF skips, skipped directory links, scan errors, conversion errors, and cleanup errors. The final summary includes counts and the exit code.

The console shows conversion totals. Individual file, folder-access, and cleanup errors are recorded only in the log, so inaccessible folders do not flood the console. Fatal errors that stop the scan still appear in the console.

Exit code `0` means the scan completed without reported errors; existing-PDF and directory-link skips do not count as errors. Exit code `2` means the scan completed with individual conversion or scan/cleanup issues; the launcher reports completion and points to the log. Exit code `1` means a fatal error stopped the run. If the log cannot be created, the script reports the error in the console and stops before converting files.

The script checks Publisher's COM registration before loading interop assemblies or scanning the drive. If the log reports `Publisher.Application` is not registered (the original error was `80040154 Class not registered`), first run the converter normally, without **Run as administrator**. The previous launcher forced elevation; Publisher on this machine is Store-installed and successfully activates in the normal user session. If normal launch still fails, open Publisher to complete setup or install/repair Publisher. Having Office interop assemblies installed does not mean Publisher itself is installed.

If Publisher is registered but cannot start, open it manually to complete setup and repair its installation if needed. If the interop assemblies cannot load, repair the Publisher/Office installation with .NET programmability support. For access errors, check folder permissions; protected folders are logged and scanning continues. If a failed export leaves a partial PDF, review and remove that partial file before retrying: any existing PDF is skipped on subsequent runs.

Verified on Windows: PowerShell syntax, Publisher COM startup, and PDF export of a new blank Publisher document. A full-drive conversion scan was not run during verification.
