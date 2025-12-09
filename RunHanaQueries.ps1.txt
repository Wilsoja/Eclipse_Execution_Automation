      <#
.SYNOPSIS
  Executes SAP HANA SQL files from a specified directory, collects results as CSV,
  zips them, and emails the archive. Uses hdbuserstore for secure connection.

.DESCRIPTION
  This script iterates through all .sql files in a source folder. For each file,
  it uses the hdbsql command-line tool (part of the SAP HANA Client) to execute
  the SQL against a HANA database specified by a secure hdbuserstore key.
  The output of each query is saved as a separate CSV file in a temporary results
  directory. All generated CSV files are then compressed into a single Zip archive.
  Finally, the Zip archive is emailed to specified recipients using Send-MailMessage.
  Temporary files are cleaned up afterwards.

.NOTES
  Author: Your Automation Guru (via AI -> Gemini 2.5 Pro Experimental [version: 3/25/2025])
  Date:   4/1/25
  Requires:
    - SAP HANA Client installed and in the system PATH or script configured with full path.
    - hdbuserstore key pre-configured for HANA connection.
    - PowerShell 5.1 or later (Standard on Windows 10/11).
    - Appropriate permissions for the user running the script (file system access, network for email).
#>

# --- CONFIGURATION SECTION - EDIT THESE VALUES ---

# --- HANA & SQL Files ---
$HanaUserStoreKey = "YourAlias" # The key you created with hdbuserstore (e.g., "HANA_PROD_REPORTING")
$SqlSourceFolder = "C:\SQL_SCRIPTS" # Folder containing your .sql files
# Optional: Full path to hdbsql if not in system PATH
$HdbsqlPath = "C:\Program Files\sap\hdbclient\hdbsql.exe"

# --- Output & Zipping ---
$BaseOutputFolder = "C:\Temp\HanaResults" # Base directory for temporary output and final zip
$TimestampFormat = "yyyyMMdd_HHmmss" # Format for timestamp in filenames
$ZipFileNamePattern = "Hana_Query_Results_{0}.zip" # {0} will be replaced by the timestamp

# --- Email Settings ---
$EmailTo = "John.Doe@something.com" # Comma-separated list of recipient emails
$EmailFrom = "nobody@something.com" # Sender email address
$EmailSubject = "Execution Results !! - $(Get-Date -Format $TimestampFormat)"
$EmailBody = "Validation Automation Results.  Attached are the results from the scheduled HANA SQL queries executed on $(Get-Date)."
$SmtpServer = "SMTP SERVER" # Your organization's SMTP server address
# Optional: SMTP Port (default is 25)
# $SmtpPort = 587
# Optional: Use SSL for SMTP
# $UseSsl = $true
# Optional: Credentials for SMTP if required (Use SecureString for better security if needed)
# $SmtpCredential = Get-Credential # Prompts interactively - not ideal for automation. Consider secure credential management methods.

# --- END CONFIGURATION SECTION ---

# --- SCRIPT LOGIC ---

# Create timestamped output directory
$Timestamp = Get-Date -Format $TimestampFormat
$OutputFolder = Join-Path -Path $BaseOutputFolder -ChildPath $Timestamp
$ZipFilePath = Join-Path -Path $BaseOutputFolder -ChildPath ($ZipFileNamePattern -f $Timestamp)

try {
    # Create the temporary output directory
    Write-Host "Creating output directory: $OutputFolder"
    $null = New-Item -ItemType Directory -Path $OutputFolder -Force -ErrorAction Stop

    # Find SQL files
    $SqlFiles = Get-ChildItem -Path $SqlSourceFolder -Filter *.sql -ErrorAction Stop
    if (-not $SqlFiles) {
        throw "No .sql files found in '$SqlSourceFolder'."
    }

    Write-Host "Found $($SqlFiles.Count) SQL files to process."
    $ResultFiles = @() # Array to hold paths of generated CSV files

    # Process each SQL file
    foreach ($SqlFile in $SqlFiles) {
        $SqlFilePath = $SqlFile.FullName
        $BaseName = $SqlFile.BaseName
        $OutputCsvPath = Join-Path -Path $OutputFolder -ChildPath "$($BaseName)_Result.csv"

        Write-Host "Processing '$($SqlFile.Name)'..."

        # Construct hdbsql command arguments
        # -U : Use secure user store key
        # -i : Input SQL file
        # -o : Output result file
        # -c ";" : Use semicolon as CSV separator (adjust if needed)
        # -a : Output data only (no headers/footers) - adjust if you WANT headers
        # You might need other options like -qe (quit on error)
        $HdbsqlArgs = @(
            "-U", $HanaUserStoreKey,
            "-I", $SqlFilePath,
            "-o", $OutputCsvPath,
            "-c", ";", # Semicolon Separator for CSV
			"-qe"
                  # "-a" Data only (adjust if you need headers)
            # "-qe"   # Optional: Quit immediately if an error occurs in SQL
        )

        # Execute hdbsql
        Write-Host "Executing: $HdbsqlPath $HdbsqlArgs"
        $process = Start-Process -FilePath $HdbsqlPath -ArgumentList $HdbsqlArgs -Wait -PassThru -NoNewWindow

        # Check execution status
        if ($process.ExitCode -ne 0) {
            Write-Warning "Error executing '$($SqlFile.Name)'. hdbsql exited with code $($process.ExitCode). Check hdbsql logs or run manually for details. Skipping this file."
            # Consider adding more robust error handling here - e.g., read stderr if possible
        } else {
            Write-Host "'$($SqlFile.Name)' executed successfully. Output: $OutputCsvPath"
            $ResultFiles += $OutputCsvPath
        }
    }

    # Check if any results were generated
    if ($ResultFiles.Count -eq 0) {
        throw "No result files were generated. Cannot create zip or send email."
    }

    # Zip the results
    Write-Host "Zipping $($ResultFiles.Count) result file(s) to '$ZipFilePath'..."
    # Compress-Archive needs the SOURCE path (the folder containing the files)
    Compress-Archive -Path $OutputFolder\* -DestinationPath $ZipFilePath -Force -ErrorAction Stop
    Write-Host "Zip file created successfully."

    # Email the zip file
    Write-Host "Sending email to: $($EmailTo -join ', ')"

    # Construct parameters for Send-MailMessage, handling optional parameters
    $MailParams = @{
        To         = $EmailTo
        From       = $EmailFrom
        Subject    = $EmailSubject
        Body       = $EmailBody
        SmtpServer = $SmtpServer
        Attachments = $ZipFilePath
        # ErrorAction = 'Stop' # Make email failure stop the script if desired
    }
    # if ($SmtpPort) { $MailParams.Add('Port', $SmtpPort) }
    # if ($UseSsl) { $MailParams.Add('UseSsl', $true) }
    # if ($SmtpCredential) { $MailParams.Add('Credential', $SmtpCredential) }

    Send-MailMessage @MailParams

    Write-Host "Email sent successfully."

} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    # Optional: Send an error notification email here
    # Exit with a non-zero code to indicate failure to Task Scheduler
    exit 1
} finally {
    # Cleanup: Remove the temporary output directory
    if (Test-Path -Path $OutputFolder) {
        Write-Host "Cleaning up temporary directory: $OutputFolder"
        Remove-Item -Path $OutputFolder -Recurse -Force
    }
    Write-Host "Script finished."
}

# Exit with 0 to indicate success to Task Scheduler
exit 0
