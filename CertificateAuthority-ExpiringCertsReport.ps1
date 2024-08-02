<#
.SYNOPSIS
    Generate a report from a Certificate Authority listing certificates that have recently expired or will expire soon, and send the report via email using an SMTP relay.

.DESCRIPTION
    This script generates a scheduled report for certificate administrators to monitor certificates that need to be renewed or confirmed for planned expiration. It retrieves certificate data from the Certificate Authority, processes the data to identify certificates that are expired or expiring soon, and sends a summary report via email. In case of errors, the script sends an error notification with a transcript.

.PARAMETER ExpiredCertAge
    Specifies the search window for expired certificates starting from the current date.
    The default value '30' will list certificates that expired between now and 30 days prior.

.PARAMETER ExpiresInDays
    Specifies the search window for certificates that will expire soon, starting from the current date.
    The default value '60' will list certificates that expire between now and 60 days later.

.PARAMETER SendFrom
    The SMTP address from which to send the report email.

.PARAMETER SendTo
    The SMTP address or array of addresses to which the report email will be sent.

.PARAMETER SendErrorTo
    The SMTP address to send the transcript as an email attachment if an error occurs.

.PARAMETER SMTPServer
    The SMTP server to use for sending the email.

.PARAMETER NoCleanup
    If specified, moves the report CSV file to the Log folder at the end of the script instead of deleting it.

.EXAMPLE
    .\CertificateAuthority-ExpiringCertsReport.ps1
    Runs the script with default parameters, sending a report of certificates that expired in the last 30 days or will expire in the next 60 days.

.EXAMPLE
    .\CertificateAuthority-ExpiringCertsReport.ps1 -ExpiredCertAge 45 -ExpiresInDays 90
    Runs the script with custom parameters, sending a report of certificates that expired in the last 45 days or will expire in the next 90 days.

.EXAMPLE
    .\CertificateAuthority-ExpiringCertsReport.ps1 -SendFrom "ca@domain.com" -SendTo "admin@domain.com"
    Runs the script with specified email addresses for sending the report.

.EXAMPLE
    .\CertificateAuthority-ExpiringCertsReport.ps1 -NoCleanup
    Runs the script and moves the report CSV file to the Log folder instead of deleting it.

.NOTES
    Author: Jeff Cloherty [https://github.com/jeffcloherty/]
    Created: 3/9/2023
    Version: 1.1
    Last Updated: 8/1/2024
    Revision History:
    - 1.0: Initial version
    - 1.1: All script variables defined as parameters, added error handling and notifications, added script synopsis, and updated comments for publishing.
#>

[CmdletBinding()]
param (
    [int]$ExpiredCertAge = 30,
    [int]$ExpiresInDays = 60,

    [ValidatePattern('^\S+@\S+\.\S+$')]
    [string]$SendFrom = "",

    [ValidatePattern('^\S+@\S+\.\S+$')]
    [string[]]$SendTo = @(""),

    [ValidatePattern('^\S+@\S+\.\S+$')]
    [string]$SendErrorTo = "",

    [string]$EmailSubject = "CA Certificate Report $(Get-Date -Format "MM-dd-yyyy")",
    [string]$SMTPServer = "",
    [switch]$NoCleanup = $false
)

# Setup logging and transcript
$pathTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logDir = "$PSScriptRoot\Logs"
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory | Out-Null
}
$transcriptPath = if(($MyInvocation.InvocationName).Length -gt ($MyInvocation.MyCommand).Length) {
        "$logDir\$(($MyInvocation.InvocationName).Split('\')[-1])_$pathTimestamp.txt"
    } else {
        "$logDir\$($MyInvocation.MyCommand)_$pathTimestamp.txt"
    }
Start-Transcript -Path $transcriptPath

Write-Output "Script started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Output "Parameters: ExpiredCertAge=$ExpiredCertAge, ExpiresInDays=$ExpiresInDays, SendFrom=$SendFrom, SendTo=$SendTo, SMTPServer=$SMTPServer, NoCleanup=$NoCleanup"

# Main script logic
try {
    # Pull certificate info from CertUtil and convert to PS object
    $certdata = ConvertFrom-Csv (certutil.exe -view log csv)

    $now = Get-Date

    # Check certificate expiration and add Status property for each item
    foreach($certificate in $certdata) {
        try {
            # Parse certificate expiration date
            $certexpiration = [datetime]$certificate.'Certificate Expiration Date'
            # Determine the certificate status
            if ($certexpiration -gt $now.AddDays(-1 * $ExpiredCertAge)) {
                if ($certexpiration -lt $now) {
                    $certstate = 'Expired'
                } elseif ($certexpiration -lt $now.AddDays($ExpiresInDays)) {
                    $certstate = 'Review'
                } else {
                    $certstate = 'Valid'
                }

                # Add Status property to the certificate
                $certificate | Add-Member -NotePropertyName Status -NotePropertyValue $certstate
                $certificate.'Request ID' = [int]$certificate.'Request ID'
            }
        } catch {
            Write-Host $certificate.'Request ID'
        }
    }

    # Create a temporary file
    $tmpFile = [System.IO.Path]::GetTempFileName()

    # Export data to file
    $certdata | Where-Object { $_.Status -ne '' } |
        Sort-Object -Property Status, 'Request ID' |
        Select-Object Status, 'Request ID', 'Requester Name', 'Issued Common Name', 'Certificate Effective Date', 'Certificate Expiration Date', 'Serial Number' |
        Export-Csv -Path $tmpFile -NoTypeInformation

    # Rename to something useful
    $tmpFile = (Rename-Item -Path $tmpFile -NewName "CA-CertReport_$(Get-Date -Format 'MM-dd-yyyy').csv" -PassThru).FullName

    # Format email body
    $Body = @"
Attached is the weekly Certificate Report.
<p>
Certificate Summary
<br>
Current Total: $(($certdata | Where-Object { $_.Status -eq 'Valid' -or $_.Status -eq 'Review' }).Count)
<p>
Expiring within $ExpiresInDays days: $(($certdata | Where-Object { $_.Status -eq 'Review' }).Count)
<table>
    <tr>
        <th>Issued Common Name</th>
        <th>Certificate Expiration Date</th>
        <th>Requester Name</th>
        <th>Request ID</th>
    </tr>
    $(($certdata | Where-Object { $_.Status -eq 'Review' }) | ForEach-Object { "<tr>
    <td>$($_.'Issued Common Name')</td>
    <td>$($_.'Certificate Expiration Date')</td>
    <td>$($_.'Requester Name')</td>
    <td>$($_.'Request ID')</td>
    </tr>" })
</table>
<br>
Expired in past $ExpiredCertAge days: $(($certdata | Where-Object { $_.Status -eq 'Expired' }).Count)
<table>
    <tr>
        <th>Issued Common Name</th>
        <th>Certificate Expiration Date</th>
        <th>Requester Name</th>
        <th>Request ID</th>
    </tr>
    $(($certdata | Where-Object { $_.Status -eq 'Expired' }) | ForEach-Object { "<tr>
    <td>$($_.'Issued Common Name')</td>
    <td>$($_.'Certificate Expiration Date')</td>
    <td>$($_.'Requester Name')</td>
    <td>$($_.'Request ID')</td>
    </tr>" })
</table>
"@

    # Send summary and report via email
    Write-Output "Sending report email"
    if ($SendTo -ne "" -and $SendFrom -ne "") {
      Send-MailMessage -Subject $EmailSubject -Body $Body -SmtpServer $SMTPServer -To $SendTo -From $SendFrom -Attachments $tmpFile -BodyAsHtml
  } else {
      Write-Output "Email sending parameters are not properly configured."
  }
  
}

catch {
    Write-Output "Error: $_"
    Write-Output "Stack Trace: $($_.Exception.StackTrace)"
    Write-Output "Deleting temp report file: $tmpFile"
    Remove-Item $tmpFile
    Write-Output "Script ended at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Saving transcript to $transcriptPath"
    
    # Send PowerShell transcript as error email
    Stop-Transcript
    Send-MailMessage -Subject "Error: $Subject" -Body "Script encountered an error: $_`nStack Trace: $($_.Exception.StackTrace)" -SmtpServer $SMTPServer -To $SendErrorTo -From $SendFrom -Attachments $transcriptPath -BodyAsHtml
    Exit
}

finally {
    if (-not $NoCleanup) {
        # Clean up temporary report file
        Write-Output "Deleting temp report file: $tmpFile"
        Remove-Item $tmpFile
    } else {
        $tmpFile = (Move-Item -Path $tmpFile -Destination "$($PSScriptRoot)\Logs\" -PassThru).FullName
        Write-Output "Report CSV saved to $tmpFile"
    }

    Write-Output "Script ended at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Saving transcript to $transcriptPath"
    Stop-Transcript
}
