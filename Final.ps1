# Unblock Tosca execution script
Unblock-File -Path "C:\Users\TE597784\OneDrive - TE Connectivity\Desktop\PipelineFiles\tosca_execution_client.ps1"

# Navigate to script directory
cd "C:\Users\TE597784\OneDrive - TE Connectivity\Desktop\PipelineFiles"

# Define log file for capturing Tosca execution output
$LogFile = "C:\Users\TE597784\OneDrive - TE Connectivity\Desktop\PipelineFiles\Results\tosca_execution_log.txt"

# Run Tosca execution client and capture output
$ExecutionOutput = ./tosca_execution_client.ps1 `
    -toscaServerUrl "https://toscadi-prod.connect.te.com" `
    -eventsConfigFilePath "C:\Users\TE597784\OneDrive - TE Connectivity\Desktop\PipelineFiles\uniqueid.json" `
    -projectName "ToscaDI_Prod" `
    -resultsFolderPath "C:\Users\TE597784\OneDrive - TE Connectivity\Desktop\PipelineFiles\Results" `
    -clientId "V0Im8oW3MUSN1OI1X0Ih0Q" `
    -clientSecret "xBspNBQS0Eyk5jwhAAos5AA1R-2DYbxUuqS7fsdvxg0g" `
    -clientTimeout 300000  *>&1 | Tee-Object -FilePath $LogFile

# Extract the result file path from the log
$ResultFile = ""
$LogContent = Get-Content $LogFile
foreach ($Line in ($LogContent | Sort-Object -Descending)) {
    if ($Line -match 'Finished writing execution results to file "(.*?)"') {
        $ResultFile = $matches[1]
        break
    }
}

# Initialize counters
$Passed = 0
$Failed = 0
$Unknown = 0
$Total = 0
$BodyIsHtml = $false

# Check if result.xml exists
if ((-not [string]::IsNullOrEmpty($ResultFile)) -and (Test-Path $ResultFile)) {
    # Parse the XML results file
    [xml]$xmlResults = Get-Content $ResultFile
    $htmlTable = ""

    foreach ($tc in $xmlResults.testsuites.testsuite.testcase) {
        $Total++  # Increment total test case count
        $tcName = $tc.name
        $tcLog = $tc.log

        # Determine test case status
        if ($tcLog -match "^\s*\+ Passed") {
            $status = "Passed"
            $Passed++
        }
        elseif ($tcLog -match "^\s*\- Failed") {
            $status = "Failed"
            $Failed++
        }
        else {
            $status = "Unknown"
            $Unknown++
        }

        # Append row to HTML table
        $htmlTable += "<tr><td>$tcName</td><td>$status</td></tr>`n"
    }

    # ✅ **Summary Table (at the top of the email)**
    $SummaryTable = @"
    <table border="1" style="border-collapse: collapse; width: 50%; margin: auto;">
        <thead style="background-color: blue; color: white;">
            <tr>
                <th>Category</th>
                <th>Count</th>
            </tr>
        </thead>
        <tbody>
            <tr><td><b>Total Test Cases</b></td><td>$Total</td></tr>
            <tr><td style="color: green;"><b>Passed</b></td><td>$Passed</td></tr>
            <tr><td style="color: red;"><b>Failed</b></td><td>$Failed</td></tr>
            <tr><td style="color: gray;"><b>Unknown</b></td><td>$Unknown</td></tr>
        </tbody>
    </table>
"@

    # ✅ **Construct Final Email Body**
    $Body = @"
<html>
<body>
    <br>
    $SummaryTable
    <br><br>
    <table border="1" style="border-collapse: collapse; width: 100%;">
        <thead style="background-color: blue; color: white;">
            <tr>
                <th>TCName</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
            $htmlTable
        </tbody>
    </table>
</body>
</html>
"@
    $Attachment = $ResultFile
    $BodyIsHtml = $true
}
else {
    # Fallback: include log file contents if results file isn't available
    $Body = "Tosca execution did not generate a valid result.xml. Below is the execution log:`n`n" + (Get-Content $LogFile -Raw)
    $Attachment = $LogFile
}

# Bypass SSL certificate validation for SMTP
Add-Type -TypeDefinition @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class SSLValidator {
    public static void OverrideValidation() {
        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
    }
}
"@
[SSLValidator]::OverrideValidation()


# Email Configuration (No Authentication Required)
$SMTPServer = "terelay.tycoelectronics.net"  # Replace with actual SMTP server
$SMTPPort = 25  # Using port 25 as per your setup
$From = "gourav.mishra@te.com"  # Replace with sender email
$To = "sri.phanindra.k@te.com"  # Replace with recipient email
$Subject = "Tosca Execution Results - Data Lake Daily"

# Send Email
if ($BodyIsHtml) {
    Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $From -To $To -Subject $Subject -Body $Body -Attachments $Attachment -BodyAsHtml
}
else {
    Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $From -To $To -Subject $Subject -Body $Body -Attachments $Attachment
}

Read-Host "Press Enter to exit"
