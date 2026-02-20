# ============================================================================
# Oracle Database Report Generator - STABLE VERSION
# Run this file by right-clicking and "Run with PowerShell"
# ============================================================================

# Set execution policy for current session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Add error handling to prevent crashes
$ErrorActionPreference = "Continue"

try {
    # Configuration
    $OracleHost = "10.2.5.50"
    $OraclePort = 1522
    $ServiceName = "ASVP1"
    $Username = "eway_owner"
    $Password = "oracle123"
    $SmtpServer = "smtp.UKPN.Local"
    $SmtpPort = 25
    $EmailFrom = "ackreport@ukpowernetworks.co.uk"
    $EmailTo = @("srikanth.naik@ukpowernetworks.co.uk")
    
    # Get the full path for the reports directory
    $ReportPath = Join-Path (Get-Location) "Reports"

    Write-Host "============================================================================" -ForegroundColor Cyan
    Write-Host "Oracle Database Report Generator - STABLE VERSION" -ForegroundColor Green
    Write-Host "============================================================================" -ForegroundColor Cyan

    # Load Oracle DLL and ImportExcel module
    try {
        Add-Type -Path ".\Oracle.ManagedDataAccess.dll"
        Write-Host "✓ Oracle DLL loaded" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to load Oracle DLL: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Press Enter to exit..." -ForegroundColor Yellow
        Read-Host
        exit 1
    }

    # Try to import ImportExcel module
    try {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Host "✓ ImportExcel module loaded" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ ImportExcel module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
            Import-Module ImportExcel
            Write-Host "✓ ImportExcel module installed and loaded" -ForegroundColor Green
        } catch {
            Write-Host "❌ Failed to install ImportExcel module: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Press Enter to exit..." -ForegroundColor Yellow
            Read-Host
            exit 1
        }
    }

    # Create directories
    if (!(Test-Path $ReportPath)) {
        New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
    }
    Write-Host "✓ Directories ready: $ReportPath" -ForegroundColor Green

    # Connection string
    $connectionString = "Data Source=$OracleHost`:$OraclePort/$ServiceName;User Id=$Username;Password=$Password;"

    # Test connection first
    Write-Host "Testing database connection..." -ForegroundColor Yellow
    try {
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)
        $connection.Open()
        $testQuery = "SELECT 'Connection OK' as result FROM dual"
        $command = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($testQuery, $connection)
        $testResult = $command.ExecuteScalar()
        $connection.Close()
        Write-Host "✓ Database connection: $testResult" -ForegroundColor Green
    } catch {
        Write-Host "❌ Database connection failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Press Enter to exit..." -ForegroundColor Yellow
        Read-Host
        exit 1
    }

    # Original query that returns comma-separated output
    $workingQuery = @'
select nvl(flow1.flow, 'N/A') || nvl(flow1.version1, '') ||',' ||
       nvl(aud1.sourcefilename, 'N/A') || ',' ||
       nvl(flow1.toparty, 'N/A') || ',' ||
       nvl(flow1.torole, 'N/A') || ',' ||
       nvl(to_char(aud2.eventtimestamp, 'yyyymmddhh24miss'), 'N/A')|| ',' ||
       nvl(to_char(flow2.created, 'yyyymmddhh24miss'), 'N/A') as output
from audit1 aud1
left join flowaudit flow1 on aud1.archive_idtnmsgid = flow1.archive_idtnmsgid
left join (
    select * from audit1 
    where eventtype = 4 and applicationname = 'ELECTRALINK'
) aud2 on aud1.archive_idtnmsgid = aud2.archive_idtnmsgid
left join (
    select * from audit1 where eventtype = 11
) aud3 on aud1.archive_idtnmsgid = aud3.origidtnmsgid
left join flowaudit flow2 on aud3.archive_idtnmsgid = flow2.archive_idtnmsgid
where aud1.eventtype = 1
and aud1.applicationname = upper('PLACEHOLDER_APP')
and aud1.eventtimestamp between sysdate-1 and sysdate
order by output
'@

    # Function to convert comma-separated data to Excel worksheet data
    function Convert-ToExcelData {
        param([System.Data.DataTable]$InputData)
        
        $excelData = @()
        
        foreach ($row in $InputData.Rows) {
            try {
                $outputValue = $row["OUTPUT"].ToString()
                $parts = $outputValue.Split(',')

                while ($parts.Length -lt 6) {
                    $parts += "N/A"
                }

                $sentTimestamp = $parts[4]
                $receivedTimestamp = $parts[5]

                $sentDate = if ($sentTimestamp.Length -ge 8) { $sentTimestamp.Substring(0, 8) } else { $sentTimestamp }
                $receivedDate = if ($receivedTimestamp.Length -ge 8) { $receivedTimestamp.Substring(0, 8) } else { $receivedTimestamp }

                $excelRow = [PSCustomObject]@{
                    Flow = $parts[0]
                    Flow_ID = $parts[1]
                    Sub = $parts[2]
                    Sub_ID = $parts[3]
                    Sent = $sentTimestamp
                    Received = $receivedTimestamp
                    Sent_Date = $sentDate
                    Received_Date = $receivedDate
                }

                $excelData += $excelRow
            } catch {
                Write-Host "⚠️ Error processing row: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        return $excelData
    }

    # Application list
    $apps = @("EMPRSP1_HUB01_19C", "LMPRSP1_HUB01_19C", "SMPRSP1_HUB01_19C")
    $appNames = @("EMPRSP1 Hub Report", "LMPRSP1 Hub Report", "SMPRSP1 Hub Report")

    # Initialize results
    $allReports = @()
    $allWorksheetData = @()
    $excelFile = Join-Path $ReportPath "MPRS_batch_Ack_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

    # Process each application
    for ($i = 0; $i -lt $apps.Length; $i++) {
        $appCode = $apps[$i]
        $appName = $appNames[$i]

        Write-Host "Processing $appName..." -ForegroundColor Yellow

        try {
            $currentQuery = $workingQuery -replace "PLACEHOLDER_APP", $appCode

            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)
            $connection.Open()
            $command = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($currentQuery, $connection)
            $adapter = New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($command)
            $rawDataTable = New-Object System.Data.DataTable
            $rowCount = $adapter.Fill($rawDataTable)
            $connection.Close()

            Write-Host "✓ Retrieved $rowCount raw records for $appName" -ForegroundColor Green

            $reportInfo = New-Object PSObject
            $reportInfo | Add-Member -Type NoteProperty -Name "AppCode" -Value $appCode
            $reportInfo | Add-Member -Type NoteProperty -Name "AppName" -Value $appName
            $reportInfo | Add-Member -Type NoteProperty -Name "RecordCount" -Value $rowCount
            $allReports += $reportInfo

            if ($rowCount -gt 0) {
                Write-Host "Converting to Excel worksheet format..." -ForegroundColor Yellow
                $excelData = Convert-ToExcelData -InputData $rawDataTable

                $worksheetInfo = @{
                    Name = $appCode
                    Data = $excelData
                    RecordCount = $excelData.Count
                }
                $allWorksheetData += $worksheetInfo

                Write-Host "✓ Prepared $($excelData.Count) records for Excel worksheet: $appCode" -ForegroundColor Green
            } else {
                Write-Host "⚠️ No data found for $appName" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "❌ Error processing $appName`: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Calculate totals
    $totalRecords = ($allReports | Measure-Object -Property RecordCount -Sum).Sum

    # Create Excel file with ONLY the 3 data worksheets
    Write-Host "Creating Excel file with 3 data worksheets..." -ForegroundColor Yellow

    if ($allWorksheetData.Count -gt 0) {
        try {
            if (Test-Path $excelFile) {
                Remove-Item $excelFile -Force
            }

            foreach ($worksheet in $allWorksheetData) {
                Write-Host "Adding worksheet: $($worksheet.Name)" -ForegroundColor Yellow
                $worksheet.Data | Export-Excel -Path $excelFile -WorksheetName $worksheet.Name -AutoSize -BoldTopRow -FreezeTopRow
                Write-Host "✓ Worksheet '$($worksheet.Name)' added with $($worksheet.RecordCount) records" -ForegroundColor Green
            }

            Write-Host "✅ Excel file created: $excelFile" -ForegroundColor Green
            $allFiles = @($excelFile)
        } catch {
            Write-Host "❌ Error creating Excel file: $($_.Exception.Message)" -ForegroundColor Red
            $allFiles = @()
        }
    } else {
        Write-Host "⚠️ No data to export to Excel" -ForegroundColor Yellow
        $allFiles = @()
    }

    Write-Host "✓ All applications processed. Total records: $totalRecords" -ForegroundColor Green

    # Dynamic HTML blocks for app cards
    $appCards = ""
    foreach ($report in $allReports) {
        $status = if ($report.RecordCount -gt 0) { "✅ Active" } else { "⚠️ No Data" }
        $color = if ($report.RecordCount -gt 0) { "#00d084" } else { "#ffb547" }

        $appCards += @"
            <tr>
                <td style='padding: 14px 16px; border-bottom: 1px solid #2e3f5f; color: #dce7ff; font-size: 14px;'>$($report.AppName)</td>
                <td style='padding: 14px 16px; border-bottom: 1px solid #2e3f5f; color: #ffffff; font-weight: 700; text-align: center;'>$($report.RecordCount)</td>
                <td style='padding: 14px 16px; border-bottom: 1px solid #2e3f5f; text-align: right;'>
                    <span style='background: $color; color: #082032; font-size: 12px; font-weight: 700; padding: 6px 10px; border-radius: 999px;'>$status</span>
                </td>
            </tr>
"@
    }

    # Stunning + animated email (with Outlook-friendly fallback structure)
    Write-Host "Generating enhanced animated email..." -ForegroundColor Yellow

    $emailBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>MPRS batch Ack report</title>
  <style>
    body { margin:0; padding:0; font-family: 'Segoe UI', Arial, sans-serif; background:#091425; color:#eaf2ff; }
    .wrapper { width:100%; padding:24px 10px; background: radial-gradient(circle at 20% 20%, #1d3f72 0%, #0a1730 38%, #050c1b 100%); }
    .container { max-width:860px; margin:0 auto; border-radius:20px; overflow:hidden; background:#0f1f3d; border:1px solid #2f4d86; box-shadow:0 20px 60px rgba(0,0,0,0.4); }

    .hero { position:relative; padding:36px 30px 28px; text-align:center; background: linear-gradient(120deg, #5b8cff, #8e6bff, #20c1ff); background-size: 200% 200%; animation: gradientMove 9s ease infinite; }
    .hero h1 { margin:0; font-size:34px; line-height:1.2; color:#fff; letter-spacing:0.3px; }
    .hero p { margin:10px 0 0; color:#eef4ff; font-size:15px; }

    .glow-orb { position:absolute; width:160px; height:160px; border-radius:50%; filter: blur(1px); opacity:.45; }
    .orb-left { top:-55px; left:-35px; background:#7b9cff; animation: floatY 5.8s ease-in-out infinite; }
    .orb-right { bottom:-65px; right:-25px; background:#35d5ff; animation: floatY 7.2s ease-in-out infinite reverse; }

    .stats { padding:22px; background:#0c1a34; }
    .stats-grid { width:100%; border-collapse:separate; border-spacing:12px; }
    .stat-card { border-radius:16px; background: linear-gradient(180deg, #1f3768, #14284f); text-align:center; padding:16px 12px; border:1px solid #355a9e; }
    .stat-number { font-size:30px; line-height:1; font-weight:800; color:#90d2ff; text-shadow:0 0 12px rgba(76, 205, 255, .4); animation:pulseNumber 2.4s ease-in-out infinite; }
    .stat-label { margin-top:8px; font-size:12px; color:#d8e6ff; text-transform:uppercase; letter-spacing:1px; }

    .content { padding:26px 22px 8px; }
    .panel { border-radius:16px; border:1px solid #2e436e; background:#101f3f; overflow:hidden; }
    .panel-header { background: linear-gradient(90deg, #203968, #17335f); padding:16px 18px; font-weight:700; color:#dce9ff; font-size:18px; }
    .footer { padding:22px; text-align:center; color:#b4c9f2; font-size:12px; background:#0a1730; border-top:1px solid #1f355f; }

    .cta { margin:18px 22px 24px; background: linear-gradient(100deg, #1d4377, #275ba3); border:1px solid #4473c8; border-radius:12px; padding:14px; text-align:center; font-size:13px; color:#dceaff; }

    @keyframes gradientMove {
      0% { background-position: 0% 50%; }
      50% { background-position: 100% 50%; }
      100% { background-position: 0% 50%; }
    }
    @keyframes floatY {
      0%, 100% { transform: translateY(0); }
      50% { transform: translateY(-16px); }
    }
    @keyframes pulseNumber {
      0%, 100% { transform: scale(1); opacity: 1; }
      50% { transform: scale(1.06); opacity: .9; }
    }

    @media (max-width: 640px) {
      .hero h1 { font-size: 27px; }
      .stat-number { font-size: 24px; }
    }
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="container">
      <div class="hero">
        <div class="glow-orb orb-left"></div>
        <div class="glow-orb orb-right"></div>
        <h1>🚀 MPRS Batch Ack Report</h1>
        <p>Generated on $(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm')</p>
      </div>

      <div class="stats">
        <table class="stats-grid" role="presentation" width="100%">
          <tr>
            <td class="stat-card">
              <div class="stat-number">$totalRecords</div>
              <div class="stat-label">Total Records</div>
            </td>
            <td class="stat-card">
              <div class="stat-number">$($allReports.Count)</div>
              <div class="stat-label">Applications</div>
            </td>
            <td class="stat-card">
              <div class="stat-number">$($allFiles.Count)</div>
              <div class="stat-label">Files Generated</div>
            </td>
          </tr>
        </table>
      </div>

      <div class="content">
        <div class="panel">
          <div class="panel-header">📊 Application Summary</div>
          <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
            <thead>
              <tr>
                <th style="text-align:left; padding:14px 16px; color:#89a8df; font-size:12px; letter-spacing:1px; text-transform:uppercase; border-bottom:1px solid #2e3f5f;">Application</th>
                <th style="text-align:center; padding:14px 16px; color:#89a8df; font-size:12px; letter-spacing:1px; text-transform:uppercase; border-bottom:1px solid #2e3f5f;">Records</th>
                <th style="text-align:right; padding:14px 16px; color:#89a8df; font-size:12px; letter-spacing:1px; text-transform:uppercase; border-bottom:1px solid #2e3f5f;">Status</th>
              </tr>
            </thead>
            <tbody>
$appCards
            </tbody>
          </table>
        </div>
      </div>

      <div class="cta">
        📎 Excel attachment includes worksheets for EMPRSP1_HUB01_19C, LMPRSP1_HUB01_19C, and SMPRSP1_HUB01_19C.
      </div>

      <div class="footer">
        <div style="font-size:14px; margin-bottom:8px;">📧 Automated MPRS Report System</div>
        <div>Database: $OracleHost`:$OraclePort/$ServiceName</div>
        <div style="margin-top:8px; opacity:.85;">Note: Some email clients (especially Outlook desktop) may reduce animations due to security limitations.</div>
      </div>
    </div>
  </div>
</body>
</html>
"@

    # Send email
    if ($allFiles.Count -gt 0) {
        Write-Host "Sending email..." -ForegroundColor Yellow

        try {
            $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
            $smtp.EnableSsl = $false

            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = New-Object System.Net.Mail.MailAddress($EmailFrom)
            $mail.To.Add($EmailTo[0])
            $mail.Subject = "✨ MPRS batch Ack report - $(Get-Date -Format 'yyyy-MM-dd')"
            $mail.Body = $emailBody
            $mail.IsBodyHtml = $true

            foreach ($file in $allFiles) {
                if (Test-Path $file) {
                    $attachment = New-Object System.Net.Mail.Attachment($file)
                    $mail.Attachments.Add($attachment)
                }
            }

            $smtp.Send($mail)
            Write-Host "✅ Email sent successfully!" -ForegroundColor Green

            $mail.Dispose()
            $smtp.Dispose()
        } catch {
            Write-Host "❌ Error sending email: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "============================================================================" -ForegroundColor Cyan
    Write-Host "✅ MPRS BATCH ACK REPORT COMPLETED!" -ForegroundColor Green
    Write-Host "============================================================================" -ForegroundColor Cyan
} catch {
    Write-Host "❌ CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
    Read-Host
}
