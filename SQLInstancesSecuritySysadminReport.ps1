#***********************************************************************
# PowerShell : SQLInstancesSecuritySysadminReport.ps1                  *
#   Function : SQL Instances Security Audit Inventory.                 *
#            :                                                         *
#            : 1) Security Audit Inventory                             *
#            :                                                         *
#***********************************************************************
#                 M O D I F I C A T I O N S                            *
# -- Date -- ---- Name ---- --------- Description -------------------- *
# 11/06/2009 Gabriel Garcia Created.                                   *
#                                                                      *
#***********************************************************************

###
### Example 
### PS C:\> C:\Scripts\SQLInstancesSecuritySysadminReport\SQLInstancesSecuritySysadminReport.ps1
###

### Host Name
$strIPGlobalProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()

### Parameters
$emailfrom = [string]$strIPGlobalProperties.HostName + "@domainname.com"
$emailto = "myemailaddres@domainname.com,anotheremailaddres@domainname.com"
$SMTPServer = "smtp.mailserver.name"

### Create a new Excel object using COM
$objExcelApp = New-Object -ComObject Excel.Application
$objExcelApp.visible = $False
$objExcelApp.DisplayAlerts = $False

### 1) Security Audit Inventory
$objWorkbook = $objExcelApp.Workbooks.Add()
$objWorksheet = $objExcelApp.Worksheets.Item(1)
$intRow = 1
$strOutFileName = "C:\Scripts\SQLInstancesSecuritySysadminReport\Archive\SQLInstancesSecuritySysadminReport$(get-date -f yyyy-MM-dd-HHmmss).xlsx"

### Variable
$CheckTime = get-Date #-Format "yyyy-MM-dd hh:mm:ss"

### Output file variable
$strTxtInFileName = "C:\Scripts\SQLInstancesSecuritySysadminReport\SQL_Servers.txt"
### File to upload to MS SQL table
### $strOutFile   = "\\FileServerName\InputFiles\tblSecurityAuditInventory.xlsx"

########################################
### Create and format column headers ###
########################################

### Instance Level
$objWorksheet.Cells.Item($intRow,1)  = "INSTANCE NAME"

### Login Level
$objWorksheet.Cells.Item($intRow,2)  = "LOGIN"
$objWorksheet.Cells.Item($intRow,3)  = "SYS"
$objWorksheet.Cells.Item($intRow,4)  = "SECURITY"
$objWorksheet.Cells.Item($intRow,5)  = "IS SERVER"
$objWorksheet.Cells.Item($intRow,6)  = "SETUP"
$objWorksheet.Cells.Item($intRow,7)  = "PROCESS"
$objWorksheet.Cells.Item($intRow,8)  = "DISK"
$objWorksheet.Cells.Item($intRow,9)  = "DBCREATOR"
$objWorksheet.Cells.Item($intRow,10) = "BULADMIN"
$objWorksheet.Cells.Item($intRow,11) = "LOGIN_TYPE"
$objWorksheet.Cells.Item($intRow,12) = "CREATE DATE"
$objWorksheet.Cells.Item($intRow,13) = "DATE LAST MODIFIED"
$objWorksheet.Cells.Item($intRow,14) = "DEFAULT DATABASE"
$objWorksheet.Cells.Item($intRow,15) = "DENY WINDOWS LOGIN"
$objWorksheet.Cells.Item($intRow,16) = "HAS ACCESS"
$objWorksheet.Cells.Item($intRow,17) = "IS DISABLED"
$objWorksheet.Cells.Item($intRow,18) = "IS LOCKED"
$objWorksheet.Cells.Item($intRow,19) = "IS PASSWORD EXPIRED"
$objWorksheet.Cells.Item($intRow,20) = "PASSWORD EXPIRATION ENABLED"
$objWorksheet.Cells.Item($intRow,21) = "WINDOWS LOGIN ACCESS TYPE"
$objWorksheet.Cells.Item($intRow,22) = "CHECK TIME"

$range = $objWorksheet.range("A1:V1")
$range.Font.Bold = $True
$range.Interior.ColorIndex = 48
$range.Font.ColorIndex = 34

$intRow ++

#################
### Main Loop ###
#################

### Read thru the contents of the SQL_Servers.txt file
foreach ($instance in get-content $strTxtInFileName) {
 ### This script gets SQL Server database information using PowerShell
 [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null

 ###  Create an SMO connection to the instance
 $srv = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $instance
 $logins = $srv.Logins 

 Write-Host -ForegroundColor Green "Checking SQL Instance" $instance "..."

 ### SQL logins
 Write-Host -ForegroundColor Yellow " Checking SQL logins" $instance "..." 
 
 ForEach ($login in $logins) {
	### use name variable for remaining script
	$name = $login.name
	Write-Host -ForegroundColor Gray "  Checking login" $name "..."	 

	if ($srv.Logins["$name"].IsMember("sysadmin") -eq $True -and $login.IsDisabled -eq $False){	
	
		### Instance level
		$objWorksheet.Cells.Item($intRow, 1) = $instance
		
		### Login level
		$objWorksheet.Cells.Item($intRow, 2)  = $name 
		$objWorksheet.Cells.Item($intRow, 3)  = $srv.Logins["$name"].IsMember("sysadmin")
		$objWorksheet.Cells.Item($intRow, 4)  = $srv.Logins["$name"].IsMember("securityadmin")
		$objWorksheet.Cells.Item($intRow, 5)  = $srv.Logins["$name"].IsMember("serveradmin")
		$objWorksheet.Cells.Item($intRow, 6)  = $srv.Logins["$name"].IsMember("setupadmin")
		$objWorksheet.Cells.Item($intRow, 7)  = $srv.Logins["$name"].IsMember("processdmin")
		$objWorksheet.Cells.Item($intRow, 8)  = $srv.Logins["$name"].IsMember("diskadmin")
		$objWorksheet.Cells.Item($intRow, 9)  = $srv.Logins["$name"].IsMember("dbcreator")
		$objWorksheet.Cells.Item($intRow, 10) = $srv.Logins["$name"].IsMember("bulkadmin")
		$objWorksheet.Cells.Item($intRow, 11) = [string]$login.logintype
		$objWorksheet.Cells.Item($intRow, 12) = $login.CreateDate
		$objWorksheet.Cells.Item($intRow, 13) = $login.DateLastModified
		$objWorksheet.Cells.Item($intRow, 14) = $login.DefaultDatabase
		$objWorksheet.Cells.Item($intRow, 15) = $login.DenyWindowsLogin
		$objWorksheet.Cells.Item($intRow, 16) = $login.HasAccess
		$objWorksheet.Cells.Item($intRow, 17) = $login.IsDisabled
		$objWorksheet.Cells.Item($intRow, 18) = $login.IsLocked
		$objWorksheet.Cells.Item($intRow, 19) = $login.IsPasswordExpired
		$objWorksheet.Cells.Item($intRow, 20) = $login.PasswordExpirationEnabled
		$objWorksheet.Cells.Item($intRow, 21) = [string]$login.WindowsLoginAccessType
		$objWorksheet.Cells.Item($intRow, 22) = $CheckTime

		$intRow ++
	}  ### End if condition
 }
 ### Disconnect from the SQL Server database
 $srv.ConnectionContext.Disconnect()
}
 
$objWorksheet.UsedRange.EntireColumn.AutoFit()
$objWorkbook.SaveAs($strOutFileName)

### Save file to upload to MS SQL table
### $objWorkbook.SaveAs($strOutFile)

$objExcelApp.Quit()

Start-Sleep -s 5

### e-mail report
$EmailSubject = "Security Audit Inventory" 
$emailbody = "Security Audit Inventory"
 
$mailmessage = New-Object system.net.mail.mailmessage
$mailmessage.from = ($emailfrom)
$mailmessage.To.add($emailto)
$mailmessage.Subject = $emailsubject
$mailmessage.Body = $emailbody

$attachment = New-Object System.Net.Mail.Attachment($strOutFileName, 'text/plain')
$mailmessage.Attachments.Add($attachment)

###$mailmessage.IsBodyHTML = $true
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")
$SMTPClient.Send($mailmessage)
