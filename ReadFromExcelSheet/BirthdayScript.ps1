Function Global:Send-Email { 
[cmdletbinding()]
Param (
[Parameter(Mandatory=$True,Position=0)]
[String]$Address,
[Parameter(Mandatory=$False,Position=1)]
[String]$Subject = "Swimming",
[Parameter(Mandatory=$False,Position=2)]
[String]$Body = "Pontypool"
      )
Begin {
Clear-Host
# Add-Type -assembly "Microsoft.Office.Interop.Outlook"
    }
Process {
# Create an instance Microsoft Outlook
# $Outlook = New-Object -ComObject Outlook.Application
# $Mail = $Outlook.CreateItem(0)
# $Mail.To = "$Address"
# $Mail.Subject = $Subject
# $Mail.Body =$Body
# # $Mail.HTMLBody = "When is swimming?"
# # $File = "D:\CP\timetable.pdf"
# # $Mail.Attachments.Add($File)
# $Mail.Send()

$From = "sdutt.243@gmail.com"
$To = $Address
$Cc = "YourBoss@YourDomain.com"
#$Attachment = "C:\temp\Some random file.txt"
$Subject = $Subject
$Body = $Body
$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject `
-Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
-Credential (Get-Credential) #-Attachments $Attachment
       } # End of Process section
# End {
# # Section to prevent error message in Outlook
# $Outlook.Quit()
# [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)
# $Outlook = $null
    #} # End of End section!
} # End of function




#Declare the file path and sheet name
$file = "C:\Users\shdutt\Desktop\test.xlsx"
$sheetName = "test"

#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

#Count max row
$rowMax = ($sheet.UsedRange.Rows).count

#Declare the starting positions
$rowName,$colName = 1,1
$rowEmail,$colEmail = 1,2
$rowDoB,$colDoB = 1,3
$rowDoJ,$colDoJ = 1,4
$today = (Get-Date)
#loop to get values and store it
for ($i=1; $i -le $rowMax-1; $i++)
{
$name = $sheet.Cells.Item($rowName+$i,$colName).text
$email = $sheet.Cells.Item($rowEmail+$i,$colEmail).text
if($sheet.Cells.Item($rowDoB+$i,$colDoB).text -ne "NULL")
{
	$dob = [datetime]::ParseExact($sheet.Cells.Item($rowDoB+$i,$colDoB).text,'%d-MMM',[CultureInfo]::InvariantCulture) 
}
else
{
	$dob = $null
}
if($sheet.Cells.Item($rowDoJ+$i,$colDoJ).text -ne "NULL")
{
	$doj = [datetime]::ParseExact($sheet.Cells.Item($rowDoJ+$i,$colDoJ).text,'%d-MMM-yy',[CultureInfo]::InvariantCulture)
}
else{
    $doj=$null
}

if($dob -ne $null)
{
	
	#$result1 = (Get-Date $today).ToShortDateString() -eq (Get-Date $dob).ToShortDateString()
	$result1 =(($today.Day -eq $dob.Day) -and ($today.Month -eq $dob.Month))
	if($result1)
	{
		write-host $name 
		write-host "Entering to send email" -Foreground Blue
		Send-Email -Address $email
	}

}

if($doj -ne $null)
{
	$result2 = (Get-Date $today) -eq (Get-Date $doj)
	#write-host $result2 -Foreground DarkGreen
	if($result2)
	{
		write-host "Entering to send email" -Foreground Blue
		Send-Email -Address $email
	}

}



# Write-Host ("My Name is: "+$name)
# Write-Host ("My Email is: "+$email)
# Write-Host ("My Date Of Birth is: "+$dob)
# Write-Host ("My Date of Joining is : "+$doj)


}



Write-Host("Max Count is : "+ $rowMax)

# catch{

	# write-host "Caught an exception:" -ForegroundColor Red
    # write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    # write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
# }

#close excel file
$objExcel.quit()
