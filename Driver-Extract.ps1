#Import Reg Key to set cleanMgr options to clean unused drivers 
# TODO: -Add switch to disable sending email notification
#       -Build XML file for configuration
# Import XML Configuration Settings
[xml]$ConfigFile = Get-Content "$PSScriptRoot\config.xml"

$RootPath = $ConfigFile.Settings.RootPath
if ($ConfigFile.Settings.CompletedNotification -eq "Email")
{
	$SendCompletionEmail = "True"
}

if (!(Test-Path -Path $RootPath))
{
	Write-Verbose -Message "[INFO] 1/2: Unable to connect to driver staging directory at: $RootPath"
	Write-Verbose -Message "[INFO] 2/2: Proceeding using $PSScriptRoot as staging directory, please move driver folder when complete"
	$RootPath = $PSScriptRoot
}

Write-Verbose -Message "[INFO] Running cleanmgr.exe and cleaning up unused driver packages"
New-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Device Driver Packages" -Name "StateFlags0003" -PropertyType Dword -Value 2 -Force
Start-Process cleanmgr -ArgumentList "/sagerun:3" -NoNewWindow -Wait


$CurrentSystemOS = Get-CIMInstance -ClassName Win32_OperatingSystem -NameSpace root\CIMV2 | select -Property OSArchitecture, Version, Caption
$CurrentModel = Get-CIMInstance -ClassName Win32_ComputerSystem -NameSpace root\CIMV2 | select -Property Manufacturer, Model, SystemSKUNumber
$BaseBoardProduct = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct
$Platform = 'ConfigMgr'

switch -wildcard ($CurrentModel.Manufacturer)
{
	"*HP*" {
		$ExtractMake = "Hewlett-Packard"
		$ExtractSKU = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct
	}
	"*Hewlett-Packard*" {
		$ExtractMake = "Hewlett-Packard"
		$ExtractSKU = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct
	}
	"*Dell*" {
		$ExtractMake = "Dell"
		$ExtractSKU = $CurrentModel.SystemSKUNumber
		$CurrentModel.Model = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).Trim()
		$ExtractSku = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).SystemSku
		if ($ExtractSku -eq "")
		{
			$ExtractSku = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct
		}
	}
	"*Lenovo*" {
		$ExtractMake = "Lenovo"
		$ExtractSKU = ((Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).SubString(0, 4)).Trim()
	}
	"*Panasonic*" {
		$ComputerManufacturer = "Panasonic Corporation"
		$ComputerModel = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).Trim()
		$SystemSKU = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct.Trim()
	}
	"*Viglen*" {
		$ComputerManufacturer = "Viglen"
		$ComputerModel = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).Trim()
		$SystemSKU = (Get-WmiObject -Class Win32_BaseBoard | Select-Object -ExpandProperty SKU).Trim()
	}
	"*AZW*" {
		$ComputerManufacturer = "AZW"
		$ComputerModel = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).Trim()
		$SystemSKU = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct.Trim()
	}
	default
	{
		$ExtractMake = $CurrentModel.Manufacturer
		$ExtractSKU = (Get-CIMInstance -ClassName MS_SystemInformation -NameSpace root\WMI).BaseBoardProduct
	}
}

switch -wildcard ($CurrentSystemOS.Caption)
{
	"*Windows 10*" {
		$OSRelease = [version]"10.0"
		$OSName = "Windows 10"
	}
	"*Windows 8.1" {
		$OSRelease = [version]"6.3"
		$OSName = "Windows 8.1"
	}
	"*Windows 8" {
		$OSRelease = [version]"6.2"
		$OSName = "Windows 8"
	}
	"*Windows 7" {
		$OSRelease = [version]"6.1"
		$OSName = "Windows 7"
	}
}

switch -wildcard ($CurrentSystemOS.OSArchitecture)
{
	"64*" {
		$OSArchitecture = "x64"
	}
	"32*" {
		$OSArchitecture = "x86"
	}
}
if ($OSName -eq "Windows 10")
{
	$ReleaseID = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name ReleaseID).ReleaseID
	$OSName = $OSName + " $ReleaseID"
	
}



# Format: Make Model - SKU - OSNameOsArchitecture
$ExtractDriverRootDir = Join-Path -Path "$RootPath" -ChildPath "$ExtractMake $($CurrentModel.Model) - $ExtractSKU - $OSName$OSArchitecture"
$ExtractDriverDir = Join-Path -Path "$ExtractDriverRootDir" -ChildPath "DriverPackage"
$ExtractDriverInfoXML = Join-Path -Path $ExtractDriverRootDir -ChildPath "ModelDetail.xml"

if (!(Test-Path -Path $ExtractDriverDir))
{
	# Folder does not exist, create it
	New-Item -Path $ExtractDriverRootDir -ItemType directory
	New-Item -Path $ExtractDriverDir -ItemType directory
}

$ExtractDriverDir = "`"$ExtractDriverDir`""


# Extract drivers and place in ExtractDriverDir
try
{
	Write-Verbose "[TRY] Attempting to extract drivers with dism..." -Verbose
	Start-Process dism -ArgumentList "/online /export-driver /destination:$($ExtractDriverDir)" -Wait
}
catch [System.Exception] {
	Write-Warning -Message "Failed to run dism"
}

if ($ConfigFile.Settings.Compression -eq "Zip")
{
	# Compress contents of ExtractDriverDir
	Write-Verbose -Message "[INFO] Creating DriverPackage.zip"
	Compress-Archive -Path (Join-Path -Path $ExtractDriverRootDir -ChildPath "DriverPackage") -DestinationPath (Join-Path -path $ExtractDriverRootDir -ChildPath "DriverPackage.zip") -CompressionLevel Fastest -Force
}

Write-Verbose -Message "[INFO] Writing ModelDetails.xml file..."
# Set XML Structure
$XmlWriter = New-Object System.XML.XmlTextWriter($ExtractDriverInfoXML, $Null)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"
$xmlWriter.WriteStartDocument()
$xmlWriter.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")

# Write Initial Header Comments
$XmlWriter.WriteComment('Created with the SCConfigMgr Driver Automation Tool')
$xmlWriter.WriteStartElement('Details')
$XmlWriter.WriteAttributeString('current', $true)

# Export Model Details 
$xmlWriter.WriteStartElement('ModelDetails')
$xmlWriter.WriteElementString('Make', $ExtractMake)
$xmlWriter.WriteElementString('Model', $CurrentModel.Model)
$xmlWriter.WriteElementString('SystemSKU', $ExtractSKU)
$xmlWriter.WriteElementString('OperatingSystem', $OSName)
$xmlWriter.WriteElementString('Architecture', $OSArchitecture)
$xmlWriter.WriteElementString('Platform', $Platform)
$xmlWriter.WriteEndElement()


# Save XML Document
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()


if ($SendCompletionEmail = "True")
{
	Write-Verbose -Message "[INFO] Sending email report..."
	
	$SendEmailFrom = "$($ConfigFile.Settings.EmailFromText) <$($ConfigFile.Settings.EmailFromAddress)>"
	$SendEmailTo = "$($ConfigFile.Settings.EmailToText) <$($ConfigFile.Settings.EmailToAddress)>"
	$SendEmailSubject = "$($ConfigFile.Settings.EmailSubject)"
	$SendEmailBody = "Driver capture completed. The latest model captured is: $($CurrentModel.Manufacturer) $($CurrentModel.Model) for $OSName $OSArchitecture - Please import the driver package using the XML file in this folder: $ExtractDriverInfoXML"
	$MailServer = "$($ConfigFile.Settings.EmailServer)"
	
	Write-Verbose -Message "[INFO] Sending email report to $SendEmailTo"
	
	Send-MailMessage -From $SendEmailFrom -To $SendEmailTo -Subject $SendEmailFrom -Body $SendEmailBody -Priority High -dno onSuccess, onFailure -SmtpServer $MailServer
}

Pause