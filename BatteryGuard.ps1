[CmdletBinding()]
Param(
	[Parameter(Mandatory = $false)]
	[switch]$Help,
	[switch]$Report,
	[int32]$Threshold
)

if ($Help) {
	Write-Output "Usage: BatteryGuard.ps1 [-Threshold] <threshold> [-Help] [-Report]"
	Write-Output "Arguments:"
	Write-Output "  -Threshold  Int value representing how low the battery can get before generating warning. Not needed if -Help or -Report are used."
	Write-Output "  -Report     Instead of monitoring battery level, will create a battery report HTML document in CWD, called 'Battery Report.html'."
	Write-Output "  -Help       This. Overrides -Report."
	exit
}

if ($Report) {
	
	# https://www.pcmag.com/how-to/how-to-check-your-laptops-battery-health-in-windows-10
	powercfg /batteryreport /output "Battery Report.html"
	exit

}

if (-not $Help -and -not $Report -and -not $Threshold) {
	Write-Error "Missing argument(s)"
	exit
}

# There's no need to do this if the charger's plugged in
# https://devblogs.microsoft.com/scripting/using-windows-powershell-to-determine-if-a-laptop-is-on-battery-power/
if ((Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnline) {
	Write-Output "Charger is plugged in, aborting."
	exit
}

# https://devblogs.microsoft.com/scripting/powertip-use-powershell-to-show-remaining-battery-time/
$BatteryPercentage = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining

while ($BatteryPercentage -ge $Threshold) {
	
	# Check for this each time, as the user might plug in the charger any time
	if ((Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnline) {
		Write-Output "Charger is plugged in, aborting."
		exit
	}

	Write-Progress -Activity "Battery" -Status "$BatteryPercentage%" -CurrentOperation "No action necessary as long as battery percentage remains at least at $Threshold%" -PercentComplete $BatteryPercentage
	$BatteryPercentage = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining
	Start-Sleep 10

}

# https://den.dev/blog/powershell-windows-notification/
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$RawXml = [xml]$Template.GetXml()
($RawXml.toast.visual.binding.text | Where-Object { $_.id -eq "1" }).AppendChild($RawXml.CreateTextNode("BATTERY WARNING")) | Out-Null
($RawXml.toast.visual.binding.text | Where-Object { $_.id -eq "2" }).AppendChild($RawXml.CreateTextNode("Battery at $BatteryPercentage%, laptop might die at any time!")) | Out-Null
$SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
$SerializedXml.LoadXml($RawXml.OuterXml)
$Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
$Toast.Tag = "PowerShell"
$Toast.Group = "PowerShell"
$Toast.ExpirationTime = [DateTimeOffset]::Now.AddDays(1)
$Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("BatteryGuard")
$Notifier.Show($Toast)
