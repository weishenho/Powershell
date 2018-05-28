[Microsoft.Office.Interop.Outlook.Application] $outlook = New-Object -ComObject Outlook.Application
$entries = $outlook.Session.GetGlobalAddressList().AddressEntries
$count = $entries.count
foreach($entry in $entries){
	$count += 1
	if($count -eq 11){break}
	[console]::WriteLine("{0} : {1} : {2}", $entry.Name, $entry.GetExchangeUser().Department, $entry.GetExchangeUser().PrimarySmtpAddress)
}




$Array = @()
[Microsoft.office.Interop.Outlook.Application] $outlook = new-object -comobject Outlook.Application
$entries = $outlook.Session.getglobaladdresslist().addressentries
$count = $entries.count
$count =0
foreach($entry in $entries){
$x = $entry.GetExchangeUser()
$obj = new-object psobject -property @{
name = $entry.name
d=$x.department}
$Array += $obj
}
