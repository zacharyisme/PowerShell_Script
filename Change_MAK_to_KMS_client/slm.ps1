# Strict-Mode - remember this, it's important!
Set-StrictMode -Version Latest

# variables
#$cimsession = New-CimSession
#$pk_type = get-ciminstance -class SoftwareLicensingProduct-CimSession $cimsession|where {$_.name -match 'windows' -AND $_.licensefamily}|where LicenseStatus -ne 0
$pk_type = get-WmiObject SoftwareLicensingProduct | where-object {$_.licensestatus -eq 1}
if ( $pk_type.Description.Contains("MAK") -and $pk_type.Description.Contains("Windows(R) 7")) {
  # We have a MAK key and Windows 7, let's change the product
  # key and activate
  Invoke-Expression "cscript //t:20 c:\windows\system32\slmgr.vbs /ipk <KMS_KEY>"
  Invoke-Expression "cscript //t:20 c:\windows\system32\slmgr.vbs /ato"
  #write-output "MAK"
}
#Else {
#  write-output "Other"
#}