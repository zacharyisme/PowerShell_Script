# New-PasswordReminder.ps1

###################################
# Get the max Password age from AD 
###################################
function Get-maxPwdAge{
   $root = [ADSI]"LDAP://dc1.company.com.tw"
   $filter = "(&(objectcategory=domainDNS)(distinguishedName=DC=company,DC=com,DC=tw))"
   $ds = New-Object system.DirectoryServices.DirectorySearcher($root,$filter)
   $dc = $ds.findone()
   [int64]$maxpwdage = [System.Math]::Abs( $dc.properties.item("maxPwdAge")[0])
   $maxpwdage/864000000000
}

###################################
# Function to send HTML email to each user
###################################

function send_email ($days_remaining, $email, $name ) 
{
 $today = Get-Date
 $today = $today.ToString("dddd (yyy-MMMM-dd)")
 $date_expire = [DateTime]::Now.AddDays($days_remaining);
 $date_expire = $date_expire.ToString("dddd (yyy-MMMM-dd)")
 $emailSmtpServer = "192.168.0.1"
 $emailSmtpServerPort = "25"
 $SmtpClient = new-object system.net.mail.smtpClient($emailSmtpServer, $emailSmtpServerPort) 
 $mailmessage = New-Object system.net.mail.mailmessage 
 #$SmtpClient.Host = "" 
 $mailmessage.from = "IT Service <itservice@company.com.tw>" 
 $mailmessage.To.add($email)
 $mailmessage.Subject = "$name, Your password will expire soon."
 $mailmessage.IsBodyHtml = $true

 $mailmessage.Body = @"
<h5><font face=Arial>Dear $name, </font></h5>
<h5><font face=Arial>Your password will expire in <font color=red><strong>$days_remaining</strong></font> days
 on <strong>$date_expire</strong><br /><br />
Your domain password is required for Computer Login, remote VPN, and Email Access.<br /><br />
To change your password, <br /><br />
1. If you are login to domain PC, press CTRL-ALT-DEL and choose Change Password.<br /><br />
2. or you can logon to web mail https://mail.company.com.tw through IE Browser to change your password.<br /><br />
For your password to be valid it must be 8 or more characters long and<br />
contain a mix of THREE of the following FOUR properties:<br /><br />
    uppercase letters (A-Z)<br />
    lowercase letters (a-z)<br />
    numbers (0-9)<br />
    symbols (!$%^&*)<br /><br />
If you have any questions, please contact the IT Service on itservice@company.com.tw <br /><br />
 Generated on : $today<br /><br />
_____________ <br />
<br /></font></h5>
"@

 $smtpclient.Send($mailmessage) 
}

###################################
# Search for Non-disabled AD users that have a Password Expiry.
###################################

$strFilter = "(&(objectCategory=User)(logonCount>=0)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(!(userAccountControl:1.2.840.113556.1.4.803:=65536)))"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000	
$objSearcher.Filter = $strFilter
$colResults = $objSearcher.FindAll();

# how many days before PW expiry do we start sending reminder emails?
$max_alert = 5


# Get the maximum password lifetime
$max_pwd_life=Get-maxPwdAge

$userlist = @()
foreach ($objResult in $colResults)
   {$objItem = $objResult.Properties; 
   if ( $objItem.mail.gettype.IsInstance -eq $True) 
      {      
         $user_name = $objItem.name
         $user_email = $objItem.email
         #Transform the DateTime readable format
         $user_logon = [datetime]::FromFileTime($objItem.lastlogon[0])
         $result = $objItem.pwdlastset 
         $user_pwd_last_set = [datetime]::FromFileTime($result[0])

         #calculate the difference in Day from last time a password was set
         $diff_date = [INT]([DateTime]::Now - $user_pwd_last_set).TotalDays;

   $Subtracted = $max_pwd_life - $diff_date
         if (($Subtracted) -le $max_alert) {
            $selected_user = New-Object psobject
            #$selected_user | Add-Member NoteProperty -Name "Name" -Value $objItem.name[0]
            $selected_user | Add-Member NoteProperty -Name "Name" -Value $objItem.Item("displayname")
            $selected_user | Add-Member NoteProperty -Name "Email" -Value $objItem.mail[0]
            $selected_user | Add-Member NoteProperty -Name "LastLogon" -Value $user_logon
            $selected_user | Add-Member NoteProperty -Name "LastPwdSet" -Value $user_pwd_last_set
            $selected_user | Add-Member NoteProperty -Name "RemainingDays" -Value ($Subtracted)
            $userlist+=$selected_user
         }
      }
   }

###################################
# Send email to each user
###################################
   foreach ($userItem in $userlist )
   {
    if ($userItem.RemainingDays -ge 0) {
      #send_email $userItem.RemainingDays $userItem.Email $userItem.Name
      send_email $userItem.RemainingDays abc@company.com.tw $userItem.Name
	  #$userItem.Name | Out-File d:\userlist.txt
       }
   }

# END

####################################
#Write-Host "UserList : " $userlist
#Write-Host "Maxpwage:" $max_pwd_life
$userlist | Out-File d:\userlist.txt