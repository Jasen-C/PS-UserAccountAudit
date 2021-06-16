# Author Jasen C
# Date 6/16/2021
# Description: Reusable script used to automate data collection of user accounts for quarterly user account audits. Outputs data to an excel sheet and emails
# to admins

# Requires the importexcel module to create and edit excel files
# Install-Module -Name ImportExcel
Import-Module ImportExcel

# Define location to safe excel file
$XLSX = "C:\Users\Me\\Documents\Scripts\Quarterly User Audits\AllUserAccountAudit.xlsx"

#Delete the file to start with a fresh copy
Remove-Item $XLSX -Force -ErrorAction SilentlyContinue

# initialize variables
$Users = ""
$User = ""
$VPNUsers = ""
$lookup = ""
$VPNGroup = "VPN-Group-name"

#region UserAudit

# Define OUs to exclude from search, other business units, service account locations, mailboxes
$ExcludeOU = @("CN=Monitoring Mailboxes,CN=Microsoft Exchange System Objects,DC=company,DC=local","CN=Users,DC=company,DC=local","OU=Recipients,OU=UTILITY,DC=company,DC=local")


$lookupOU = $ExcludeOU | Group-Object -AsHashTable -AsString

# Properties can probably be scaled back, currently pulling all users properties
$Users = get-aduser -Filter * -Properties * | Sort-Object | Select-Object Name,Description,extensionAttribute1,LastLogonDate,PasswordLastSet,@{n='OU';e={$_.distinguishedname -replace '^.+?,(CN|OU.+)','$1'}} |
? { ($_.ParentContainer -notlike '*Builtin*')}

#Get users authorized for VPN access
$VPNUsers = Get-ADGroupMember -Identity $VPNGroup | Select-Object Name

# Add all the VPN Users into a hash table to perform lookups against
$lookup = $VPNUsers | Group-Object -AsHashTable -AsString -Property Name

foreach ($User in $Users){
    # if the user's OU is in our Excluded OU list, skip the account and continue to next account
    if ($lookupOU.ContainsKey($User.OU)){
        #Write-Host $User.name #used to debug and see what accounts are being excluded
        Continue
    }

    #Set our searchTerm to the user's name to check for VPN access
    $searchTerm = $User.name
    
if ($lookup.ContainsKey($searchTerm)){
    
    # Create an obeject attribute for VPN Access and set it to the user's name
    $User | Add-Member -MemberType NoteProperty -Name 'VPN Access' -Value $User.Name
}else{
    
    # Create an obeject attribute for VPN Access and set it to null
    $User | Add-Member -MemberType NoteProperty -Name 'VPN Access' -Value ""
    } 
    #Set additional Object Attributes with the names we will use in our excel file
    $User | Add-Member -MemberType NoteProperty -Name 'Location' -Value $User.extensionAttribute1
    $User | Add-Member -MemberType NoteProperty -Name 'Current User' -Value $User.name
    $User | Add-Member -MemberType NoteProperty -Name 'Action to be taken' -Value "" 
    $User | Add-Member -MemberType NoteProperty -Name 'Reason for Action' -Value "" 
    $User | Add-Member -MemberType NoteProperty -Name 'Continue VPN Access?' -Value ""

    if ($User.LastLogondate -le $Date){
        #Flag user accounts not logged in in the last 90 days, if under 90 days set to null
        $User.LastLogondate = "Over 90 days -" + $User.LastLogondate
    }else{$User.LastLogondate = ""}
    # Reorder User fields and set user object to selected items
    $User | Select-Object "Current User","Action to be taken","Reason for Action","Location",Description,"VPN Access","Continue VPN Access?","OU",LastLogonDate,PasswordLastSet |
    export-excel $XLSX -WorkSheetname 'User Accounts' -Append
}
#endregion


#Region Admin Account Audit

write-host "Members of Administrators"
$Administrators = Get-ADGroup Administrators -Properties members |select  -expand members | sort
foreach ($i in $Administrators){
try {
        $Account = ""
        $Account = Get-ADObject $i -Properties Samaccountname,objectclass,description -ErrorAction 'SilentlyContinue'| Select Samaccountname,objectclass,description
        $Account | Add-Member -MemberType NoteProperty -Name 'Remove' -Value ""
        $Account | Add-Member -MemberType NoteProperty -Name 'Reason' -Value ""
        $Account = $Account | Select-Object Samaccountname,objectclass,description,"Remove","Reason"
        $Account | export-excel $XLSX -WorkSheetname 'Administrators Group' -Append
}
catch{

}
}

write-host "Members of Domain Administrators"
$Administrators = Get-ADGroup "Domain Admins" -Properties members |select  -expand members | sort
$Administrators += Get-ADGroup 'ADAdmin' -Properties members |select  -expand members | sort
foreach ($i in $Administrators){
try {
        $Account = ""
        $Account = Get-ADObject $i -Properties Samaccountname,objectclass,description -ErrorAction 'SilentlyContinue'| Select Samaccountname,objectclass,description
        $Account | Add-Member -MemberType NoteProperty -Name 'Remove' -Value ""
        $Account | Add-Member -MemberType NoteProperty -Name 'Reason' -Value ""
        $Account = $Account | Select-Object Samaccountname,objectclass,description,"Remove","Reason"
        $Account | export-excel $XLSX -WorkSheetname 'Domain Admins Group' -Append
}
catch{

}
}

write-host "Members of Enterprise Admins"
$Administrators = Get-ADGroup "Enterprise Admins" -Properties members |select  -expand members | sort
$Administrators += Get-ADGroup 'ADAdmin' -Properties members |select  -expand members | sort
foreach ($i in $Administrators){
try {
    $Account = ""
$Account = Get-ADObject $i -Properties Samaccountname,objectclass,description -ErrorAction 'SilentlyContinue'| Select Samaccountname,objectclass,description
$Account | Add-Member -MemberType NoteProperty -Name 'Remove' -Value ""
$Account | Add-Member -MemberType NoteProperty -Name 'Reason' -Value ""
$Account = $Account | Select-Object Samaccountname,objectclass,description,"Remove","Reason"
$Account | export-excel $XLSX -WorkSheetname 'Enterprise Admins Group' -Append
}
catch{

}
}
#endregion



#### Send Email ####
$smtphost = "dns.emailserver.local" 
$from = "sendingAccount@company.local"
$email1 = "User1@company.local"
$email2 = "User2@company.local"

$subject = "Quarterly User Account Audit" 
$attachment1 = $XLSX
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage 
$msg.To.Add($email1)
$msg.To.Add($email2)
$msg.Attachments.Add($attachment1)
$msg.from = $from
$msg.subject = $subject
$smtp.send($msg) 
