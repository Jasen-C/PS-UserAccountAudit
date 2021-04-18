Import-Module ImportExcel
$XLSX = "C:\Users\jcrisp\OneDrive - Agri Beef Co\Documents\Scripts\Quarterly User Audits\UserAccountAudit.xlsx"


Remove-Item $XLSX -Force -ErrorAction SilentlyContinue
$Users = ""
$User = ""
$VPNUsers = ""
$lookup = ""

$ExcludeOU = @("CN=Monitoring Mailboxes,CN=Microsoft Exchange System Objects,DC=rebco,DC=local","CN=Users,DC=rebco,DC=local","OU=BOIIT SERVICE,OU=BOIIT,DC=rebco,DC=local","OU=Recipients,OU=UTILITY,DC=rebco,DC=local","OU=WAB CONTRACTORS,OU=WAB,DC=rebco,DC=local","OU=WAB IT USERS,OU=WAB-IT,DC=rebco,DC=local","OU=WAB OFFICE USERS,OU=WAB,DC=rebco,DC=local")
$lookupOU = $ExcludeOU | Group-Object -AsHashTable -AsString
$Users = get-aduser -Filter * -Properties * | Sort-Object | Select-Object Name,Created,SamAccountName,GivenName,Surname,Enabled,Description,extensionAttribute1,LastLogonDate,PasswordLastSet,@{n='OU';e={$_.distinguishedname -replace '^.+?,(CN|OU.+)','$1'}} |
? { ($_.ParentContainer -notlike '*Builtin*')}
$VPNUsers = Get-ADGroupMember -Identity SOPHOSVPN | Select-Object Name 
$lookup = $VPNUsers | Group-Object -AsHashTable -AsString -Property Name
foreach ($User in $Users){
    if ($lookupOU.ContainsKey($User.OU)){
        # Write-Host $User.name
        Continue
    }

    $searchTerm = $User.name
if ($lookup.ContainsKey($searchTerm)){
    # "Found $searchTerm"
    $User | Add-Member -MemberType NoteProperty -Name 'VPN Access' -Value $User.Name
}else{
    # "$searchTerm not found"
    $User | Add-Member -MemberType NoteProperty -Name 'VPN Access' -Value ""
    } 
    $User | Add-Member -MemberType NoteProperty -Name 'Location' -Value $User.extensionAttribute1
    $User | Add-Member -MemberType NoteProperty -Name 'User' -Value $User.name
    $User | Add-Member -MemberType NoteProperty -Name 'First Name' -Value $User.GivenName
    # $User | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $User.SamAccountName
    $User | Add-Member -MemberType NoteProperty -Name 'Last Name' -Value $User.Surname
    $User | Add-Member -MemberType NoteProperty -Name 'Shared Account' -Value ""
    $User | Add-Member -MemberType NoteProperty -Name 'Status' -Value ""
    $User | Add-Member -MemberType NoteProperty -Name 'Notes' -Value ""
    $User | Add-Member -MemberType NoteProperty -Name 'Continue VPN Access?' -Value "" 
    # if ($User.LastLogondate -le $Date){
    #     $User.LastLogondate = "Over 90 days"
    # }else{$User.LastLogondate = ""}
    $User = $User | Select-Object "Shared Account","Status",SamAccountName,"Notes","First Name", "Last Name",Enabled,"Location",Description,"OU",Created,LastLogonDate,PasswordLastSet,"VPN Access"
    # $User | export-csv test.csv -Append
    $User | export-excel $XLSX -WorkSheetname 'All User Accounts' -Append
}

