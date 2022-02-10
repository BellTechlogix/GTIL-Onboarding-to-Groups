<#
Provided By: CapGemini
Last Updated By: BTL
Last Updated On: 17Nov2021
#>


$User= Read-Host "Enter Email Address" 
Write-Host "Is $User GTIL employees/secondees OR GTIL consultants/contractors? (
1 for  GTIL employees/secondees or
2 for GTIL consultants/contractors
3 for GTIL consultants/contractors with laptops)  " -ForegroundColor White -BackgroundColor Blue
$resp1 = Read-host 'Enter Choice' 
if ($resp1 -eq 1) 
{
$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Users”}
$member = Get-MsolUser -UserPrincipalName $User
Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId
$Group="GTIL_Employees_Global@gti.gt.com"
#$Group1="GTIL-Users"
Add-DistributionGroupMember -Identity $Group  -Member $User
#Add-UnifiedGroupLinks -Identity $Group1 -LinkType Members -Links $User -ErrorAction Stop
#Write-Host $User has been added to $Group and $Group1 as $User is an GTIL employees/secondees -ForegroundColor White -BackgroundColor Blue
Set-AzureADUserExtension -ObjectID $User -ExtensionName extension_7ad0543d182445dcbce5d98a226ce6e2_gtIDHubFilterCloud -ExtensionValue 'IDHub=Include'
Write-Host $User has been included to ID HUB -ForegroundColor Red -BackgroundColor White
Write-Host "Enter User's Location? (
1 for  London or 
2 for Chicago or
3 for Downers Grove
4 for No Groups requires to be Added )" -ForegroundColor White -BackgroundColor Blue
$resp = Read-host 'Enter Choice' 
if ($resp -eq 1) 
{ $Group="GTIL_Employees_London@gti.gt.com"
Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
Write-Host $User has been added to $Group -ForegroundColor White -BackgroundColor Blue
} 
if ($resp -eq 2) 
{ 
 $Group="GTIL_Employees_Chicago@gti.gt.com"
Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
Write-Host $User has been added to $Group -ForegroundColor White -BackgroundColor Blue 
} 
if ($resp -eq 3) 
{ $Group="GTIL_Employees_DownersGrove@gti.gt.com"
Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
Write-Host $User has been added to $Group -ForegroundColor White -BackgroundColor Blue
}
if ($resp -eq 4) 
{ 
Write-Host $User has not been added to any Location Group -ForegroundColor White -BackgroundColor Blue
}  }
if ($resp1 -eq 2)
{$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support”}
$member = Get-MsolUser -UserPrincipalName $User
Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId
#$Group="GTIL-Support"
#Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $User
#Write-Host $User has been added to $Group -ForegroundColor White -BackgroundColor Blue
}
if ($resp1 -eq 3)
{$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support”}
$member = Get-MsolUser -UserPrincipalName $User
Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId
$securityGroup1 = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support with laptops”}
$member1 = Get-MsolUser -UserPrincipalName $User
Add-MsolGroupMember -GroupObjectId $securityGroup1.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member1.ObjectId

}



#GTINetTest.Employee1@gti.gt.com--London
#GTINetTest.Employee2@gti.gt.com--Chicago
#GTINetTest.Employee3@gti.gt.com--Downers Grove
#GTINetTest.Contractor@gti.gt.com--Contractor