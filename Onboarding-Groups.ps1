<#
Provided By: CapGemini
Last Updated By: BTL
Last Updated On: 17Nov2021
#>

#This function lets you build an array of specific list items you wish
Function MultipleSelectionBox ($inputarray,$label,$directions,$listboxtype, $icon) {
 
# Taken from Technet - http://technet.microsoft.com/en-us/library/ff730950.aspx
# This version has been updated to work with Powershell v3.0.
# Had to replace $x with $Script:x throughout the function to make it work. 
# This specifies the scope of the X variable.  Not sure why this is needed for v3.
# http://social.technet.microsoft.com/Forums/en-SG/winserverpowershell/thread/bc95fb6c-c583-47c3-94c1-f0d3abe1fafc
#
# Function has 3 inputs:
#     $inputarray = Array of values to be shown in the list box.
#     $prompt = The title of the list box
#     $listboxtype = system.windows.forms.selectionmode (None, One, MutiSimple, or MultiExtended)
 
$Script:x = @()
 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = $label
$objForm.Size = New-Object System.Drawing.Size(310,200) 
$objForm.StartPosition = "CenterScreen"
$objIcon = New-Object system.drawing.icon ($icon)
$objForm.Icon = $objIcon
 
$objForm.KeyPreview = $True
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
    }
    })
 
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})
 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(35,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
 
$OKButton.Add_Click(
   {
        foreach ($objItem in $objListbox.SelectedItems)
            {$Script:x += $objItem}
        $objForm.Close()
   })
 
$objForm.Controls.Add($OKButton)
 
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)
 
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(230,30) 
$objLabel.Text = $directions
#"Please make a selection from the list below:"
$objForm.Controls.Add($objLabel) 
 
$objListbox = New-Object System.Windows.Forms.Listbox 
$objListbox.Location = New-Object System.Drawing.Size(12,55) 
$objListbox.Size = New-Object System.Drawing.Size(270,55) 
 
$objListbox.SelectionMode = $listboxtype
 
$inputarray | ForEach-Object {[void] $objListbox.Items.Add($_)}
 
$objListbox.Height = 60
$objForm.Controls.Add($objListbox) 
$objForm.Topmost = $True
 
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
 
Return $Script:x
}

#creates a prompt for the new user being input
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Data Entry Form'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the information in the space below:'
$form.Controls.Add($label)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)
$form.Topmost = $true
$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $User = $textBox.Text
    $User
}

#$User= Read-Host "Enter Email Address"

MultipleSelectionBox
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