<#
Created By: BTL - Kristopher Roy
Created On: 10Feb22
Last Updated On: 10Feb22
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

#This is a Function that allows a tailord input for recording
Function InputBox($header,$text)
{
    #creates a prompt for the new user being input
    $Input = @()
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $header
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
    $label.Text = $text
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
    Return $textBox.Text
    }
}

#A Basic Dynamic 2 option User Prompt Function
function user-prompt
{
	#If Select1 and 2 are not defined deafault is Yes,No
	param(
		#Title Defines Header for Prompt
		[Parameter(Mandatory=$false)]
		[string] $Title,
		#Message Defines Body of Prompt
		[Parameter(Mandatory=$False)]
		[string]$Message,
		#Select1 Defines First of two options
		[Parameter(Mandatory=$False)]
		[string]$Select1,
		#Select2 Define Second of two options
		[Parameter(Mandatory=$False)]
		[string]$Select2,
		#String to Display what first option means in a tooltip
		[Parameter(Mandatory=$False)]
		[string]$Selection1ToolTip,
		#String to Display what second option means in a tooltip
		[Parameter(Mandatory=$False)]
		[string]$Selection2ToolTip
    )

    If($Title -eq $Null -or $Title -eq ""){$title = "Selection"}
    If($Select1 -eq $Null -or $Select1 -eq ""){$Select1 = " No "}
	$selection1 = New-Object System.Management.Automation.Host.ChoiceDescription "&$Select1", `
	$selection1ToolTip
    If($Select2 -eq $Null -or $Select2 -eq ""){$Select2 = " Yes "}
	$selection2 = New-Object System.Management.Automation.Host.ChoiceDescription "&$Select2", `
	$selection2ToolTip
    If($message -eq $Null -or $Message -eq ""){$message = "Basic $select1,$select2 Options"}
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($selection1, $selection2)
    $choice = $host.ui.PromptForChoice($title, $message, $options, 0)    
    $choice = [int]$choice
    Return $choice
    #Returns selection1 = 0, and selection2 = 1

<#
	example 1 - No switches defined:
	user-prompt
	Result = A Yes,No prompt with Title "Selection", Body "Basic Yes,No Options"
		
	example 2 - Selections Defined
	user-prompt -select1 "blue" -select2 "green"
	Result = A blue,green prompt with Title "Selection" Body "Basic Blue,Green Options"
		
	Output will always be 0,1
#>
}

#Begin Script
#Authenticate and add Azure/O365 modules
$credential= Get-Credential
Connect-AzureAD -Credential $credential
Connect-MsolService -Credential $credential
Connect-ExchangeOnline -ShowProgress $true

#Verify Connections are good
$tenantDomain = ((Get-AzureADTenantDetail).VerifiedDomains|where{$_.Name -eq 'gti.gt.com'}).name
$MsolDomain = (Get-MsolDomain|where{$_.Name -eq 'gti.gt.com'}).name
$EXOdomain = (Get-EXOMailbox -ResultSize 1|select PrimarySmtpAddress).PrimarySmtpAddress.split('@')[1]

$connectionverify = user-prompt -Title "Verify Connections" -Message "AZDomain = $tenantDomain, MSOLDomain = $MsolDomain, EXODomain = $EXODomain is this correct"

If($connectionverify -eq 0){powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Connections Failed Stopping')}"
	Exit}

#Gather and verify Ticket Number
while($ticketverify -eq '0'){
    $ticket = InputBox -header "Ticket Number" -text "Input the related Ticket Number"

    $ticketverify = user-prompt -Title "Verify Ticket" -Message "Is $ticket correct?"
}

#Gather and Verify New User Email
while($userverify -eq '0'){
    $user = InputBox -header "New User" -text "Input the New User Email from $ticket"

    $userverify = user-prompt -Title "Verify User" -Message "Is $user correct?"
}



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