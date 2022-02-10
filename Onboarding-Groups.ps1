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

If($connectionverify -eq '0'){powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Connections Failed Stopping')}"
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



#Select the User type
while($employeetypeverify -eq '0'){
	$options01 = "GTIL employees/secondees","GTIL consultants/contractors","GTIL consultants/contractors with laptops"
	$employeetype = MultipleSelectionBox -listboxtype one -inputarray $options01 -label 'Groups Onboarding' -directions 'Verify selection in ticket and check hardware required' -icon "C:\Windows\SystemApps\Microsoft.Windows.SecHealthUI_cw5n1h2txyewy\Assets\Account.theme-light.ico"
	$employeetypeverify = user-prompt -Title "Verify Employee Type" -Message "You selected $employeetype is this correct?"
}
if ($employeetype -eq 'GTIL employees/secondees') 
{
	#Get GTIL-Users security group and add employee	
	$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Users”}
	$member = Get-MsolUser -UserPrincipalName $User
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL-Users Security Group')}"
	Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId

	#Get GTIL_Employees_Global Distribution Group and add employee
	$Group="GTIL_Employees_Global@gti.gt.com"
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL_Employees_Global Distribution Group')}"
	Add-DistributionGroupMember -Identity $Group  -Member $User

	#Add the IDHubInclude extension
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding IDHubInclude extension to $User')}"
	Set-AzureADUserExtension -ObjectID $User -ExtensionName extension_7ad0543d182445dcbce5d98a226ce6e2_gtIDHubFilterCloud -ExtensionValue 'IDHub=Include'
	Write-Host $User has been included to ID HUB -ForegroundColor Red -BackgroundColor White
	
	#Select the employees location
	while($employeelocationverify -eq '0'){
		$options02 = "London","Chicago","Downers Grove","No Location Groups"
		$employeelocation = MultipleSelectionBox -listboxtype one -inputarray $options01 -label 'Employee Location' -directions 'Verify location in ticket' -icon "C:\Windows\SystemApps\Microsoft.Windows.SecHealthUI_cw5n1h2txyewy\Assets\Account.theme-light.ico"
		$employeelocationverify = user-prompt -Title "Verify Employee location" -Message "You selected $employeelocation is this correct?"
	}
	
	#Add employee to location specific groups
	if ($employeelocation -eq 'London') 
	{ $Group="GTIL_Employees_London@gti.gt.com"
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL_Employees_London Distribution Group')}"
	Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
	} 
	if ($employeelocation -eq 'Chicago') 
	{ 
	 $Group="GTIL_Employees_Chicago@gti.gt.com"
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL_Employees_Chicago Distribution Group')}"
	Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
	} 
	if ($employeelocation -eq 'Downers Grove') 
	{ $Group="GTIL_Employees_DownersGrove@gti.gt.com"
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL_Employees_DownersGrove Distribution Group')}"
	Add-DistributionGroupMember -Identity $Group  -Member $User -ErrorAction Stop
	}
	if ($employeelocation -eq 'No Location Groups') 
	{ 
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('$User is not being added to any location specific Distribution Groups')}"
	}  
}


if ($employeetype -eq 'GTIL consultants/contractors')
{
	#Get GTIL-Support security group and add Consultant/Contractor		
	$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support”}
	$member = Get-MsolUser -UserPrincipalName $User
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL-Support Security Group')}"
	Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId
}

if ($employeetype -eq 'GTIL consultants/contractors with laptops')
{
	#Get GTIL-Support security group and add Consultant/Contractor	
	$securityGroup = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support”}
	$member = Get-MsolUser -UserPrincipalName $User
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL-Support Security Group')}"
	Add-MsolGroupMember -GroupObjectId $securityGroup.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member.ObjectId
	
	#Get GTIL-Support with laptops security group and add Consultant/Contractor
	$securityGroup1 = Get-MsolGroup -GroupType “Security” | Where-Object {$_.DisplayName -eq “GTIL-Support with laptops”}
	$member1 = Get-MsolUser -UserPrincipalName $User
	powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Adding $User to GTIL-Support with laptops Security Group')}"
	Add-MsolGroupMember -GroupObjectId $securityGroup1.ObjectId -GroupMemberType “User” -GroupMemberObjectId $member1.ObjectId
}