### Loads ###

# Loads the necessary components for creating the forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")


### Arrays ###

# Lists departments for Department dropdown in main form
$DepartmentOptions = @(
    # List of Available Department names
)

# Lists user types for User Type dropdown in the main form
$UserTypes = @(
    # Array of available user types
)

# The array that stores all information entered by the form
$currentUser = @{
    FirstName = ''
    LastName = ''
    Email = ''
    Password = ''
    Department = ''
    Title = ''
    Manager = ''
    Alias = [System.Collections.ArrayList]@()
    UserType = ''
    Licenses = [System.Collections.ArrayList]@()
    MirrorGroup = ''
    PhoneNumber = ''
}

# Information for each possible license choice
$availableLicenseChoice = [Ordered]@{
    # Hash table of License names to SKU values
}


### Functions ###

# This function gets you connected to AzureAD, ExchangeOnlineManagement, and MicrosoftTeams
function GetConnected {
    # Checks to make sure the modules are available, installs if not
    if (Get-Module -ListAvailable -Name AzureAD) {
        Write-Host "Azure AD Module exists"
    } 
    else {
        Write-Host "Module does not exist"
        Install-Module -Name AzureAD
        Import-Module -Name AzureAD
    }
    
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "Exchange Online Module exists"
    } 
    else {
        Write-Host "Module does not exist"
        Install-Module -Name ExchangeOnlineManagement
        Import-Module -Name ExchangeOnlineManagement
    }
    
    if (Get-Module -ListAvailable -Name MicrosoftTeams) {
        Write-Host "Microsoft Teams Module exists"
    } 
    else {
        Write-Host "Module does not exist"
        Install-Module -Name MicrosoftTeams
        Import-Module -Name MicrosoftTeams
    }
    
    # Checks to see if you are connected to the modules.  Connects if not.
    try { 
        $var = Get-AzureADTenantDetail 
        if ($var) {
            Write-Host "You are connected to Azure AD"
        }
    } 
    
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { 
        Write-Host "You're not connected to Azure AD"
        Connect-AzureAD
    }

    $EOConnected = $false
    $psSessions = Get-PSSession | Select-Object -Property State, Name
    if (((@($psSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -eq $true) {
        Write-Host "You are connected to Exchange Online"
        $EOConnected = $true
    }

    else {
        Write-Host "You're not connected to Exchange Online"
        Connect-ExchangeOnline
        if (((@($psSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0) -eq $true) {
            $EOConnected = $true
        }
    }

    $MTConnected = $false
    try {    
        $TeamsTest = Get-CsTenant
        if ($TeamsTest) {
            Write-Host "You are connected to Microsoft Teams"
            $MTConnected = $true
        }
    }
    catch {
        Write-Host "You're not connected to Microsoft Teams"
        Connect-MicrosoftTeams
        $TeamsCheckAgain = Get-CsTenant
        if ($TeamsCheckAgain) {
            $MTConnected = $true
        }
    }

    $azureADConnected = Get-AzureADTenantDetail

    # Returns true if all three are currently connected
    if ($azureADConnected -AND $EOConnected -AND $MTConnected) {
        return $true
    }
}


# Displays available licenses to verify before running the script.
function checkCurrentLicenses {
    Write-Host "`nCurrent Available Licenses:" -ForegroundColor Magenta

    # Creates table of available licenses based on license array above
    $availableLicenseChoice | Format-Table Name, @{L="Available";E={$currentLicense = Get-AzureADSubscribedSku -ObjectID $_.Value; ($currentLicense | Select-Object -ExpandProperty PrepaidUnits).Enabled - $currentLicense.ConsumedUnits}} -AutoSize

    # Prompts user to continue if appropriate
    while ($true) {
        $continueLicense = Read-Host -Prompt "Given the available licenses, do you want to continue? (y / n)"
        
        # Moves on if 'y' is entered
        if ($continueLicense -eq 'y') {
            break
        }

        # Explains next steps and terminates script if 'n' is entered
        elseif ($continueLicense -eq 'n') {
            Write-Host "`n**Please purchase the appropriate licenses and then run the script again.**`n"
            exit
        }

        # Checks for valid entry
        else {
            Write-Host "`nThis is not a valid option.  Please select y or n.`n"
        }
    }
}


###################################
### Form Creation and Operation ###
###################################

function CreateForm{
    # Main Form Creation
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.Text = "Onboarding Form"
    $MainForm.Size = New-Object System.Drawing.Size(400,600)
    $MainForm.KeyPreview = $True
    $MainForm.FormBorderStyle = "1"
    $MainForm.MaximizeBox = $false
    $MainForm.StartPosition = "CenterScreen"
    $MainForm.Topmost = $true

    ### Create Tab Control ###

    $TabControl = New-Object System.Windows.Forms.TabControl
    $TabControl.Size = New-Object System.Drawing.Size(400,520)
    $TabControl.TabIndex = 0
    $MainForm.Controls.Add($TabControl)

    ### Add Tabs ###
    # User Information Tab
    $UserInfoTab = New-Object System.Windows.Forms.TabPage
    $UserInfoTab.Text = "User Info"
    $UserInfoTab.TabIndex = 0
    $TabControl.Controls.Add($UserInfoTab)

    # Licenses Tab
    $LicensesTab = New-Object System.Windows.Forms.TabPage
    $LicensesTab.Text = "Licenses"
    $LicensesTab.TabIndex = 1
    $TabControl.Controls.Add($LicensesTab)

    # Optional Tab
    $OptionalTab = New-Object System.Windows.Forms.TabPage
    $OptionalTab.Text = "Optional"
    $OptionalTab.TabIndex = 2
    $TabControl.Controls.Add($OptionalTab)


    ### Add Buttons ###

    # Submit Button
    $Submit = New-Object System.Windows.Forms.Button
    $Submit.Size = New-Object System.Drawing.Size(75,25)
    $Submit.Location = New-Object System.Drawing.Point(200,530)
    $Submit.Text = "Submit"
    $MainForm.Controls.Add($Submit)
    $Submit.add_click({
        $confirm, $message = VerifyMandatoryFields

        if ($confirm -eq $True) {
            $ConfirmResult = ConfirmFormResult
            
            if ($ConfirmResult -eq [System.Windows.Forms.DialogResult]::OK) {
                $MainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
        }

        else {
            $MissingInfo = New-Object System.Windows.Forms.Label
            $MissingInfo.Size = New-Object System.Drawing.Size(200,40)
            $MissingInfo.Location = New-Object System.Drawing.Point(10,528)
            $MissingInfo.Text = "Please Enter the Mandatory Information"
            $MissingInfo.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0)
            $MainForm.Controls.Add($MissingInfo)

            [System.Windows.Forms.MessageBox]::Show($message, "Missing Parameters")
        }
    })

    # Cancel Button
    $Cancel = New-Object System.Windows.Forms.Button
    $Cancel.Size = New-Object System.Drawing.Size(75,25)
    $Cancel.Location = New-Object System.Drawing.Point(300,530)
    $Cancel.Text = "Cancel"
    $MainForm.Controls.Add($Cancel)
    $Cancel.add_click({
        $MainForm.Close()
    })

    
    ### Create Form Items for User Info Tab ###
    ###########################################

    ### Add Lables ###

    # First Name Label
    $FN_Label = New-Object System.Windows.Forms.Label
    $FN_Label.Size = New-Object System.Drawing.Size(70,20)
    $FN_Label.Location = New-Object System.Drawing.Point(5,13)
    $FN_Label.Text = "First Name*"
    $UserInfoTab.Controls.Add($FN_Label)

    # Last Name Label
    $LN_Label = New-Object System.Windows.Forms.Label
    $LN_Label.Size = New-Object System.Drawing.Size(70,20)
    $LN_Label.Location = New-Object System.Drawing.Point(5,43)
    $LN_Label.Text = "Last Name*"
    $UserInfoTab.Controls.Add($LN_Label)

    # Department Label
    $Dept_Label = New-Object System.Windows.Forms.Label
    $Dept_Label.Size = New-Object System.Drawing.Size(90,20)
    $Dept_Label.Location = New-Object System.Drawing.Point(5,73)
    $Dept_Label.Text = "Department*"
    $UserInfoTab.Controls.Add($Dept_Label)

    # Employee Type Label
    $Dept_Label = New-Object System.Windows.Forms.Label
    $Dept_Label.Size = New-Object System.Drawing.Size(95,20)
    $Dept_Label.Location = New-Object System.Drawing.Point(5,103)
    $Dept_Label.Text = "Employee Type*"
    $UserInfoTab.Controls.Add($Dept_Label)

    # Title Label
    $Title_Label = New-Object System.Windows.Forms.Label
    $Title_Label.Size = New-Object System.Drawing.Size(90,20)
    $Title_Label.Location = New-Object System.Drawing.Point(5,133)
    $Title_Label.Text = "Title*"
    $UserInfoTab.Controls.Add($Title_Label)

    # Manager Email Label
    $ME_Label = New-Object System.Windows.Forms.Label
    $ME_Label.Size = New-Object System.Drawing.Size(95,20)
    $ME_Label.Location = New-Object System.Drawing.Point(5,163)
    $ME_Label.Text = "Manager Email*"
    $UserInfoTab.Controls.Add($ME_Label)

    # Mirror Groups Email Label
    $PW_Label = New-Object System.Windows.Forms.Label
    $PW_Label.Size = New-Object System.Drawing.Size(95,20)
    $PW_Label.Location = New-Object System.Drawing.Point(5,193)
    $PW_Label.Text = "Mirror Groups On"
    $UserInfoTab.Controls.Add($PW_Label)

    # Password Label
    $PW_Label = New-Object System.Windows.Forms.Label
    $PW_Label.Size = New-Object System.Drawing.Size(95,20)
    $PW_Label.Location = New-Object System.Drawing.Point(5,223)
    $PW_Label.Text = "Password*"
    $UserInfoTab.Controls.Add($PW_Label)


    ### Add Text Forms ###

    # First Name Text Box
    $FN_TB = New-Object System.Windows.Forms.TextBox
    $FN_TB.Location = New-Object System.Drawing.Point(105,10)
    $FN_TB.Size = New-Object System.Drawing.Size(200,20)
    $UserInfoTab.Controls.Add($FN_TB)

    # Last Name Text Box
    $LN_TB = New-Object System.Windows.Forms.TextBox
    $LN_TB.Location = New-Object System.Drawing.Point(105,40)
    $LN_TB.Size = New-Object System.Drawing.Size(200,20)
    $UserInfoTab.Controls.Add($LN_TB)

    # Title Text Box
    $Title_TB = New-Object System.Windows.Forms.TextBox
    $Title_TB.Location = New-Object System.Drawing.Point(105,130)
    $Title_TB.Size = New-Object System.Drawing.Size(200,20)
    $UserInfoTab.Controls.Add($Title_TB)

    # Manager Email Text Box
    $ME_TB = New-Object System.Windows.Forms.TextBox
    $ME_TB.Location = New-Object System.Drawing.Point(105,160)
    $ME_TB.Size = New-Object System.Drawing.Size(200,20)
    $UserInfoTab.Controls.Add($ME_TB)

    # Mirror Groups Email Text Box
    $MG_TB = New-Object System.Windows.Forms.TextBox
    $MG_TB.Location = New-Object System.Drawing.Point(105,190)
    $MG_TB.Size = New-Object System.Drawing.Size(200,20)
    $UserInfoTab.Controls.Add($MG_TB)


    ### Add Masked Text Box ###

    # Password Masked Text Box
    $PW_MTB = New-Object Windows.Forms.MaskedTextBox
    $PW_MTB.PasswordChar = '*'
    $PW_MTB.Size = New-Object System.Drawing.Size(200,20)
    $PW_MTB.Location = New-Object System.Drawing.Point(105,220)
    $UserInfoTab.Controls.Add($PW_MTB)


    ### Add Drop Down Menus ###

    # Department Dropdown
    $Dep_DD = New-Object System.Windows.Forms.ComboBox
    $Dep_DD.Size = New-Object System.Drawing.Size(200,20)
    $Dep_DD.Location = New-Object System.Drawing.Point(105,70)
    $Dep_DD.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    # Add options to Dep_DD
    ForEach ($Department in $DepartmentOptions) {
        $Dep_DD.Items.Add($Department) | Out-Null
    }
    $UserInfoTab.Controls.Add($Dep_DD)

    # User Type Dropdown
    $UT_DD = New-Object System.Windows.Forms.ComboBox
    $UT_DD.Size = New-Object System.Drawing.Size(200,20)
    $UT_DD.Location = New-Object System.Drawing.Point(105,100)
    $UT_DD.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    # Add options to UT_DD
    ForEach ($UserType in $UserTypes) {
        $UT_DD.Items.Add($UserType) | Out-Null
    }
    $UserInfoTab.Controls.Add($UT_DD)


    ### Create Form Items for Licenses Tab ###
    ##########################################

    ### Add Labels ###

    # Fulltime Employee / Intern Label
    $FTE_Label = New-Object System.Windows.Forms.Label
    $FTE_Label.Size = New-Object System.Drawing.Size(400,20)
    $FTE_Label.Location = New-Object System.Drawing.Point(5,10)
    $FTE_Label.Text = "Fulltime Employee / Intern License Options"
    $LicensesTab.Controls.Add($FTE_Label)

    # Contractor Employee Label
    $Con_Label = New-Object System.Windows.Forms.Label
    $Con_Label.Size = New-Object System.Drawing.Size(400,20)
    $Con_Label.Location = New-Object System.Drawing.Point(5,190)
    $Con_Label.Text = "Contractor License Options"
    $LicensesTab.Controls.Add($Con_Label)

    # Mandatory License Label
    $ML_Label = New-Object System.Windows.Forms.Label
    $ML_Label.Size = New-Object System.Drawing.Size(400,20)
    $ML_Label.Location = New-Object System.Drawing.Point(5,340)
    $ML_Label.Text = "** One of these licenses is mandatory for the above User Type."
    $LicensesTab.Controls.Add($ML_Label)


    ### Add Checkboxes ###

    # Add Checkboxes for license options


    ### Add Dropdowns ###

    # Add options for Phone numbers


    ### Add Form Items for Optional Tab ###
    #######################################



    $MainForm.Add_Shown({$MainForm.Activate()})
    $Result = $MainForm.ShowDialog()

    Return $Result
}

# Verifies that all fields are correctly entered
function VerifyMandatoryFields{
    # Established return variables
    $verified = $true
    $message = ''

    # First Name not empty
    if ($FN_TB.Text -eq "") {
        $message += "First Name is missing`n"
        $verified = $false
    }

    # Last Name not empty
    if ($LN_TB.Text -eq "") {
        $message += "Last Name is missing`n"
        $verified = $false
    }

    # Verify user does not already exist
    if ($FN_TB.Text -ne "" -AND $LN_TB.Text -ne "") {
        # Add logic to put together email string to verify in Azure AD
        
        $userCheck = Get-AzureADUser -SearchString $checkExist

        # If the user exists already in Azure AD based on email, throw error
        if ($userCheck) {
            $message += "A user account with this email already exists`n"
            $verified = $false
        }
    }

    # Department not empty
    if ($Dep_DD.Text -eq "") {
        $message += "Department is missing`n"
        $verified = $false
    }

    # User Type not empty
    if ($UT_DD.Text -eq "") {
        $message += "User Type is missing`n"
        $verified = $false
    }

    # Title not empty
    if ($Title_TB.Text -eq "") {
        $message += "Title is missing`n"
        $verified = $false
    }

    # Manager Email not empty
    if ($ME_TB.Text -eq "") {
        $message += "Manager Email is missing`n"
        $verified = $false
    }
    
    # Verifies that the manager email exists in Azure
    else {
        $managerCheck = Get-AzureADUser -SearchString $ME_TB.Text  
        if (!($managerCheck)) {
            $message += "The manager does not exist`n"
            $verified = $false
        }        
    }

    # Verifies that the mirror email is present in Azure
    if ($MG_TB.Text -ne "") {
        $mirrorCheck = Get-AzureADUser -SearchString $MG_TB.Text  
        if (!($mirrorCheck)) {
            $message += "The mirror group email does not exist`n"
            $verified = $false
        }  
    }

    # Password not missing
    if ($PW_MTB.Text -eq "") {
        $message += "Password is missing`n"
        $verified = $false
    }

    # Checks licenses based on User type

    # Add logic to determine whether licenses are appropriate for selected User type

    # Returns the previous variables
    Return $verified, $message
}

function ConfirmFormResult(){
    # Enters all the information into the $currentUser array for use later
    $currentUser.FirstName = $FN_TB.Text
    $currentUser.LastName = $LN_TB.Text
    
    # Create Email Address string from First Name and Last Name
    

    $currentUser.Password = $PW_MTB.Text | ConvertTo-SecureString -AsPlainText -Force
    $currentUser.Department = $Dep_DD.Text
    $currentUser.Title = $Title_TB.Text
    $currentUser.Manager = $ME_TB.Text
    $currentUser.UserType = $UT_DD.Text
    $currentUser.MirrorGroup = $MG_TB.Text
    
    # Add All Licenses to Array

    $currentUser.Licenses.Clear()
    $currentUser.PhoneNumber = ''

    # Add logic for license checkboxes here.  For instance, if *LicenseName.Checked, add to $currentUser.Licenses array


    ### Create Confirm Form ###
    ###########################

    $ConfirmForm = New-Object System.Windows.Forms.Form
    $ConfirmForm.Text = "Confirm User Information"
    $ConfirmForm.Size = New-Object System.Drawing.Size(400,300)
    $ConfirmForm.KeyPreview = $True
    $ConfirmForm.FormBorderStyle = "1"
    $ConfirmForm.MaximizeBox = $false
    $ConfirmForm.StartPosition = "CenterScreen"
    $ConfirmForm.Topmost = $true
    $ConfirmForm.AutoSize = $True

    ### Create Buttons for Confirm Form

    # Confirm Button
    $CC_B = New-Object System.Windows.Forms.Button
    $CC_B.Size = New-Object System.Drawing.Size(75,25)
    $CC_B.Location = New-Object System.Drawing.Point(200,230)
    $CC_B.Text = "Confirm"
    $CC_B.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Right
    $ConfirmForm.Controls.Add($CC_B)
    $CC_B.add_click({
        $ConfirmForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })

    # Cancel Button
    $CanC_B = New-Object System.Windows.Forms.Button
    $CanC_B.Size = New-Object System.Drawing.Size(75,25)
    $CanC_B.Location = New-Object System.Drawing.Point(300,230)
    $CanC_B.Text = "Cancel"
    $CanC_B.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
        -bor [System.Windows.Forms.AnchorStyles]::Right
    $ConfirmForm.Controls.Add($CanC_B)
    $CanC_B.add_click({
        $ConfirmForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

    ### Create Confirm Form Label

    # Information Message Label
    $confirmMessage = "Name: " + $currentUser.FirstName + " " + $currentUser.LastName + "`nEmail: " + $currentUser.Email + "`nEmail Aliases: " + $currentUser.Alias + "`nPhone Number: " + $currentUser.PhoneNumber + "`nDepartment: " + $currentUser.Department + "`nTitle: " + $currentUser.Title + "`nManager Email: " + $currentUser.Manager + "`nUser Type: " + $currentUser.UserType + "`nLicenses to Add: " + $currentUser.Licenses + "`nMirror Groups On: " + $currentUser.MirrorGroup

    # Display ConfirmMessage in The Form
    $CM_Label = New-Object System.Windows.Forms.Label
    $CM_Label.Size = New-Object System.Drawing.Size(300,200)
    $CM_Label.Location = New-Object System.Drawing.Point(10,13)
    $CM_Label.Text = $confirmMessage
    $CM_Label.AutoSize = $True
    $CM_Label.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
    $ConfirmForm.Controls.Add($CM_Label)

    $ConfirmForm.Add_Shown({$ConfirmForm.Activate()})
    $ConfirmResult = $ConfirmForm.ShowDialog()

    # Returns the DialogResult to be used in the MainForm
    return $ConfirmResult
}


##################################################
### Create and Configure Azure AD User Account ###
##################################################

### Create the User Account

function createNewUser {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [String]$logPath
    )

    Write-Host "Creating User Profile in Azure AD...`n"

    # Puts together appropriate strings and objects to enter into Azure
    $fullName = $currentUser.FirstName + " " + $currentUser.LastName
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $currentUser.Password 
    $MailNickname = # Designate MailNickname which is a required parameter
    
    # Enters the new user into Azure and uses a variable to facilitate the user object
    $onboardUser = New-AzureADUser -AccountEnabled $true -DisplayName $fullName -UserPrincipalName $currentUser.Email -MailNickName $MailNickname -OtherMails $currentUser.Alias -Department $currentUser.Department -JobTitle $currentUser.Title -PasswordProfile $PasswordProfile -UsageLocation # Specify Usage Locationi
    
    # Logs user creation
    if ($onboardUser.UserPrincipalName -eq $currentUser.Email) {
        Write-Host "User Has Been Successfully Created`n"
        "[" +(Get-Date -Format "HH:mm:ss") + "] " + "User has been created with the above parameters.  See below for Azure AD User dump." | Add-Content -Path $logPath -NoNewline
        (Get-AzureADUser -ObjectID $onboardUser.ObjectID | Select-Object DisplayName,UserPrincipalName,OtherMails,Department,JobTitle) >> $logPath
    }
    else {
        Write-Host "There May Have Been An Error With User Creation.  Script exiting`n"
        "[" +(Get-Date -Format "HH:mm:ss") + "] " + "FATAL ERROR: Could not find selected user after creation.  Script closed due to fatal error" >> $logPath
        exit
    }

    Write-Host "Assigning Manager...`n"

    # Assigns manager based on $currentUser.Manager
    $managerUser = Get-AzureADUser -ObjectId $currentUser.Manager
    Set-AzureADUserManager -ObjectId $onboardUser.ObjectId -RefObjectId $managerUser.ObjectId

    # Logs manager assignment
    if ((Get-AzureADUserManager -ObjectID $onboardUser.ObjectID).UserPrincipalName -eq $managerUser.UserPrincipalName) {
        Write-Host "Manager Has Been Added Successfully`n"
        "[" +(Get-Date -Format "HH:mm:ss") + "] " + (Get-AzureADUserManager -ObjectID $onboardUser.ObjectID).DisplayName + " added as " + ($onboardUser.DisplayName) + "'s Manager." >> $logPath
    }
    else {
        Write-Host "There May Have Been An Error With Adding Manager, Please Check Log`n"
        "[" +(Get-Date -Format "HH:mm:ss") + "] " + "ERROR: Manager was not the created as expected.  Please correct the manager in Azure AD."
    }

    Write-Host "Assigning Selected Licenses...`n"

    # Assigns licenses
    assignLicenses $onboardUser $logPath

    Write-Host "Assigning Selected Groups...`n"

    # Assigns groups
    assignGroups $onboardUser $logPath

    Write-Host "Working on Final Tasks...`n"

    # Runs final misc tasks
    finalTasks $onboardUser $logPath
}


### Assign Licenses to New User

function assignLicenses {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [Microsoft.Open.AzureAD.Model.User]$onboardUser,

        [Parameter(Mandatory=$true, Position=1)]
        [String]$logPath
    )

    "`n[" +(Get-Date -Format "HH:mm:ss") + "] " + "Logging Assigned Licenses:" >> $logPath
    
    # Runs through the $currentUser.Licenses array and assigns each one to the user
    foreach ($license in $currentUser.Licenses) {
        # Creates a license object using the $availableLicenseChoice array
        $currentLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $licenseObject = Get-AzureADSubscribedSku -ObjectID $availableLicenseChoice[$license]
        $currentLicense.SkuID = $licenseObject.SkuID

        $licensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $licensesToAssign.AddLicenses = $currentLicense

        # Assigns that license object to the user
        Set-AzureADUserLicense -ObjectID $onboardUser.ObjectID -AssignedLicenses $licensesToAssign
        
        # Uses the licenses array to verify and then log
        $logLicense = Get-AzureADUserLicenseDetail -ObjectID $onboardUser.ObjectID | Where-Object {$_.SkuID -eq $availableLicenseChoice[$license]}

        if ($logLicense) {
            $license + " has been assigned" >> $logPath
        }

        else {
            "ERROR: " + " has not been assigned" >> $logPath
        }
    }

    Write-Host "Licenses Assigned, Check Log For Details`n"
}



### Assign New User to Groups

function assignGroups {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [Microsoft.Open.AzureAD.Model.User]$onboardUser,

        [Parameter(Mandatory=$true, Position=1)]
        [String]$logPath
    )
    
    # Check to make sure the user mailbox has been created
    while($true) {
        $userMailboxPresent = Get-Mailbox -Identity $onboardUser.UserPrincipalName -ErrorAction SilentlyContinue
        if (!($userMailboxPresent)) {
            Write-Host "User Mailbox Has Not Yet Been Instantiated.  Checking Again in 60 Seconds"
            Start-Sleep -Seconds 60
        }

        else {
            break
        }
    }

    "`n[" +(Get-Date -Format "HH:mm:ss") + "] " + "Logging Assigned Groups:" >> $logPath

    # Add User to specific group based on User Type and logs after verification
    Switch ($currentUser.UserType) {
        # Logic to add users to specific groups based on User Type
    }

    # Adds users to non security enabled groups based on mirror email if exists
    if ($currentUser.MirrorGroup -ne "") {
        # Gets assigned groups based on mirror group email address
        $mirrorUser = Get-AzureADUser -SearchString $currentUser.MirrorGroup
        $groupsToMirror = Get-AzureADUserMembership -ObjectID $mirrorUser.ObjectID | Where-Object {$_.ObjectType -eq "Group" -AND $_.SecurityEnabled -eq $false}

        # Cycle through each group and add new user to the group
        foreach ($group in $groupsToMirror) {
            try {
                # This will only work for 365 groups
                Add-AzureADGroupMember -ObjectID $group.ObjectID -RefObjectId $onboardUser.ObjectID 
                $checkGroup = Get-AzureADGroupMember -ObjectID $group.ObjectID | Where-Object {$_.ObjectID -eq $onboardUser.ObjectID}
                
                # Log on successful and go to next group
                if ($checkGroup) {
                    Write-Host "User Added to" $group.DisplayName "`n"
                    "Added to " + $group.DisplayName >> $logPath
                }
            }

            catch {
                # If the group is not a 365 and an error occurs, try to add to Distribution Group
                Write-Host "Group Was Not A 365 Group, Trying Distribution / Security Group...`n"
                Add-DistributionGroupMember -Identity $group.DisplayName -Member $onboardUser.UserPrincipalName
                $checkDistroGroup = Get-DistributionGroupMember -Identity $group.DisplayName | Where-Object {$_.PrimarySMTPAddress -eq $onboardUser.UserPrincipalName}
                
                # Verifies and logs group membership
                if ($checkDistroGroup) {
                    Write-Host "User Added to" $group.DisplayName "`n"
                    "Added to " + $group.DisplayName >> $logPath
                }
                else {
                    Write-Host "User May Not Have Been Added to" $group.DisplayName "`n"
                    "ERROR: User may not have been added to " + $group.DisplayName + ".  Please assign manually." >> $logPath
                }
            }
        }
    }
}


# Run misc final set up tasks

function finalTasks {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [Microsoft.Open.AzureAD.Model.User]$onboardUser,

        [Parameter(Mandatory=$true, Position=1)]
        [String]$logPath
    )

    # Add any extra tasks that you may always run on a new user
}


# Create Log file and pass path

function createLogFile {
    Write-Host "`nCreating New Log File at 'C:\OnboardingLogFiles\'`n"

    # Createa strings to facilitate file creation
    $folderName = "OnboardingLogFiles"
    $path = "C:\OnboardingLogFiles"
    $logName = $currentUser.Email + "_" + (Get-Date -Format "yyyyMMdd").ToString() + ".txt"
    $logFilePath = $path + "\" + $logName

    # Check if the folder already exists at C:\ and create if not
    if (!(Test-Path $path)) {
        New-Item -ItemType Directory -Path "C:\" -Name $folderName | Out-Null
    }

    # Create new log file useing the $logName string
    New-Item -ItemType File -Path $path -Name $logName | Out-Null

    # Write design elements to the log file
    # Make a cool design for your Log file here and add that to the file

    # Write intro to the log file
    $introLine = "`nOnboarding Script for " + $currentUser.FirstName + " " + $currentUser.LastName + "`nScript Run On " + (Get-Date -Format "yyyy/MM/dd").ToString() + "`n`n"
    $introLine >> $logFilePath

    # Write $currentUser dump to the log file for reference
    $informationLine = "Information Entered Into Form:`n----------------------------`nName: " + $currentUser.FirstName + " " + $currentUser.LastName + "`nEmail: " + $currentUser.Email + "`nEmail Aliases: " + $currentUser.Alias + "`nPhone Number: " + $currentUser.PhoneNumber + "`nDepartment: " + $currentUser.Department + "`nTitle: " + $currentUser.Title + "`nManager Email: " + $currentUser.Manager + "`nUser Type: " + $currentUser.UserType + "`nLicenses to Add: " + $currentUser.Licenses + "`nMirror Groups On: " + $currentUser.MirrorGroup + "`n----------------------------`n"
    $informationLine >> $logFilePath

    # Return the log file path
    return $logFilePath
}

#####################
### Main Function ###
#####################

# Run GetConnected to verify connection to necessary modules
$connected = GetConnected

# If connection succeeded, run the rest of the script
if ($connected) {
    # Check licenses before anything else
    checkCurrentLicenses

    # Run the form and enter Dialog Result into veriable
    $FormResult = CreateForm

    # If the DialogResult is OK, start user creation
    if ($FormResult -eq [System.Windows.Forms.DialogResult]::OK) {
        # Create log file and enter filepath into a variable
        $logPath = createLogFile

        # Run createNewUser and pass the pathfile for the log
        createNewUser($logPath)

        # Write closing information to the console and the log file.
        Write-Host "-- User Creation Completed. `n-- Please find the log file at" $logPath ". `n-- Print this file to PDF and attach to the ticket."
    }
}

# If you are not able to connect, end the program and print the below message
else {
    Write-Host "We were not able to get you connected.  Please try again."
}