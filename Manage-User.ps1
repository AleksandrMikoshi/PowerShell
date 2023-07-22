<#Manage-User - menu for command for create, edit, delete account in Active Directory end Exchange.
The data for creating users is collected from the ticket in JIRA. Use module from Atlassian company - https://atlassianps.org/docs/JiraPS/
Manage-User — меню для команды создания, редактирования, удаления учетной записи в Active Directory и Exchange.
Данные для создания пользователей собираются из заявки в JIRA. Используйте официальный модуль от компании Atlassian - https://atlassianps.org/docs/JiraPS/ #>

function Manage-user{
    Add-Type -assembly System.Windows.Forms

    $global:DC = 'Domain Controller'
    $global:Mail_serv = 'Mail Server'
    $global:Jira = 'URL Jira Server'
    $global:UserName = "Jira user"
    $global:SecurePasswd = "Jira password" | ConvertTo-SecureString -AsPlainText -Force
    $global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePasswd
    $global:Path_log = 'Path logs'
    $global:Path_OU = 'OU=Users,DC=My,DC=Company,DC=com'
    $global:Path_fired = 'OU=Fired_users,DC=My,DC=Company,DC=com'
    $global:Path_OU_Outstaff = 'OU=Outstaff,DC=My,DC=Company,DC=com'
    $domain = 'Company.com'
    $Label_Text = "Company name or other"

    $Login_Mail = 'Mail server login'
    $Password_Mail = 'Mail server password' | ConvertTo-SecureString -AsPlainText -Force
    $Cred_Exch = New-Object System.Management.Automation.PSCredential ($Login_Mail, $Password_Mail)    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Add-MailboxPermission,Enable-Mailbox,New-Mailbox,Disable-mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Update-EmailAddressPolicy,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList,Add-ADPermission,Set-Mailbox, Get-Mailbox -DisableNameChecking
 
    <# Menu Active Directory
    Меню Active Directory #>
    $ActiveDirectoryMenu = {
        <# Form for an issue in Jira
        Форма для указания задачи в Jira #>
        $Ticket_Jira_User = {
            $Button_Ticket_Jira = New-Object System.Windows.Forms.Button
            $Button_Ticket_Jira.Location = New-Object System.Drawing.Size(400,50)
            $Button_Ticket_Jira.AutoSize = $true
            $Button_Ticket_Jira.Text = "Continue"
            $Button_Ticket_Jira.Add_click({Info_From_Jira -PipeLine $TextBox_Ticket_Jira.Text})

            $New_Label = New-Object System.Windows.Forms.Label
            $New_Label.Text = "Specify the task number"
            $New_Label.Location  = New-Object System.Drawing.Point(10,20)
            $New_Label.AutoSize = $true

            $TextBox_Ticket_Jira = New-Object System.Windows.Forms.TextBox
            $TextBox_Ticket_Jira.Location  = New-Object System.Drawing.Point(10,50)
            $TextBox_Ticket_Jira.Text = ''
            $TextBox_Ticket_Jira.Size = New-Object System.Drawing.Size(270,30)

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Ticket_Jira.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Form_Ticket_Jira = New-Object System.Windows.Forms.Form
            $Form_Ticket_Jira.Text ='Account Management Active Directory'
            $Form_Ticket_Jira.Width = 500
            $Form_Ticket_Jira.Height = 500
            $Form_Ticket_Jira.AutoSize = $true
            $Form_Ticket_Jira.Controls.Add($Button_Ticket_Jira)
            $Form_Ticket_Jira.Controls.Add($New_Label)
            $Form_Ticket_Jira.Controls.Add($TextBox_Ticket_Jira)
            $Form_Ticket_Jira.Controls.Add($Label)
            $Form_Ticket_Jira.Controls.Add($Button_Back)
            $Form_Ticket_Jira.Controls.Add($Button_Exit)

            $ActiveDirectoryMenu.Hide()
            $Form_Ticket_Jira.ShowDialog()
        }
        <# Getting information from Jira
        Получение информации из Jira #>
        function Info_From_Jira{
            param (
                [Parameter(Mandatory)]
                [string]$PipeLine
            )   
            Set-JiraConfigServer -Server "$Jira"
            $Full_Data = Get-JiraIssue -Issue "$PipeLine" -Credential $cred -Fields customfield_10409,customfield_10708,customfield_10408,customfield_10506,customfield_10410,customfield_10502,customfield_10503,customfield_10505,customfield_10501,customfield_10700,Key #получение данных из Jira
            [array]$Array_City               =   ($Full_Data.customfield_10501)
            [System.DateTime]$Data           =   get-date ($Full_Data.customfield_10505)
            [string]$Global:Manager          =   $Full_Data.customfield_10708.name
            [string]$Global:LastName         =   ($Full_Data.customfield_10408).Split(" ")[0]
            [string]$Global:FirstName        =   ($Full_Data.customfield_10408).Split(" ")[1]
            [string]$Global:Initials         =   ($Full_Data.customfield_10408).Split(" ")[2]
            [string]$Global:SamAccountName   =   ($Full_Data.customfield_10506).ToLower()
            [string]$Global:Office           =   $Array_City.Value
            [string]$Global:JobTitle         =   $Full_Data.customfield_10410
            [string]$Global:Department       =   $Full_Data.customfield_10502
            [string]$Global:Phone            =   $Full_Data.customfield_10503
            [string]$Global:Birth            =   Get-Date $Data -format "dd.MM.yyyy"

            $Label_Info_User = New-Object System.Windows.Forms.Label
            $Label_Info_User.Text = "Check the data"
            $Label_Info_User.Location = New-Object System.Drawing.Point(10,20)
            $Label_Info_User.AutoSize = $true

            $Label_ticket_1 = New-Object System.Windows.Forms.Label
            $Label_ticket_1.Text = "Last name"
            $Label_ticket_1.Location = New-Object System.Drawing.Point(10,75)
            $Label_ticket_1.AutoSize = $true

            $TextBox_LastName = New-Object System.Windows.Forms.TextBox
            $TextBox_LastName.Location  = New-Object System.Drawing.Point(150,75)
            $TextBox_LastName.Text = $LastName
            $TextBox_LastName.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_2 = New-Object System.Windows.Forms.Label
            $Label_ticket_2.Text = "Firts name"
            $Label_ticket_2.Location = New-Object System.Drawing.Point(10,100)
            $Label_ticket_2.AutoSize = $true

            $TextBox_FirstName = New-Object System.Windows.Forms.TextBox
            $TextBox_FirstName.Location  = New-Object System.Drawing.Point(150,100)
            $TextBox_FirstName.Text = $FirstName
            $TextBox_FirstName.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_3 = New-Object System.Windows.Forms.Label
            $Label_ticket_3.Text = "Middle name"
            $Label_ticket_3.Location = New-Object System.Drawing.Point(10,125)
            $Label_ticket_3.AutoSize = $true

            $TextBox_Initials = New-Object System.Windows.Forms.TextBox
            $TextBox_Initials.Location  = New-Object System.Drawing.Point(150,125)
            $TextBox_Initials.Text = $Initials
            $TextBox_Initials.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_4 = New-Object System.Windows.Forms.Label
            $Label_ticket_4.Text = "Login"
            $Label_ticket_4.Location = New-Object System.Drawing.Point(10,150)
            $Label_ticket_4.AutoSize = $true

            $TextBox_SamAccountName = New-Object System.Windows.Forms.TextBox
            $TextBox_SamAccountName.Location  = New-Object System.Drawing.Point(150,150)
            $TextBox_SamAccountName.Text = $SamAccountName
            $TextBox_SamAccountName.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_5 = New-Object System.Windows.Forms.Label
            $Label_ticket_5.Text = "Office"
            $Label_ticket_5.Location = New-Object System.Drawing.Point(10,175)
            $Label_ticket_5.AutoSize = $true

            $TextBox_Office = New-Object System.Windows.Forms.TextBox
            $TextBox_Office.Location  = New-Object System.Drawing.Point(150,175)
            $TextBox_Office.Text = $Office
            $TextBox_Office.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_6 = New-Object System.Windows.Forms.Label
            $Label_ticket_6.Text = "Department"
            $Label_ticket_6.Location = New-Object System.Drawing.Point(10,200)
            $Label_ticket_6.AutoSize = $true

            $TextBox_Department = New-Object System.Windows.Forms.TextBox
            $TextBox_Department.Location  = New-Object System.Drawing.Point(150,200)
            $TextBox_Department.Text = $Department
            $TextBox_Department.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_7 = New-Object System.Windows.Forms.Label
            $Label_ticket_7.Text = "Job title"
            $Label_ticket_7.Location = New-Object System.Drawing.Point(10,275)
            $Label_ticket_7.AutoSize = $true

            $TextBox_JobTitle = New-Object System.Windows.Forms.TextBox
            $TextBox_JobTitle.Location  = New-Object System.Drawing.Point(150,275)
            $TextBox_JobTitle.Text = $JobTitle
            $TextBox_JobTitle.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_8 = New-Object System.Windows.Forms.Label
            $Label_ticket_8.Text = "Manager"
            $Label_ticket_8.Location = New-Object System.Drawing.Point(10,300)
            $Label_ticket_8.AutoSize = $true

            $TextBox_Manager = New-Object System.Windows.Forms.TextBox
            $TextBox_Manager.Location  = New-Object System.Drawing.Point(150,300)
            $TextBox_Manager.Text = $Manager
            $TextBox_Manager.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_9 = New-Object System.Windows.Forms.Label
            $Label_ticket_9.Text = "Phone number"
            $Label_ticket_9.Location = New-Object System.Drawing.Point(10,325)
            $Label_ticket_9.AutoSize = $true

            $TextBox_Phone = New-Object System.Windows.Forms.TextBox
            $TextBox_Phone.Location  = New-Object System.Drawing.Point(150,325)
            $TextBox_Phone.Text = $Phone
            $TextBox_Phone.Size = New-Object System.Drawing.Size(300,30)

            $Label_ticket_10 = New-Object System.Windows.Forms.Label
            $Label_ticket_10.Text = "Brithday"
            $Label_ticket_10.Location = New-Object System.Drawing.Point(10,350)
            $Label_ticket_10.AutoSize = $true

            $TextBox_Birth = New-Object System.Windows.Forms.TextBox
            $TextBox_Birth.Location  = New-Object System.Drawing.Point(150,350)
            $TextBox_Birth.Text = $Birth
            $TextBox_Birth.Size = New-Object System.Drawing.Size(300,30)

            $Button_Info_User = New-Object System.Windows.Forms.Button
            $Button_Info_User.Location = New-Object System.Drawing.Size(400,20)
            $Button_Info_User.AutoSize = $true
            $Button_Info_User.Text = "Create account"
            $Button_Info_User.Add_click(
                {
                    $UserAttr = @{
                        LastName = $TextBox_LastName.Text
                        FirstName = $TextBox_FirstName.Text
                        Initials = $TextBox_Initials.Text
                        SamAccountName = $TextBox_SamAccountName.Text
                        Office = $TextBox_Office.Text
                        JobTitle = $TextBox_JobTitle.Text
                        Department = $TextBox_Department.Text
                        Phone = $TextBox_Phone.Text
                        Brith = $TextBox_Birth.Text
                        PipeLine = $PipeLine
                        Manager = $TextBox_Manager.Text}
                    Create_User -UserAttr $UserAttr -DC $DC -Path_OU $Path_OU
                }
            )
            
            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Info_User.Dispose(),$Form_Ticket_Jira.Show()})
            $Button_Back.AutoSize = $true

            $Form_Info_User = New-Object System.Windows.Forms.Form
            $Form_Info_User.Text ='Account Management Active Directory'
            $Form_Info_User.Width = 500
            $Form_Info_User.Height = 500
            $Form_Info_User.AutoSize = $true
            $Form_Info_User.Controls.Add($Label_Info_User)
            $Form_Info_User.Controls.Add($Label_ticket_1)
            $Form_Info_User.Controls.Add($TextBox_LastName)
            $Form_Info_User.Controls.Add($Label_ticket_2)
            $Form_Info_User.Controls.Add($TextBox_FirstName)
            $Form_Info_User.Controls.Add($Label_ticket_3)
            $Form_Info_User.Controls.Add($TextBox_Initials)
            $Form_Info_User.Controls.Add($Label_ticket_4)
            $Form_Info_User.Controls.Add($TextBox_SamAccountName)
            $Form_Info_User.Controls.Add($Label_ticket_5)
            $Form_Info_User.Controls.Add($TextBox_Office)
            $Form_Info_User.Controls.Add($Label_ticket_6)
            $Form_Info_User.Controls.Add($TextBox_Department)
            $Form_Info_User.Controls.Add($Label_ticket_7)
            $Form_Info_User.Controls.Add($TextBox_JobTitle)
            $Form_Info_User.Controls.Add($Label_ticket_8)
            $Form_Info_User.Controls.Add($TextBox_Manager)
            $Form_Info_User.Controls.Add($Label_ticket_9)
            $Form_Info_User.Controls.Add($TextBox_Phone)
            $Form_Info_User.Controls.Add($Label_ticket_10)
            $Form_Info_User.Controls.Add($TextBox_Birth)
            $Form_Info_User.Controls.Add($Button_Info_User)
            $Form_Info_User.Controls.Add($Button_Back)
            $Form_Info_User.Controls.Add($Button_Exit)
            $Form_Info_User.Controls.Add($Label)

            $Form_Ticket_Jira.Hide()
            $Form_Info_User.ShowDialog()

        }
        <# Create account
        Создание пользователя #>
        function Create_User {
            param (
                [Parameter(Mandatory)]
                [hashtable]$UserAttr,
                [Parameter(Mandatory)]
                [string]$DC,
                [Parameter(Mandatory)]
                [string]$Path_OU
            )
                $LastName        = $UserAttr.LastName
                $FirstName       = $UserAttr.FirstName
                $Initials        = $UserAttr.Initials
                $SamAccountName  = $UserAttr.SamAccountName
                $Office          = $UserAttr.Office
                $JobTitle        = $UserAttr.JobTitle
                $Department      = $UserAttr.Department
                $Phone           = $UserAttr.Phone
                $Birth           = $UserAttr.Brith
                $PipeLine        = $UserAttr.PipeLine
                $Manager         = $UserAttr.Manager
                try{
                    $UserPrincipalName = $SamAccountName + "@" + $domain
                    $DisplayName = $LastName + " " + $FirstName
                    $FullName = $LastName + " " + $FirstName + " " + $Initials
                    Add-Type -AssemblyName System.Web
                    $length = Get-Random -Minimum 9 -Maximum 12
                    $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
                    New-ADUser -Server $DC `
                    -Name $FullName `
                    -DisplayName $DisplayName `
                    -GivenName $FirstName `
                    -Office $Office `
                    -Surname $LastName `
                    -OfficePhone $Phone `
                    -Department $Department `
                    -Manager $Manager `
                    -Title $JobTitle `
                    -UserPrincipalName $UserPrincipalName `
                    -SamAccountName $SamAccountName `
                    -Path $Path_OU `
                    -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force) `
                    -ChangePasswordAtLogon $true `
                    -Enabled $true
                    Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute1 = $Birth }
                    Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute2 = $Initials }
                    Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute15 = "$Birth" }
    
                    $user_creds = "Login: $SamAccountName`nPassword: $password"
                    $parameters = @{
                        Fields = @{
                            customfield_10641 = "$user_creds"
                        }
                    }
                    Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $cred
    
                    Enable-Mailbox -identity $SamAccountName -Alias $SamAccountName -Database DAG_DB02 -DomainController $DC
                    Enable-Mailbox -identity $SamAccountName -archive -ArchiveDatabase ARCHIVE -DomainController $DC
                    Get-GlobalAddressList | Update-GlobalAddressList
                    Get-OfflineAddressBook | Update-OfflineAddressBook
                    Get-AddressList | Update-AddressList

                    $Color = "green"
                    $Outcome="
                        Account $LastName $FirstName create
                        Mailbox '$SamAccountName@m2.ru' create"
                }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value

                    $Color = "red"
                    $Outcome = 'Execution failed. Contact your system administrator'
                }
                finally{
                    $Text_Create_User = New-Object System.Windows.Forms.Label
                    $Text_Create_User.Text = $Outcome
                    $Text_Create_User.ForeColor = $Color
                    $Text_Create_User.Location = New-Object System.Drawing.Point(10,20)
                    $Text_Create_User.AutoSize = $true

                    $Button_Exit = New-Object System.Windows.Forms.Button
                    $Button_Exit.Text = 'Exit'
                    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $Button_Exit.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true

                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_Create_User.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_Create_User = New-Object System.Windows.Forms.Form
                    $Form_Create_User.Text ='Account Management Active Directory'
                    $Form_Create_User.Width = 500
                    $Form_Create_User.Height = 500
                    $Form_Create_User.AutoSize = $true
                    $Form_Create_User.Controls.Add($Text_Create_User)
                    $Form_Create_User.Controls.Add($Button_Back)
                    $Form_Create_User.Controls.Add($Button_Exit)
                    $Form_Create_User.Controls.Add($Label)

                    $Form_Info_User.Hide()
                    $Form_Create_User.ShowDialog()
                }
        }
        <# Form for an issue in Jira
        Форма для указания задачи в Jira #>
        $Ticket_Jira_Outstaff = {
            $Button_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.Button
            $Button_Ticket_Jira_Outstaff.Location = New-Object System.Drawing.Size(400,50)
            $Button_Ticket_Jira_Outstaff.AutoSize = $true
            $Button_Ticket_Jira_Outstaff.Text = "Continue"
            $Button_Ticket_Jira_Outstaff.Add_click({Info_User_Outstaff -PipeLine $TextBox_Ticket_Jira.Text})

            $New_Label = New-Object System.Windows.Forms.Label
            $New_Label.Text = "Specify the task number"
            $New_Label.Location  = New-Object System.Drawing.Point(10,20)
            $New_Label.AutoSize = $true

            $TextBox_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.TextBox
            $TextBox_Ticket_Jira_Outstaff.Location  = New-Object System.Drawing.Point(10,50)
            $TextBox_Ticket_Jira_Outstaff.Text = ''
            $TextBox_Ticket_Jira_Outstaff.Size = New-Object System.Drawing.Size(270,30)

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Ticket_Jira_Outstaff.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.Form
            $Form_Ticket_Jira_Outstaff.Text ='Account Management Active Directory'
            $Form_Ticket_Jira_Outstaff.Width = 500
            $Form_Ticket_Jira_Outstaff.Height = 500
            $Form_Ticket_Jira_Outstaff.AutoSize = $true
            $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Ticket_Jira_Outstaff)
            $Form_Ticket_Jira_Outstaff.Controls.Add($New_Label)
            $Form_Ticket_Jira_Outstaff.Controls.Add($TextBox_Ticket_Jira_Outstaff)
            $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Back)
            $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Exit)
            $Form_Ticket_Jira_Outstaff.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Ticket_Jira_Outstaff.ShowDialog()
        }
        <# Getting information from Jira
        Получение информации из Jira #>
        function Info_User_Outstaff {
            param (
                [Parameter(Mandatory)]
                [string]$PipeLine
            )   
            process{
                Set-JiraConfigServer -Server "$Jira"
                $Full_Data = Get-JiraIssue -Issue "$PipeLine" -Credential $cred -Fields customfield_10708,customfield_10408,customfield_10506,customfield_10502,customfield_10503,customfield_11305,customfield_10501 #получение данных из Jira
                [array]$Array_City = $Full_Data.customfield_10501
                [string]$LastName         =   ($Full_Data.customfield_10408).Split(" ")[0]
                [string]$FirstName        =   ($Full_Data.customfield_10408).Split(" ")[1]
                [string]$Initials         =   ($Full_Data.customfield_10408).Split(" ")[2]
                [string]$SamAccountName   =   ($Full_Data.customfield_10506).ToLower()
                [string]$Office           =   $Array_City.Value
                [string]$Department       =   $Full_Data.customfield_10502
                [string]$Phone            =   $Full_Data.customfield_10503
                [string]$Company          =   $Full_Data.customfield_11305
            
                $Text_Outstaff_Info_User = New-Object System.Windows.Forms.Label
                $Text_Outstaff_Info_User.Text = "Check data"
                $Text_Outstaff_Info_User.Location = New-Object System.Drawing.Point(10,20)
                $Text_Outstaff_Info_User.AutoSize = $true

                $Label_ticket_1 = New-Object System.Windows.Forms.Label
                $Label_ticket_1.Text = "Last name"
                $Label_ticket_1.Location = New-Object System.Drawing.Point(10,50)
                $Label_ticket_1.AutoSize = $true

                $TextBox_LastName = New-Object System.Windows.Forms.TextBox
                $TextBox_LastName.Location  = New-Object System.Drawing.Point(150,50)
                $TextBox_LastName.Text = $LastName
                $TextBox_LastName.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_2 = New-Object System.Windows.Forms.Label
                $Label_ticket_2.Text = "First name"
                $Label_ticket_2.Location = New-Object System.Drawing.Point(10,75)
                $Label_ticket_2.AutoSize = $true

                $TextBox_FirstName = New-Object System.Windows.Forms.TextBox
                $TextBox_FirstName.Location  = New-Object System.Drawing.Point(150,75)
                $TextBox_FirstName.Text = $FirstName
                $TextBox_FirstName.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_3 = New-Object System.Windows.Forms.Label
                $Label_ticket_3.Text = "Middle name"
                $Label_ticket_3.Location = New-Object System.Drawing.Point(10,100)
                $Label_ticket_3.AutoSize = $true

                $TextBox_Initials = New-Object System.Windows.Forms.TextBox
                $TextBox_Initials.Location  = New-Object System.Drawing.Point(150,100)
                $TextBox_Initials.Text = $Initials
                $TextBox_Initials.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_4 = New-Object System.Windows.Forms.Label
                $Label_ticket_4.Text = "Login"
                $Label_ticket_4.Location = New-Object System.Drawing.Point(10,125)
                $Label_ticket_4.AutoSize = $true

                $TextBox_SamAccountName = New-Object System.Windows.Forms.TextBox
                $TextBox_SamAccountName.Location  = New-Object System.Drawing.Point(150,125)
                $TextBox_SamAccountName.Text = $SamAccountName
                $TextBox_SamAccountName.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_5 = New-Object System.Windows.Forms.Label
                $Label_ticket_5.Text = "Office"
                $Label_ticket_5.Location = New-Object System.Drawing.Point(10,150)
                $Label_ticket_5.AutoSize = $true

                $TextBox_Office = New-Object System.Windows.Forms.TextBox
                $TextBox_Office.Location  = New-Object System.Drawing.Point(150,150)
                $TextBox_Office.Text = $Office
                $TextBox_Office.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_6 = New-Object System.Windows.Forms.Label
                $Label_ticket_6.Text = "Department"
                $Label_ticket_6.Location = New-Object System.Drawing.Point(10,175)
                $Label_ticket_6.AutoSize = $true

                $TextBox_Department = New-Object System.Windows.Forms.TextBox
                $TextBox_Department.Location  = New-Object System.Drawing.Point(150,175)
                $TextBox_Department.Text = $Department
                $TextBox_Department.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_7 = New-Object System.Windows.Forms.Label
                $Label_ticket_7.Text = "Phone number"
                $Label_ticket_7.Location = New-Object System.Drawing.Point(10,200)
                $Label_ticket_7.AutoSize = $true

                $TextBox_Phone = New-Object System.Windows.Forms.TextBox
                $TextBox_Phone.Location  = New-Object System.Drawing.Point(150,200)
                $TextBox_Phone.Text = $Phone
                $TextBox_Phone.Size = New-Object System.Drawing.Size(300,30)

                $Label_ticket_8 = New-Object System.Windows.Forms.Label
                $Label_ticket_8.Text = "Company"
                $Label_ticket_8.Location = New-Object System.Drawing.Point(10,225)
                $Label_ticket_8.AutoSize = $true
            
                $TextBox_Company = New-Object System.Windows.Forms.TextBox
                $TextBox_Company.Location  = New-Object System.Drawing.Point(150,225)
                $TextBox_Company.Text = $Company
                $TextBox_Company.Size = New-Object System.Drawing.Size(300,30)

                $Button_Outstaff_Info_User = New-Object System.Windows.Forms.Button
                $Button_Outstaff_Info_User.Location = New-Object System.Drawing.Size(400,20)
                $Button_Outstaff_Info_User.AutoSize = $true
                $Button_Outstaff_Info_User.Text = "Create account"
                $Button_Outstaff_Info_User.Add_click(
                    {
                        $UserAttr = @{
                            LastName        = $TextBox_LastName.Text
                            FirstName       = $TextBox_FirstName.Text
                            Initials        = $TextBox_Initials.Text
                            SamAccountName  = $TextBox_SamAccountName.Text
                            Office          = $TextBox_Office.Text
                            Department      = $TextBox_Department.Text
                            Phone           = $TextBox_Phone.Text
                            Company         = $TextBox_Company.Text
                            }
                    Create_Outstaff_User -UserAttr $UserAttr -DC $DC
                    }
                )
            
                $Button_Exit = New-Object System.Windows.Forms.Button
                $Button_Exit.Text = 'Exit'
                $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $Button_Exit.AutoSize = $true

                $Label = New-Object System.Windows.Forms.Label
                $Label.Text = $Label_Text
                $Label.Location  = New-Object System.Drawing.Point(110,450)
                $Label.AutoSize = $true

                $Button_Back = New-Object System.Windows.Forms.Button
                $Button_Back.Text = 'Back'
                $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                $Button_Back.Add_click({$Form_Info_User.Dispose(),$Form_Ticket_Jira.Show()})
                $Button_Back.AutoSize = $true

                $Form_Outstaff_Info_User = New-Object System.Windows.Forms.Form
                $Form_Outstaff_Info_User.Text ='Account Management Active Directory'
                $Form_Outstaff_Info_User.Width = 500
                $Form_Outstaff_Info_User.Height = 500
                $Form_Outstaff_Info_User.AutoSize = $true
                $Form_Outstaff_Info_User.Controls.Add($Text_Outstaff_Info_User)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_1)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_LastName)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_2)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_FirstName)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_3)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_Initials)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_4)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_SamAccountName)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_5)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_Office)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_6)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_Department)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_7)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_Phone)
                $Form_Outstaff_Info_User.Controls.Add($Label_ticket_8)
                $Form_Outstaff_Info_User.Controls.Add($TextBox_Company)
                $Form_Outstaff_Info_User.Controls.Add($Button_Outstaff_Info_User)
                $Form_Outstaff_Info_User.Controls.Add($Button_Back)
                $Form_Outstaff_Info_User.Controls.Add($Button_Exit)
                $Form_Outstaff_Info_User.Controls.Add($Label)

                $Form_Ticket_Jira.Hide()
                $Form_Outstaff_Info_User.ShowDialog()
            }
        }
        <# Create account
        Создание пользователя #>
        function Create_Outstaff_User {
             param (
                [Parameter(Mandatory)]
                [hashtable]$UserAttr,
                [Parameter(Mandatory)]
                [string]$DC
                ) 
                $LastName        = $UserAttr.LastName
                $FirstName       = $UserAttr.FirstName
                $Initials        = $UserAttr.Initials
                $SamAccountName  = $UserAttr.SamAccountName
                $Office          = $UserAttr.Office
                $Department      = $UserAttr.Department
                $Phone           = $UserAttr.Phone
                $Company         = $UserAttr.Company
                try{
                    $UserPrincipalName = $SamAccountName + "@" + $domain
                    $FullName = $LastName + " " + $FirstName + " " + $Initials
                    $DisplayName = $LastName + " " + $FirstName
                    $length = Get-Random -Minimum 8 -Maximum 11
                    $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
                    New-ADUser -Server $DC `
                    -Name $FullName `
                    -DisplayName $DisplayName `
                    -GivenName $FirstName `
                    -Office $Office `
                    -Surname $LastName `
                    -OfficePhone $Phone `
                    -Department $Department `
                    -UserPrincipalName $UserPrincipalName `
                    -SamAccountName $SamAccountName `
                    -Path $Path_OU_Outstaff `
                    -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force) `
                    -ChangePasswordAtLogon $true `
                    -Enabled $true
                    Set-ADUser -Server $DC $SamAccountName -add @{company = $Company}
                    Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute2 = $Initials }

                    Add-ADGroupMember jira_users -Server $DC $SamAccountName
                    Add-ADGroupMember Confluence_Users -Server $DC $SamAccountName
                    Add-ADGroupMember owncloud_users -Server $DC $SamAccountName
    
                    $user_creds = "Login: $SamAccountName`nPassword: $password"
                    $parameters = @{
                        Fields = @{
                            customfield_10641 = "$user_creds"
                            }
                        }
                    Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $cred

                    $Color = "green"
                    $Outcome="
                    Account $LastName $FirstName create
                    Mailbox '$SamAccountName@m2.ru' create"
                }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value

                    $Color = "red"
                    $Outcome = 'Execution failed. Contact your system administrator'
                }
                finally{
                    $Text_Create_Outstaff_User = New-Object System.Windows.Forms.Label
                    $Text_Create_Outstaff_User.Text = $Outcome
                    $Text_Create_Outstaff_User.ForeColor = $Color
                    $Text_Create_Outstaff_User.Location = New-Object System.Drawing.Point(10,20)
                    $Text_Create_Outstaff_User.AutoSize = $true

                    $Button_Exit = New-Object System.Windows.Forms.Button
                    $Button_Exit.Text = 'Exit'
                    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $Button_Exit.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true

                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_Create_User.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_Create_Outstaff_User = New-Object System.Windows.Forms.Form
                    $Form_Create_Outstaff_User.Text ='Account Management Active Directory'
                    $Form_Create_Outstaff_User.Width = 500
                    $Form_Create_Outstaff_User.Height = 500
                    $Form_Create_Outstaff_User.AutoSize = $true
                    $Form_Create_Outstaff_User.Controls.Add($Text_Create_Outstaff_User)
                    $Form_Create_Outstaff_User.Controls.Add($Button_Back)
                    $Form_Create_Outstaff_User.Controls.Add($Button_Exit)
                    $Form_Create_Outstaff_User.Controls.Add($Label)

                    $Form_Outstaff_Info_User.Hide()
                    $Form_Create_Outstaff_User.ShowDialog()
                }
        }
        <# User selection FROM WHO and TO WHOM to copy groups
        Выбор пользователей ОТ КОГО и КОМУ копировать группы #>
        $Copy_group = {
            $Text_From_Copy_Group = New-Object System.Windows.Forms.Label
            $Text_From_Copy_Group.Text = "User from which we copy groups"
            $Text_From_Copy_Group.Location = New-Object System.Drawing.Point(10,30)
            $Text_From_Copy_Group.AutoSize = $true

            $ComboBox_From_Copy_Group = New-Object System.Windows.Forms.ComboBox
            $ComboBox_From_Copy_Group.Width = 250
            $FromUsers = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($FromUser in $FromUsers){
                $ComboBox_From_Copy_Group.Items.Add($FromUser.cn);
            }
            $ComboBox_From_Copy_Group.Location = New-Object System.Drawing.Point(10,50)

            $Text_To_Copy_Group = New-Object System.Windows.Forms.Label
            $Text_To_Copy_Group.Text = "User to whom we copy groups"
            $Text_To_Copy_Group.Location = New-Object System.Drawing.Point(10,80)
            $Text_To_Copy_Group.AutoSize = $true

            $ComboBox_To_Copy_Group = New-Object System.Windows.Forms.ComboBox
            $ComboBox_To_Copy_Group.Width = 250
            $ToUsers = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($ToUser in $ToUsers){
                $ComboBox_To_Copy_Group.Items.Add($ToUser.cn);
            }
            $ComboBox_To_Copy_Group.Location = New-Object System.Drawing.Point(10,100)

            $Button_Copy_Group_Cont = New-Object System.Windows.Forms.Button
            $Button_Copy_Group_Cont.Location = New-Object System.Drawing.Size(350,50)
            $Button_Copy_Group_Cont.AutoSize = $true
            $Button_Copy_Group_Cont.Text = "Copy group"
            $Button_Copy_Group_Cont.Add_Click(
                {
                    [string]$SourceUser = $ComboBox_From_Copy_Group.selectedItem
                    [string]$TargetUser = $ComboBox_To_Copy_Group.selectedItem
                    CopyGroup -SourceUser $SourceUser -TargetUser $TargetUser
                }
            )

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Copy_Group.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Copy_Group = New-Object System.Windows.Forms.Form
            $Form_Copy_Group.Text ='Account Management Active Directory'
            $Form_Copy_Group.Width = 500
            $Form_Copy_Group.Height = 500
            $Form_Copy_Group.AutoSize = $true
            $Form_Copy_Group.Controls.Add($Text_From_Copy_Group)
            $Form_Copy_Group.Controls.Add($ComboBox_From_Copy_Group)
            $Form_Copy_Group.Controls.Add($Text_To_Copy_Group)
            $Form_Copy_Group.Controls.Add($ComboBox_To_Copy_Group)
            $Form_Copy_Group.Controls.Add($Button_Copy_Group_Cont)
            $Form_Copy_Group.Controls.Add($Button_Back)
            $Form_Copy_Group.Controls.Add($Button_Exit)
            $Form_Copy_Group.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Copy_Group.ShowDialog()
        }
        <# Group copy function
        Функция копирования групп #>
        function CopyGroup {
            param (
                [Parameter(Mandatory)]
                $SourceUser,
                [Parameter(Mandatory)]
                $TargetUser
            )
            process{
                $SourceUserLogin = Get-ADUser -Filter "cn -eq '$SourceUser'" -Properties SamAccountName
                $TargetUserLogin = Get-ADUser -Filter "cn -eq '$TargetUser'" -Properties SamAccountName

                try{
                    $SourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUserLogin.SamAccountName | Where-Object {$_.name -ne "Domain Users"}
                    Add-ADPrincipalGroupMembership -Identity $TargetUserLogin.SamAccountName -MemberOf $sourceGroups

                    $Color = "green"
                    $Outcome="Groups from user $SourceUser copied"
                                    }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value

                    $Color = "red" 
                    $Outcome = 'Execution failed. Contact your system administrator'
                }
                finally{
                    $Text_Copy_group_done = New-Object System.Windows.Forms.Label
                    $Text_Copy_group_done.Text = $Outcome
                    $Text_Copy_group_done.ForeColor = $Color
                    $Text_Copy_group_done.Location = New-Object System.Drawing.Point(10,20)
                    $Text_Copy_group_done.AutoSize = $true
                
                    $Button_Exit = New-Object System.Windows.Forms.Button
                    $Button_Exit.Text = 'Exit'
                    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $Button_Exit.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true

                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_Copy_group_done.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_Copy_group_done = New-Object System.Windows.Forms.Form
                    $Form_Copy_group_done.Text ='Account Management Active Directory'
                    $Form_Copy_group_done.Width = 500
                    $Form_Copy_group_done.Height = 500
                    $Form_Copy_group_done.AutoSize = $true
                    $Form_Copy_group_done.Controls.Add($Text_Copy_group_done)
                    $Form_Copy_group_done.Controls.Add($Button_Exit)
                    $Form_Copy_group_done.Controls.Add($Button_Back)
                    $Form_Copy_group_done.Controls.Add($Label)

                    $Form_Copy_Group.Hide()
                    $Form_Copy_group_done.ShowDialog()
                }
            }
        } 
        <# Selecting a user to reset the password
        Выбор пользователя для сброса пароля #>
        $Reset_Pass = {
            $Text_Reset_Pass = New-Object System.Windows.Forms.Label
            $Text_Reset_Pass.Text = "Select user"
            $Text_Reset_Pass.Location = New-Object System.Drawing.Point(10,20)
            $Text_Reset_Pass.AutoSize = $true

            $ComboBox_Reset_Pass = New-Object System.Windows.Forms.ComboBox
            $ComboBox_Reset_Pass.Width = 250
            $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($User in $Users){
                $ComboBox_Reset_Pass.Items.Add($User.cn);
                }
            $ComboBox_Reset_Pass.Location = New-Object System.Drawing.Point(10,50)

            $Button_Reset_Pass = New-Object System.Windows.Forms.Button
            $Button_Reset_Pass.Location = New-Object System.Drawing.Size(350,50)
            $Button_Reset_Pass.AutoSize = $true
            $Button_Reset_Pass.Text = "Reset password"
            $Button_Reset_Pass.Add_Click({
                [string]$User = $ComboBox_Reset_Pass.selectedItem
                ResetPasswd -User $User
                }
            )

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Reset_Pass.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Reset_Pass = New-Object System.Windows.Forms.Form
            $Form_Reset_Pass.Text ='Account Management Active Directory'
            $Form_Reset_Pass.Width = 500
            $Form_Reset_Pass.Height = 500
            $Form_Reset_Pass.AutoSize = $true
            $Form_Reset_Pass.Controls.Add($Text_Reset_Pass)
            $Form_Reset_Pass.Controls.Add($ComboBox_Reset_Pass)
            $Form_Reset_Pass.Controls.Add($Button_Reset_Pass)
            $Form_Reset_Pass.Controls.Add($Button_Back)
            $Form_Reset_Pass.Controls.Add($Button_Exit)
            $Form_Reset_Pass.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Reset_Pass.ShowDialog()
        }
        <# Password reset function
        Функция сброса пароля #>
        function ResetPasswd{
            param(
                    [Parameter(Mandatory)]
                    $User
                    )
            process{
                $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName

                try{
                    $Text_New_Pass = New-Object System.Windows.Forms.Label
                    $Text_New_Pass.Text = "New password:"
                    $Text_New_Pass.Location = New-Object System.Drawing.Point(10,20)
                    $Text_New_Pass.AutoSize = $true

                    Add-Type -AssemblyName System.Web
                    $length = Get-Random -Minimum 9 -Maximum 12
                    $Generate = [System.Web.Security.Membership]::GeneratePassword($length, 2)
                    $NewPasswd = ConvertTo-SecureString $Generate -AsPlainText -force
                    Set-ADAccountPassword $UserLogin -NewPassword $NewPasswd -Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True
                    
                    $Color = "green"
                    $Outcome="$Generate"
                }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value

                    $Color = "red"
                    $Outcome = 'Execution failed. Contact your system administrator'
                }
                finally{
                    $TextBox_Success_Copy = New-Object System.Windows.Forms.TextBox
                    $TextBox_Success_Copy.Text = $Outcome
                    $TextBox_Success_Copy.ForeColor = $Color
                    $TextBox_Success_Copy.Location = New-Object System.Drawing.Point(10,40)
                    $TextBox_Success_Copy.AutoSize = $true
            
                    $Button_Exit = New-Object System.Windows.Forms.Button
                    $Button_Exit.Text = 'Exit'
                    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $Button_Exit.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true

                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_New_Pass.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_New_Pass = New-Object System.Windows.Forms.Form
                    $Form_New_Pass.Text ='Account Management Active Directory'
                    $Form_New_Pass.Width = 500
                    $Form_New_Pass.Height = 500
                    $Form_New_Pass.AutoSize = $true
                    $Form_New_Pass.Controls.Add($Text_New_Pass)
                    $Form_New_Pass.Controls.Add($TextBox_Success_Copy)
                    $Form_New_Pass.Controls.Add($Button_Back)
                    $Form_New_Pass.Controls.Add($Button_Exit)
                    $Form_New_Pass.Controls.Add($Label)

                    $Form_Reset_Pass.Hide()
                    $Form_New_Pass.ShowDialog()
                }
            }
        }
        <# User selection to unlock
        Выбор пользователя для разблокировки #>
        $Unlock_user = {
            $Text_Unlock_user = New-Object System.Windows.Forms.Label
            $Text_Unlock_user.Text = "Select user"
            $Text_Unlock_user.Location = New-Object System.Drawing.Point(10,20)
            $Text_Unlock_user.AutoSize = $true
            
            $ComboBox_Unlock_user = New-Object System.Windows.Forms.ComboBox
            $ComboBox_Unlock_user.Width = 250
            $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($User in $Users){
                $ComboBox_Unlock_user.Items.Add($User.cn);
                }
            $ComboBox_Unlock_user.Location = New-Object System.Drawing.Point(10,50)

            $Button_Unlock_user = New-Object System.Windows.Forms.Button
            $Button_Unlock_user.Location = New-Object System.Drawing.Size(300,50)
            $Button_Unlock_user.AutoSize = $true
            $Button_Unlock_user.Text = "Unblock user"
            $Button_Unlock_user.Add_Click({
                [string]$User = $ComboBox_Unlock_user.selectedItem
                UnlockAccount -User $User
                }
            )

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Unlock_user.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Unlock_user = New-Object System.Windows.Forms.Form
            $Form_Unlock_user.Text ='Account Management Active Directory'
            $Form_Unlock_user.Width = 500
            $Form_Unlock_user.Height = 500
            $Form_Unlock_user.AutoSize = $true
            $Form_Unlock_user.Controls.Add($Text_Unlock_user)
            $Form_Unlock_user.Controls.Add($ComboBox_Unlock_user)
            $Form_Unlock_user.Controls.Add($Button_Unlock_user)
            $Form_Unlock_user.Controls.Add($Button_Back)
            $Form_Unlock_user.Controls.Add($Button_Exit)
            $Form_Unlock_user.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Unlock_user.ShowDialog()
        }
        <# Function to unlock user
        Функция для разблокировки пользователя #>
        function UnlockAccount{
            param(
                [Parameter(Mandatory)]
                $User
                )
            process{
                $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName

                try{
                    $Text_Unlock = New-Object System.Windows.Forms.Label
                    $Text_Unlock.Text = "Result of checking:"
                    $Text_Unlock.Location = New-Object System.Drawing.Point(10,20)
                    $Text_Unlock.AutoSize = $true

                    $Color = "green"
                    $Check = Get-ADUser -Identity $UserLogin.SamAccountName -Properties LockedOut
                        if ($Check.LockedOut -eq $True) {
                            Unlock-ADAccount -Identity $UserLogin.SamAccountName
                            $Outcome="$User account unlocked"
                        }
                        else{
                            $Outcome="$User account is not locked out"
                        }
                }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value
                    $Color = "red"
                    $Outcome = 'Execution failed. Contact your system administrator'
                    }
                finally{
                    $Text_Unlock_1 = New-Object System.Windows.Forms.Label
                    $Text_Unlock_1.Text = $Outcome
                    $Text_Unlock_1.ForeColor = $Color
                    $Text_Unlock_1.Location = New-Object System.Drawing.Point(10,40)
                    $Text_Unlock_1.AutoSize = $true
        
                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_New_Pass.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_Unlock = New-Object System.Windows.Forms.Form
                    $Form_Unlock.Text ='Account Management Active Directory'
                    $Form_Unlock.Width = 500
                    $Form_Unlock.Height = 500
                    $Form_Unlock.AutoSize = $true
                    $Form_Unlock.Controls.Add($Text_Unlock)
                    $Form_Unlock.Controls.Add($Text_Unlock_1)
                    $Form_Unlock.Controls.Add($Button_Back)
                    $Form_Unlock.Controls.Add($Button_Exit)
                    $Form_Unlock.Controls.Add($Label)

                    $Form_Unlock_user.Hide()
                    $Form_Unlock.ShowDialog()
                }
            }
        }
        <# Selecting a user to disable an account
        Выбор пользователя для отключения учетной записи #>
        $Fired_user = {
            $Text_Fired_user = New-Object System.Windows.Forms.Label
            $Text_Fired_user.Text = "Select user"
            $Text_Fired_user.Location = New-Object System.Drawing.Point(10,20)
            $Text_Fired_user.AutoSize = $true
            
            $ComboBox_Fired_user = New-Object System.Windows.Forms.ComboBox
            $ComboBox_Fired_user.Width = 250
            $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($User in $Users){
                $ComboBox_Fired_user.Items.Add($User.cn);
                }
            $ComboBox_Fired_user.Location = New-Object System.Drawing.Point(10,50)

            $Button_Fired_user = New-Object System.Windows.Forms.Button
            $Button_Fired_user.Location = New-Object System.Drawing.Size(300,50)
            $Button_Fired_user.AutoSize = $true
            $Button_Fired_user.Text = "Carry out user dismissal"
            $Button_Fired_user.Add_Click({
                [string]$User = $ComboBox_Fired_user.selectedItem
                Fireduser -User $User
                }
            )

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Fired_user.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Fired_user = New-Object System.Windows.Forms.Form
            $Form_Fired_user.Text ='Account Management Active Directory'
            $Form_Fired_user.Width = 500
            $Form_Fired_user.Height = 500
            $Form_Fired_user.AutoSize = $true
            $Form_Fired_user.Controls.Add($Text_Fired_user)
            $Form_Fired_user.Controls.Add($ComboBox_Fired_user)
            $Form_Fired_user.Controls.Add($Button_Fired_user)
            $Form_Fired_user.Controls.Add($Button_Back)
            $Form_Fired_user.Controls.Add($Button_Exit)
            $Form_Fired_user.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Fired_user.ShowDialog()
        }
        <# Account disable feature
        Функция отключения учетной записи #>
        function Fired_User{
            param(
                [Parameter(Mandatory)]
                $User
                )
            process{
            $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName
                try{
                    $group_Member = Get-ADPrincipalGroupMembership -Identity $UserLogin | Where-Object { $_.Name -ne "Domain Users" }
                    Remove-ADPrincipalGroupMembership -Identity $UserLogin -MemberOf $group_Member -Confirm:$false
                    }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path Path_log -Value $Value
                    Write-Host "There are no access groups. Shutdown continues" -ForegroundColor Red
                    }
                finally{
                    Get-ADUser -Identity $UserLogin -server $DC | Move-ADObject -targetpath $Path_fired
                    Set-Mailbox -Identity $UserLogin -HiddenFromAddressListsEnabled $true
                    Set-ADUser -Identity $UserLogin -Enabled $false
                    Set-ADUser -Identity $UserLogin -Manager $null
            
                    Disable-mailbox -Identity $UserLogin -Confirm:$false
                    Get-GlobalAddressList | Update-GlobalAddressList
                    Get-OfflineAddressBook | Update-OfflineAddressBook
                    Get-AddressList | Update-AddressList
                    
                    $Color = "green"
                    $Outcome="$User account disabled"
                    
                    $Text_Fired = New-Object System.Windows.Forms.Label
                    $Text_Fired.Text = $Outcome
                    $Text_Fired.ForeColor = $Color
                    $Text_Fired.Location = New-Object System.Drawing.Point(10,40)
                    $Text_Fired.AutoSize = $true
        
                    $Button_Exit = New-Object System.Windows.Forms.Button
                    $Button_Exit.Text = 'Exit'
                    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
                    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $Button_Exit.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true

                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_Fired.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Form_Fired = New-Object System.Windows.Forms.Form
                    $Form_Fired.Text ='Account Management Active Directory'
                    $Form_Fired.Width = 500
                    $Form_Fired.Height = 500
                    $Form_Fired.AutoSize = $true
                    $Form_Fired.Controls.Add($Text_Fired)
                    $Form_Fired.Controls.Add($Button_Back)
                    $Form_Fired.Controls.Add($Button_Exit)
                    $Form_Fired.Controls.Add($Label)

                    $Form_Fired_user.Hide()
                    $Form_Fired.ShowDialog()
                    }
            }
        }
        <# User Selection to Disable VPN Access Groups
        Выбор пользователя для отключения групп доступа VPN #>
        $Remove_Group = {
            $Text_Remove_Group = New-Object System.Windows.Forms.Label
            $Text_Remove_Group.Text = "Select user"
            $Text_Remove_Group.Location = New-Object System.Drawing.Point(10,20)
            $Text_Remove_Group.AutoSize = $true
            
            $ComboBox_Remove_Group = New-Object System.Windows.Forms.ComboBox
            $ComboBox_Remove_Group.Width = 250
            $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($User in $Users){
                $ComboBox_Remove_Group.Items.Add($User.cn);
                }
            $ComboBox_Remove_Group.Location = New-Object System.Drawing.Point(10,50)

            $Button_Remove_Group = New-Object System.Windows.Forms.Button
            $Button_Remove_Group.Location = New-Object System.Drawing.Size(300,50)
            $Button_Remove_Group.AutoSize = $true
            $Button_Remove_Group.Text = "Disable access groups"
            $Button_Remove_Group.Add_Click({
                [string]$User = $ComboBox_Remove_Group.selectedItem
                RemoveGroup -User $User
                }
            )

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_Remove_Group.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_Remove_Group = New-Object System.Windows.Forms.Form
            $Form_Remove_Group.Text ='Account Management Active Directory'
            $Form_Remove_Group.Width = 500
            $Form_Remove_Group.Height = 500
            $Form_Remove_Group.AutoSize = $true
            $Form_Remove_Group.Controls.Add($Text_Remove_Group)
            $Form_Remove_Group.Controls.Add($ComboBox_Remove_Group)
            $Form_Remove_Group.Controls.Add($Button_Remove_Group)
            $Form_Remove_Group.Controls.Add($Button_Back)
            $Form_Remove_Group.Controls.Add($Button_Exit)
            $Form_Remove_Group.Controls.Add($Label)

            $ActiveDirectoryMenu.Hide()
            $Form_Remove_Group.ShowDialog()
        }
        <# Disable VPN access group feature
        Функция отключения групп доступа VPN #>
        function Removegroup{
            param(
                [Parameter(Mandatory)]
                $User
                )
            process{
                $UserLogin = Get-ADUser -Filter "cn -eq '$User'" -Properties SamAccountName

                try{
                    $Text_Removegroup = New-Object System.Windows.Forms.Label
                    $Text_Removegroup.Text = "Result of disabling access groups:"
                    $Text_Removegroup.Location = New-Object System.Drawing.Point(10,20)
                    $Text_Removegroup.AutoSize = $true

                    $group_Member = Get-ADPrincipalGroupMembership -Identity $UserLogin | Where-Object { $_.Name -ne "Domain Users" }
                    Remove-ADPrincipalGroupMembership -Identity $UserLogin -MemberOf $group_Member -Confirm:$false

                    $Color = "green"
                    $Outcome="$User access groups removed"
                    }
                catch {
                    $Data = Get-Date
                    $Exception = ($Error[0].Exception).Message
                    $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
                    $Value = "$Data" + " " + "$Exception"+"
                    " + "$InvocationInfo"
                    Add-Content -Path $Path_log -Value $Value

                    $Color = "red"
                    $Outcome = 'Execution failed. Contact your system administrator'
                    }
                finally{
                    $Text_Removegroup_1 = New-Object System.Windows.Forms.Label
                    $Text_Removegroup_1.Text = $Outcome
                    $Text_Removegroup_1.ForeColor = $Color
                    $Text_Removegroup_1.Location = New-Object System.Drawing.Point(10,40)
                    $Text_Removegroup_1.AutoSize = $true
        
                    $Button_Back = New-Object System.Windows.Forms.Button
                    $Button_Back.Text = 'Main menu'
                    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
                    $Button_Back.Add_click({$Form_New_Pass.Dispose(),$main_form.Show()})
                    $Button_Back.AutoSize = $true

                    $Label = New-Object System.Windows.Forms.Label
                    $Label.Text = $Label_Text
                    $Label.Location  = New-Object System.Drawing.Point(110,450)
                    $Label.AutoSize = $true        

                    $Form_Removegroup = New-Object System.Windows.Forms.Form
                    $Form_Removegroup.Text ='Account Management Active Directory'
                    $Form_Removegroup.Width = 500
                    $Form_Removegroup.Height = 500
                    $Form_Removegroup.AutoSize = $true
                    $Form_Removegroup.Controls.Add($Text_Removegroup)
                    $Form_Removegroup.Controls.Add($Text_Removegroup_1)
                    $Form_Removegroup.Controls.Add($Button_Back)
                    $Form_Removegroup.Controls.Add($Button_Exit)
                    $Form_Removegroup.Controls.Add($Label)

                    $Form_Remove_Group.Hide()
                    $Form_Removegroup.ShowDialog()
                }
            }
        }
        <# Cat show :)
        Показ котиков :) #>
        $For_Kuzia = {
            $PictureBox_For_Kuzia = New-Object System.Windows.Forms.PictureBox
            $PictureBox_For_Kuzia.Load('C:\cat.jpg')
            $PictureBox_For_Kuzia.Location = New-Object System.Drawing.Point(10,10)
            $PictureBox_For_Kuzia.AutoSize = $true

            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Button_Exit.AutoSize = $true

            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true

            $Button_Back = New-Object System.Windows.Forms.Button
            $Button_Back.Text = 'Back'
            $Button_Back.Location = New-Object System.Drawing.Point(10,400)
            $Button_Back.Add_click({$Form_For_Kuzia.Dispose(),$ActiveDirectoryMenu.Show()})
            $Button_Back.AutoSize = $true

            $Form_For_Kuzia = New-Object System.Windows.Forms.Form
            $Form_For_Kuzia.Text ='Kittens go!'
            $Form_For_Kuzia.Width = 500
            $Form_For_Kuzia.Height = 500
            $Form_For_Kuzia.AutoSize = $true
            $Form_For_Kuzia.Controls.Add($Label)
            $Form_For_Kuzia.Controls.add($PictureBox_For_Kuzia)
            $Form_For_Kuzia.Controls.Add($Button_Exit)
            $Form_For_Kuzia.Controls.Add($Button_Back)

            $ActiveDirectoryMenu.Hide()
            $Form_For_Kuzia.ShowDialog()
        }

        <# Button for further user creation
        Кнопка для дальнейшего создания пользователя #>
        $Button_Info_user = New-Object System.Windows.Forms.Button
        $Button_Info_user.Text = 'Create a new user'
        $Button_Info_user.Location = New-Object System.Drawing.Point(10,20)
        $Button_Info_user.Add_click($Ticket_Jira_User)
        $Button_Info_user.AutoSize = $true

        <# Button to further create an Outstaff user
        Кнопка для дальнейшего создания Outstaff-пользователя #>
        $Button_Info_user_outstaff = New-Object System.Windows.Forms.Button
        $Button_Info_user_outstaff.Text = 'Create a new Outstaff user'
        $Button_Info_user_outstaff.Location = New-Object System.Drawing.Point(10,50)
        $Button_Info_user_outstaff.Add_click($Ticket_Jira_Outstaff)
        $Button_Info_user_outstaff.AutoSize = $true

        <# Button for copying groups
        Кнопка для копирования групп #>
        $Button_Copy_group = New-Object System.Windows.Forms.Button
        $Button_Copy_group.Text = 'Copy groups from user to user'
        $Button_Copy_group.Location = New-Object System.Drawing.Point(10,80)
        $Button_Copy_group.Add_click($Copy_group)
        $Button_Copy_group.AutoSize = $true

        <# Password reset button
        Кнопка для сброса пароля #>
        $Button_Reset_Pass = New-Object System.Windows.Forms.Button
        $Button_Reset_Pass.Text = 'Reset account password'
        $Button_Reset_Pass.Location = New-Object System.Drawing.Point(10,110)
        $Button_Reset_Pass.Add_click($Reset_Pass)
        $Button_Reset_Pass.AutoSize = $true

        <# User unlock button
        Кнопка для разблокировки пользователя #>
        $Button_Unlock_user = New-Object System.Windows.Forms.Button
        $Button_Unlock_user.Text = 'Unlock account'
        $Button_Unlock_user.Location = New-Object System.Drawing.Point(10,140)
        $Button_Unlock_user.Add_click($Unlock_user)
        $Button_Unlock_user.AutoSize = $true

        <# Button for dismissing a user
        Кнопка для увольнения пользователя #>
        $Button_Fired_user = New-Object System.Windows.Forms.Button
        $Button_Fired_user.Text = 'Dismissing an employee'
        $Button_Fired_user.Location = New-Object System.Drawing.Point(10,170)
        $Button_Fired_user.Add_click($Fired_user)
        $Button_Fired_user.AutoSize = $true

        <# Button for deleting access groups
        Кнопка для удаления групп доступа #>
        $Button_Remove_group = New-Object System.Windows.Forms.Button
        $Button_Remove_group.Text = 'Removing access groups (access to mail will remain)'
        $Button_Remove_group.Location = New-Object System.Drawing.Point(10,200)
        $Button_Remove_group.Add_click($Remove_group)
        $Button_Remove_group.AutoSize = $true

        <# Button for cats
        Кнопка для котиков #>
        $Button_For_Kuzia = New-Object System.Windows.Forms.Button
        $Button_For_Kuzia.Text = 'Cat'
        $Button_For_Kuzia.Location = New-Object System.Drawing.Point(10,230)
        $Button_For_Kuzia.Add_click($For_Kuzia)
        $Button_For_Kuzia.AutoSize = $true

        <# Button to return to the main menu
        Кнопка для возврата в главное меню #>
        $Button_Main_Menu = New-Object System.Windows.Forms.Button
        $Button_Main_Menu.Text = 'Main menu'
        $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
        $Button_Main_Menu.Add_click({$ActiveDirectoryMenu.Dispose(),$main_form.Show()})
        $Button_Main_Menu.AutoSize = $true

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $Label_Text
        $Label.Location  = New-Object System.Drawing.Point(110,450)
        $Label.AutoSize = $true

        <# Creating the main menu form
        Создание формы главного меню #>
        $ActiveDirectoryMenu = New-Object System.Windows.Forms.Form
        $ActiveDirectoryMenu.Text ='Account Management Active Directory'
        $ActiveDirectoryMenu.Width = 500
        $ActiveDirectoryMenu.Height = 500
        $ActiveDirectoryMenu.AutoSize = $true
        $ActiveDirectoryMenu.Controls.Add($Button_Info_user)
        $ActiveDirectoryMenu.Controls.Add($Button_Info_user_outstaff)
        $ActiveDirectoryMenu.Controls.Add($Button_Copy_group)
        $ActiveDirectoryMenu.Controls.Add($Button_Reset_Pass)
        $ActiveDirectoryMenu.Controls.Add($Button_Unlock_user)
        $ActiveDirectoryMenu.Controls.Add($Button_Fired_user)
        $ActiveDirectoryMenu.Controls.Add($Button_Remove_group)
        $ActiveDirectoryMenu.Controls.Add($Button_For_Kuzia)
        $ActiveDirectoryMenu.Controls.Add($Button_Main_Menu)
        $ActiveDirectoryMenu.Controls.Add($Button_Exit)
        $ActiveDirectoryMenu.Controls.Add($Label)

        $main_form.Hide()
        $ActiveDirectoryMenu.ShowDialog()
    }

    <# Menu Exchange
    Меню Exchange #>
    $ExchangeMenu = {
        <# Button to create a mailbox
        Кнопка для создания почтового ящика #>
        $Button_Create_mail = New-Object System.Windows.Forms.Button
        $Button_Create_mail.Text = 'Create a mailbox'
        $Button_Create_mail.Location = New-Object System.Drawing.Point(10,20)
        #$Create_mail.Add_click($Create_mail)
        $Button_Create_mail.AutoSize = $true

        <# Button for granting rights "Send as"
        Кнопка для предоставления прав "Отправить как" #>
        $Button_Send_as = New-Object System.Windows.Forms.Button
        $Button_Send_as.Text = 'Grant "Send As" rights'
        $Button_Send_as.Location = New-Object System.Drawing.Point(10,50)
        #$Send_as.Add_click($Send_as)
        $Button_Send_as.AutoSize = $true

        <# Button for granting "Full Control" rights
        Кнопка для предоставления прав "Полный доступ" #>
        $Button_Full_access = New-Object System.Windows.Forms.Button
        $Button_Full_access.Text = 'Granting "Full Control" rights'
        $Button_Full_access.Location = New-Object System.Drawing.Point(10,80)
        #$Full_access.Add_click($Full_access)
        $Button_Full_access.AutoSize = $true

        <# Button for connecting an Outstaff mailbox
        Кнопка для подклчения почтового ящика Outstaff #>
        $Button_Mail_for_Outstaff = New-Object System.Windows.Forms.Button
        $Button_Mail_for_Outstaff.Text = 'Connecting a mailbox for Outstaff'
        $Button_Mail_for_Outstaff.Location = New-Object System.Drawing.Point(10,140)
        #$Mail_for_Outstaff.Add_click($Mail_for_Outstaff)
        $Button_Mail_for_Outstaff.AutoSize = $true

        <# Button to return to the main menu
        Кнопка для возврата в главное меню #>
        $Button_Main_Menu = New-Object System.Windows.Forms.Button
        $Button_Main_Menu.Text = 'Main menu'
        $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
        $Button_Main_Menu.Add_click({$ExchangeMenu.Dispose(),$main_form.Show()})
        $Button_Main_Menu.AutoSize = $true

        <# Button to close the form
        Кнопка для закрытия формы #>
        $Button_Exit = New-Object System.Windows.Forms.Button
        $Button_Exit.Text = 'Exit'
        $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
        $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Button_Exit.AutoSize = $true

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $Label_Text
        $Label.Location  = New-Object System.Drawing.Point(110,450)
        $Label.AutoSize = $true

        $ExchangeMenu = New-Object System.Windows.Forms.Form
        $ExchangeMenu.Text ='Exchange account management'
        $ExchangeMenu.Width = 500
        $ExchangeMenu.Height = 500
        $ExchangeMenu.AutoSize = $true
        $ExchangeMenu.Controls.Add($Button_Create_mail)
        $ExchangeMenu.Controls.Add($Button_Send_as)
        $ExchangeMenu.Controls.Add($Button_Full_access)
        $ExchangeMenu.Controls.Add($Button_Mail_for_Outstaff)
        $ExchangeMenu.Controls.Add($Button_Main_Menu)
        $ExchangeMenu.Controls.Add($Button_Exit)
        $ActiveDirectoryMenu.Controls.Add($Label)

        $main_form.Hide()
        $ExchangeMenu.ShowDialog()
    }

    <# Help Menu
    Меню справочной информации #>
    $InfoMenu = {

        $Info_Pass = {
            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = $Label_Text
            $Label.Location  = New-Object System.Drawing.Point(110,450)
            $Label.AutoSize = $true
    
            $Text_Info_Pass = New-Object System.Windows.Forms.Label
            $Text_Info_Pass.Text = "Select user"
            $Text_Info_Pass.Location = New-Object System.Drawing.Point(10,20)
            $Text_Info_Pass.AutoSize = $true
            $Info_Pass.Controls.Add($Text_Info_Pass)
    
            $ComboBox_Info_Pass = New-Object System.Windows.Forms.ComboBox
            $ComboBox_Info_Pass.Width = 250
            $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
            Foreach ($User in $Users)
            {
            $ComboBox_Info_Pass.Items.Add($User.cn);
            }
            $ComboBox_Info_Pass.Location = New-Object System.Drawing.Point(10,50)
            $Info_Pass.Controls.Add($ComboBox_Info_Pass)
    
            $Text_Info_Pass_2 = New-Object System.Windows.Forms.Label
            $Text_Info_Pass_2.Text = "Last password change:"
            $Text_Info_Pass_2.Location = New-Object System.Drawing.Point(10,80)
            $Text_Info_Pass_2.AutoSize = $true
            $Info_Pass.Controls.Add($Text_Info_Pass_2)
    
            $Text_Info_Pass_3 = New-Object System.Windows.Forms.Label
            $Text_Info_Pass_3.Text = ""
            $Text_Info_Pass_3.Location = New-Object System.Drawing.Point(10,110)
            $Text_Info_Pass_3.AutoSize = $true
            $Info_Pass.Controls.Add($Text_Info_Pass_3)
    
            $Button_Check_Pass = New-Object System.Windows.Forms.Button
            $Button_Check_Pass.Location = New-Object System.Drawing.Size(400,20)
            $Button_Check_Pass.AutoSize = $true
            $Button_Check_Pass.Text = "Check"
            $Info_Pass.Controls.Add($Button_Check_Pass)
            $Button_Check_Pass.Add_Click(
                {
                [string]$Name = $ComboBox_Info_Pass.selectedItem
                $Text_Info_Pass_3.Text = [datetime]::FromFileTime((Get-ADUser -Filter "cn -eq $Name" -Properties pwdLastSet).pwdLastSet).ToString('dd.MM.yyyy hh:mm')
                }
            )
    
            <# Button to return to the main menu
            Кнопка для возврата в главное меню #>
            $Button_Main_Menu = New-Object System.Windows.Forms.Button
            $Button_Main_Menu.Text = 'Main menu'
            $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
            $Button_Main_Menu.Add_click({$Info_Pass.Dispose(),$main_form.Show()})
            $Info_Pass.Controls.Add($Button_Main_Menu)
            $Button_Main_Menu.AutoSize = $true
    
            <# Button to close the form
            Кнопка для закрытия формы #>
            $Button_Exit = New-Object System.Windows.Forms.Button
            $Button_Exit.Text = 'Exit'
            $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
            $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $Info_Pass.Controls.Add($Button_Exit)
            $Button_Exit.AutoSize = $true
    
            $Info_Pass = New-Object System.Windows.Forms.Form
            $Info_Pass.Text ='Reference Information'
            $Info_Pass.Width = 500
            $Info_Pass.Height = 500
            $Info_Pass.AutoSize = $true
            $Info_Pass.Controls.Add($Label)
    
            $InfoMenu.Dispose()
            $Info_Pass.ShowDialog()
        }
    
        <# Button to view when the password was last changed
        Кнопка для просмотра когда последний раз меняли пароль #>
        $Button_Info_Pass = New-Object System.Windows.Forms.Button
        $Button_Info_Pass.Text = 'See when the password was last changed'
        $Button_Info_Pass.Location = New-Object System.Drawing.Point(10,20)
        $Button_Info_Pass.Add_click($Info_Pass)
        $Button_Info_Pass.AutoSize = $true

        <# Getting a list of PCs that were last online 40+ days ago
        Получение списка ПК которые были последний раз в сети 40+ дней назад #>
        $Button_PC_old = New-Object System.Windows.Forms.Button
        $Button_PC_old.Text = 'View PCs that have not been online for a long time'
        $Button_PC_old.Location = New-Object System.Drawing.Point(10,50)
        #$Button_PC_old.Add_click($PC_old)
        $Button_PC_old.AutoSize = $true

        <# View account attributes
        Посмотреть атрибуты учетной записи #>
        $Button_Attributes = New-Object System.Windows.Forms.Button
        $Button_Attributes.Text = 'View account attributes'
        $Button_Attributes.Location = New-Object System.Drawing.Point(10,80)
        #$Button_Attributes.Add_click($Acc_Attributes)
        $Button_Attributes.AutoSize = $true

        <# Button to return to the main menu
        Кнопка для возврата в главное меню #>
        $Button_Main_Menu = New-Object System.Windows.Forms.Button
        $Button_Main_Menu.Text = 'Main menu'
        $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
        $Button_Main_Menu.Add_click({$InfoMenu.Dispose(),$main_form.Show()})
        $Button_Main_Menu.AutoSize = $true

        <# Button to close the form
        Кнопка для закрытия формы #>
        $Button_Exit = New-Object System.Windows.Forms.Button
        $Button_Exit.Text = 'Exit'
        $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
        $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Button_Exit.AutoSize = $true

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $Label_Text
        $Label.Location  = New-Object System.Drawing.Point(110,450)
        $Label.AutoSize = $true

        $InfoMenu = New-Object System.Windows.Forms.Form
        $InfoMenu.Text ='Reference Information'
        $InfoMenu.Width = 500
        $InfoMenu.Height = 500
        $InfoMenu.AutoSize = $true
        $InfoMenu.Controls.Add($Button_Info_Pass)
        $InfoMenu.Controls.Add($Button_PC_old)
        $InfoMenu.Controls.Add($Button_Attributes)
        $InfoMenu.Controls.Add($Button_Main_Menu)
        $InfoMenu.Controls.Add($Button_Exit)
        $InfoMenu.Controls.Add($Label)

        $main_form.Hide()
        $InfoMenu.ShowDialog()
        }

    <# Button to go to the Active Directory menu
    Кнопка перехода в меню Active Directory #>
    $global:Button_ADMenu = New-Object System.Windows.Forms.Button
    $global:Button_ADMenu.Text = 'Active Directory Account Management'
    $global:Button_ADMenu.Location = New-Object System.Drawing.Point(10,20)
    $global:Button_ADMenu.Add_click($ActiveDirectoryMenu)
    $global:Button_ADMenu.AutoSize = $true
    
    <# Exchange menu button
    Кнопка перехода в меню Exchange #>
    $global:Button_ExchMenu = New-Object System.Windows.Forms.Button
    $global:Button_ExchMenu.Text = 'Exchange account management'
    $global:Button_ExchMenu.Location = New-Object System.Drawing.Point(10,50)
    $global:Button_ExchMenu.Add_click($ExchangeMenu)
    $global:Button_ExchMenu.AutoSize = $true

    <# Button to go to the help menu
    Кнопка для перехода в меню справочной информации #>
    $global:Button_Info = New-Object System.Windows.Forms.Button
    $global:Button_Info.Text = 'Reference Information'
    $global:Button_Info.Location = New-Object System.Drawing.Point(10,80)
    $global:Button_Info.Add_click($InfoMenu)
    $global:Button_Info.AutoSize = $true

    <# Button to close the form
    Кнопка для закрытия формы #>
    $global:Button_Exit = New-Object System.Windows.Forms.Button
    $global:Button_Exit.Text = 'Exit'
    $global:Button_Exit.Location = New-Object System.Drawing.Point(400,400)
    $global:Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $global:Button_Exit.AutoSize = $true

    <# Signature at the bottom of the form
    Подпись снизу формы #>
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Label_Text
    $Label.Location  = New-Object System.Drawing.Point(110,450)
    $Label.AutoSize = $true

    <# Main menu form
    Форма главного меню #>
    $global:main_form = New-Object System.Windows.Forms.Form
    $global:main_form.Text ='Account Management'
    $global:main_form.Width = 500
    $global:main_form.Height = 500
    $global:main_form.AutoSize = $true
    $global:main_form.Controls.Add($Button_ADMenu)
    $global:main_form.Controls.Add($Button_ExchMenu)
    $global:main_form.Controls.Add($Button_Info)
    $global:main_form.Controls.Add($Button_Exit)
    $global:main_form.Controls.Add($Label)

    $global:main_form.ShowDialog()
}

Manage-user