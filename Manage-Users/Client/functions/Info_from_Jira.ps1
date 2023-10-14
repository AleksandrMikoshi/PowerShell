<#Menu function showing data from Jira to create an account for an collaborator
Under the asterisk (*) are required fields for further filling#>
<#Функция меню с отображением данных из Jira для создания учетной записи для сотрудника
Под звездочкой (*) обязательные поля для дальнейшего заполнения#>
function Info_from_Jira {
    param (
        [Parameter(Mandatory)]
        $Data
    )   
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $Label_Info_User = New-Object System.Windows.Forms.Label
    $Label_Info_User.Text = "Check data"
    $Label_Info_User.Location = New-Object System.Drawing.Point(10,20)
    $Label_Info_User.AutoSize = $true

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Application number*"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,50)
    $Label_ticket_1.AutoSize = $true

    $TextBox_Task = New-Object System.Windows.Forms.TextBox
    $TextBox_Task.Location  = New-Object System.Drawing.Point(150,50)
    $TextBox_Task.Text = $Data.Task
    $TextBox_Task.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "Last Name*"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,75)
    $Label_ticket_2.AutoSize = $true

    $TextBox_LastName = New-Object System.Windows.Forms.TextBox
    $TextBox_LastName.Location  = New-Object System.Drawing.Point(150,75)
    $TextBox_LastName.Text = $Data.LastName
    $TextBox_LastName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_3 = New-Object System.Windows.Forms.Label
    $Label_ticket_3.Text = "First Name*"
    $Label_ticket_3.Location = New-Object System.Drawing.Point(10,100)
    $Label_ticket_3.AutoSize = $true

    $TextBox_FirstName = New-Object System.Windows.Forms.TextBox
    $TextBox_FirstName.Location  = New-Object System.Drawing.Point(150,100)
    $TextBox_FirstName.Text = $Data.FirstName
    $TextBox_FirstName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_4 = New-Object System.Windows.Forms.Label
    $Label_ticket_4.Text = "Surname*"
    $Label_ticket_4.Location = New-Object System.Drawing.Point(10,125)
    $Label_ticket_4.AutoSize = $true

    $TextBox_Initials = New-Object System.Windows.Forms.TextBox
    $TextBox_Initials.Location  = New-Object System.Drawing.Point(150,125)
    $TextBox_Initials.Text = $Data.Initials
    $TextBox_Initials.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_5 = New-Object System.Windows.Forms.Label
    $Label_ticket_5.Text = "Login*"
    $Label_ticket_5.Location = New-Object System.Drawing.Point(10,150)
    $Label_ticket_5.AutoSize = $true

    $TextBox_SamAccountName = New-Object System.Windows.Forms.TextBox
    $TextBox_SamAccountName.Location  = New-Object System.Drawing.Point(150,150)
    $TextBox_SamAccountName.Text = $Data.SamAccountName
    $TextBox_SamAccountName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_6 = New-Object System.Windows.Forms.Label
    $Label_ticket_6.Text = "Office*"
    $Label_ticket_6.Location = New-Object System.Drawing.Point(10,175)
    $Label_ticket_6.AutoSize = $true

    $TextBox_Office = New-Object System.Windows.Forms.TextBox
    $TextBox_Office.Location  = New-Object System.Drawing.Point(150,175)
    $TextBox_Office.Text = $Data.Office
    $TextBox_Office.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_7 = New-Object System.Windows.Forms.Label
    $Label_ticket_7.Text = "Department*"
    $Label_ticket_7.Location = New-Object System.Drawing.Point(10,200)
    $Label_ticket_7.AutoSize = $true

    $TextBox_Department = New-Object System.Windows.Forms.TextBox
    $TextBox_Department.Location  = New-Object System.Drawing.Point(150,200)
    $TextBox_Department.Text = $Data.Department
    $TextBox_Department.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_8 = New-Object System.Windows.Forms.Label
    $Label_ticket_8.Text = "Management*"
    $Label_ticket_8.Location = New-Object System.Drawing.Point(10,225)
    $Label_ticket_8.AutoSize = $true

    $TextBox_Management = New-Object System.Windows.Forms.TextBox
    $TextBox_Management.Location  = New-Object System.Drawing.Point(150,225)
    $TextBox_Management.Text = $Data.Management
    $TextBox_Management.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_9 = New-Object System.Windows.Forms.Label
    $Label_ticket_9.Text = "Division"
    $Label_ticket_9.Location = New-Object System.Drawing.Point(10,250)
    $Label_ticket_9.AutoSize = $true

    $TextBox_Division = New-Object System.Windows.Forms.TextBox
    $TextBox_Division.Location  = New-Object System.Drawing.Point(150,250)
    $TextBox_Division.Text = $Data.Division
    $TextBox_Division.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_10 = New-Object System.Windows.Forms.Label
    $Label_ticket_10.Text = "Job Title*"
    $Label_ticket_10.Location = New-Object System.Drawing.Point(10,275)
    $Label_ticket_10.AutoSize = $true

    $TextBox_JobTitle = New-Object System.Windows.Forms.TextBox
    $TextBox_JobTitle.Location  = New-Object System.Drawing.Point(150,275)
    $TextBox_JobTitle.Text = $Data.JobTitle
    $TextBox_JobTitle.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_11 = New-Object System.Windows.Forms.Label
    $Label_ticket_11.Text = "Manager*"
    $Label_ticket_11.Location = New-Object System.Drawing.Point(10,300)
    $Label_ticket_11.AutoSize = $true

    $TextBox_Manager = New-Object System.Windows.Forms.TextBox
    $TextBox_Manager.Location  = New-Object System.Drawing.Point(150,300)
    $TextBox_Manager.Text = $Data.Manager
    $TextBox_Manager.Size = New-Object System.Drawing.Size(300,30)
            
    $Label_ticket_12 = New-Object System.Windows.Forms.Label
    $Label_ticket_12.Text = "Phone*"
    $Label_ticket_12.Location = New-Object System.Drawing.Point(10,325)
    $Label_ticket_12.AutoSize = $true

    $TextBox_Phone = New-Object System.Windows.Forms.TextBox
    $TextBox_Phone.Location  = New-Object System.Drawing.Point(150,325)
    $TextBox_Phone.Text = $Data.Phone
    $TextBox_Phone.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_13 = New-Object System.Windows.Forms.Label
    $Label_ticket_13.Text = "Date of Birth*"
    $Label_ticket_13.Location = New-Object System.Drawing.Point(10,350)
    $Label_ticket_13.AutoSize = $true

    $TextBox_Birth = New-Object System.Windows.Forms.TextBox
    $TextBox_Birth.Location  = New-Object System.Drawing.Point(150,350)
    $TextBox_Birth.Text = $Data.Birth
    $TextBox_Birth.Size = New-Object System.Drawing.Size(300,30)

    $Button_Info_User = New-Object System.Windows.Forms.Button
    $Button_Info_User.Location = New-Object System.Drawing.Size(330,20)
    $Button_Info_User.AutoSize = $true
    $Button_Info_User.Text = "Create an account"
    $Button_Info_User.Add_click(
        {
            $UserAttr = @{
            LastName        = $TextBox_LastName.Text
            FirstName       = $TextBox_FirstName.Text
            Initials        = $TextBox_Initials.Text
            SamAccountName  = $TextBox_SamAccountName.Text
            Office          = $TextBox_Office.Text
            JobTitle        = $TextBox_JobTitle.Text
            Department      = $TextBox_Department.Text
            Management      = $TextBox_Management.Text
            Phone           = $TextBox_Phone.Text
            Birth           = $TextBox_Birth.Text
            Manager         = $TextBox_Manager.Text
            Division        = $TextBox_Division.Text
            PipeLine        = $TextBox_Task.Text
        }
        $Form_Info_User.Controls.Remove($Label_Info_User)
        $Form_Info_User.Controls.Remove($Label_ticket_1)
        $Form_Info_User.Controls.Remove($TextBox_Task)
        $Form_Info_User.Controls.Remove($Label_ticket_2)
        $Form_Info_User.Controls.Remove($TextBox_LastName)
        $Form_Info_User.Controls.Remove($Label_ticket_3)
        $Form_Info_User.Controls.Remove($TextBox_FirstName)
        $Form_Info_User.Controls.Remove($Label_ticket_4)
        $Form_Info_User.Controls.Remove($TextBox_Initials)
        $Form_Info_User.Controls.Remove($Label_ticket_5)
        $Form_Info_User.Controls.Remove($TextBox_SamAccountName)
        $Form_Info_User.Controls.Remove($Label_ticket_6)
        $Form_Info_User.Controls.Remove($TextBox_Office)
        $Form_Info_User.Controls.Remove($Label_ticket_7)
        $Form_Info_User.Controls.Remove($TextBox_Department)
        $Form_Info_User.Controls.Remove($Label_ticket_8)
        $Form_Info_User.Controls.Remove($TextBox_Management)
        $Form_Info_User.Controls.Remove($Label_ticket_9)
        $Form_Info_User.Controls.Remove($TextBox_Division)
        $Form_Info_User.Controls.Remove($Label_ticket_10)
        $Form_Info_User.Controls.Remove($TextBox_JobTitle)
        $Form_Info_User.Controls.Remove($Label_ticket_11)
        $Form_Info_User.Controls.Remove($TextBox_Manager)
        $Form_Info_User.Controls.Remove($Label_ticket_12)
        $Form_Info_User.Controls.Remove($TextBox_Phone)
        $Form_Info_User.Controls.Remove($Label_ticket_13)
        $Form_Info_User.Controls.Remove($TextBox_Birth)
        $Form_Info_User.Controls.Remove($Button_Info_User)

        $Text_Create_User = New-Object System.Windows.Forms.Label
        $Text_Create_User.Text = "- Create an account"
        $Text_Create_User.Location = New-Object System.Drawing.Point(10,20)
        $Text_Create_User.AutoSize = $true
        $Text_Create_User.BackColor = $Label_Color
        $Text_Create_User.ForeColor = $Text_Label
        $Form_Info_User.Controls.Add($Text_Create_User)
            
        $Text_Add_Mail = New-Object System.Windows.Forms.Label
        $Text_Add_Mail.Text = "- Creating a mailbox"
        $Text_Add_Mail.Location = New-Object System.Drawing.Point(10,40)
        $Text_Add_Mail.AutoSize = $true
        $Text_Add_Mail.BackColor = $Label_Color
        $Text_Add_Mail.ForeColor = $Text_Label
        $Form_Info_User.Controls.Add($Text_Add_Mail)

        $Text_Add_Group = New-Object System.Windows.Forms.Label
        $Text_Add_Group.Text = "- Adding groups to Active Directory"
        $Text_Add_Group.Location = New-Object System.Drawing.Point(10,60)
        $Text_Add_Group.AutoSize = $true
        $Text_Add_Group.BackColor = $Label_Color
        $Text_Add_Group.ForeColor = $Text_Label
        $Form_Info_User.Controls.Add($Text_Add_Group)

        $Create = Create_User -UserAttr $UserAttr -DC $DC -Path_OU $Path_OU -Path_log $Path_log
        $Outcome_User = $Create.Outcome
        [string]$Color_User = $Create.Color
        $Text_Create_User.Text = $Outcome_User
        $Text_Create_User.ForeColor = $Color_User

        sleep 1

        $CreateUserMail = Create_Mail_Jea -UserAttr $UserAttr -DC $DC -Mail_serv $Mail_serv -Path_log $Path_log -DB_Mail $DB_Mail -ArchiveDatabase $ArchiveDatabase
        $Outcome_User_Mail = $CreateUserMail.Outcome
        [string]$Color_User_Mail = $CreateUserMail.Color
        $Text_Add_Mail.Text = $Outcome_User_Mail
        $Text_Add_Mail.ForeColor = $Color_User_Mail

        sleep 1

        $AddGroup = Add_Group -UserAttr $UserAttr -DC $DC -Path_log $Path_log
        $Outcome_Group = $AddGroup.Outcome
        [string]$Color_Group = $AddGroup.Color
        $Text_Add_Group.Text = $Outcome_Group
        $Text_Add_Group.ForeColor = $Color_Group

        }
    )
             
    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$Back = Ticket_Jira_User $Form_Info_User.Dispose(),$Back})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Info_User = New-Object System.Windows.Forms.Form
    $Form_Info_User.Icon = $Icon
    $Form_Info_User.BackgroundImage = $Background
    $Form_Info_User.BackgroundImageLayout = "None"
    $Form_Info_User.Text ='Active Directory Account Management'
    $Form_Info_User.Width = 500
    $Form_Info_User.Height = 500
    $Form_Info_User.AutoSize = $true
    $Form_Info_User.Controls.Add($Label_Info_User)
    $Form_Info_User.Controls.Add($Label_ticket_1)
    $Form_Info_User.Controls.Add($TextBox_Task)
    $Form_Info_User.Controls.Add($Label_ticket_2)
    $Form_Info_User.Controls.Add($TextBox_LastName)
    $Form_Info_User.Controls.Add($Label_ticket_3)
    $Form_Info_User.Controls.Add($TextBox_FirstName)
    $Form_Info_User.Controls.Add($Label_ticket_4)
    $Form_Info_User.Controls.Add($TextBox_Initials)
    $Form_Info_User.Controls.Add($Label_ticket_5)
    $Form_Info_User.Controls.Add($TextBox_SamAccountName)
    $Form_Info_User.Controls.Add($Label_ticket_6)
    $Form_Info_User.Controls.Add($TextBox_Office)
    $Form_Info_User.Controls.Add($Label_ticket_7)
    $Form_Info_User.Controls.Add($TextBox_Department)
    $Form_Info_User.Controls.Add($Label_ticket_8)
    $Form_Info_User.Controls.Add($TextBox_Management)
    $Form_Info_User.Controls.Add($Label_ticket_9)
    $Form_Info_User.Controls.Add($TextBox_Division)
    $Form_Info_User.Controls.Add($Label_ticket_10)
    $Form_Info_User.Controls.Add($TextBox_JobTitle)
    $Form_Info_User.Controls.Add($Label_ticket_11)
    $Form_Info_User.Controls.Add($TextBox_Manager)
    $Form_Info_User.Controls.Add($Label_ticket_12)
    $Form_Info_User.Controls.Add($TextBox_Phone)
    $Form_Info_User.Controls.Add($Label_ticket_13)
    $Form_Info_User.Controls.Add($TextBox_Birth)
    $Form_Info_User.Controls.Add($Button_Info_User)
    $Form_Info_User.Controls.Add($Button_Back)
    $Form_Info_User.Controls.Add($Button_Exit)
    $Form_Info_User.Controls.Add($Label)
    $Form_Info_User.StartPosition = 'CenterScreen'
    $Form_Info_User.MaximizeBox = $false
    $Form_Info_User.FormBorderStyle = "FixedDialog"

    $Form_Ticket_Jira.Dispose()
    $Form_Info_User.ShowDialog()
}