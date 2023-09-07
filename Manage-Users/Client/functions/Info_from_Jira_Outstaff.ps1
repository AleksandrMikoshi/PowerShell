<#Menu function showing data from Jira to create an account for an external collaborator#>
<#Функция меню с отображением данных из Jira для создания учетной записи для внешнего сотрудника#>
function Info_from_Jira_Outstaff {
    param (
        [Parameter(Mandatory)]
        $Data
    )   
    Add-Type -assembly System.Windows.Forms

    $Text_Outstaff_Info_User = New-Object System.Windows.Forms.Label
    $Text_Outstaff_Info_User.Text = "Check the data"
    $Text_Outstaff_Info_User.Location = New-Object System.Drawing.Point(10,20)
    $Text_Outstaff_Info_User.AutoSize = $true

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Last Name"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,50)
    $Label_ticket_1.AutoSize = $true

    $TextBox_LastName = New-Object System.Windows.Forms.TextBox
    $TextBox_LastName.Location  = New-Object System.Drawing.Point(150,50)
    $TextBox_LastName.Text = $Data.LastName
    $TextBox_LastName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "First Name"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,75)
    $Label_ticket_2.AutoSize = $true

    $TextBox_FirstName = New-Object System.Windows.Forms.TextBox
    $TextBox_FirstName.Location  = New-Object System.Drawing.Point(150,75)
    $TextBox_FirstName.Text = $Data.FirstName
    $TextBox_FirstName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_3 = New-Object System.Windows.Forms.Label
    $Label_ticket_3.Text = "Surname"
    $Label_ticket_3.Location = New-Object System.Drawing.Point(10,100)
    $Label_ticket_3.AutoSize = $true

    $TextBox_Initials = New-Object System.Windows.Forms.TextBox
    $TextBox_Initials.Location  = New-Object System.Drawing.Point(150,100)
    $TextBox_Initials.Text = $Data.Initials
    $TextBox_Initials.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_4 = New-Object System.Windows.Forms.Label
    $Label_ticket_4.Text = "Login"
    $Label_ticket_4.Location = New-Object System.Drawing.Point(10,125)
    $Label_ticket_4.AutoSize = $true

    $TextBox_SamAccountName = New-Object System.Windows.Forms.TextBox
    $TextBox_SamAccountName.Location  = New-Object System.Drawing.Point(150,125)
    $TextBox_SamAccountName.Text = $Data.SamAccountName
    $TextBox_SamAccountName.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_5 = New-Object System.Windows.Forms.Label
    $Label_ticket_5.Text = "Office"
    $Label_ticket_5.Location = New-Object System.Drawing.Point(10,150)
    $Label_ticket_5.AutoSize = $true

    $TextBox_Office = New-Object System.Windows.Forms.TextBox
    $TextBox_Office.Location  = New-Object System.Drawing.Point(150,150)
    $TextBox_Office.Text = $Data.Office
    $TextBox_Office.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_6 = New-Object System.Windows.Forms.Label
    $Label_ticket_6.Text = "Department"
    $Label_ticket_6.Location = New-Object System.Drawing.Point(10,175)
    $Label_ticket_6.AutoSize = $true

    $TextBox_Department = New-Object System.Windows.Forms.TextBox
    $TextBox_Department.Location  = New-Object System.Drawing.Point(150,175)
    $TextBox_Department.Text = $Data.Department
    $TextBox_Department.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_7 = New-Object System.Windows.Forms.Label
    $Label_ticket_7.Text = "Phone"
    $Label_ticket_7.Location = New-Object System.Drawing.Point(10,200)
    $Label_ticket_7.AutoSize = $true

    $TextBox_Phone = New-Object System.Windows.Forms.TextBox
    $TextBox_Phone.Location  = New-Object System.Drawing.Point(150,200)
    $TextBox_Phone.Text = $Data.Phone
    $TextBox_Phone.Size = New-Object System.Drawing.Size(300,30)

    $Label_ticket_8 = New-Object System.Windows.Forms.Label
    $Label_ticket_8.Text = "Company"
    $Label_ticket_8.Location = New-Object System.Drawing.Point(10,225)
    $Label_ticket_8.AutoSize = $true
            
    $TextBox_Company = New-Object System.Windows.Forms.TextBox
    $TextBox_Company.Location  = New-Object System.Drawing.Point(150,225)
    $TextBox_Company.Text = $Data.Company
    $TextBox_Company.Size = New-Object System.Drawing.Size(300,30)

    $Button_Outstaff_Create_User = New-Object System.Windows.Forms.Button
    $Button_Outstaff_Create_User.Location = New-Object System.Drawing.Size(330,20)
    $Button_Outstaff_Create_User.AutoSize = $true
    $Button_Outstaff_Create_User.Text = "Create an account"
    $Button_Outstaff_Create_User.Add_click(
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
            $Create_Outstaff = Create_Outstaff_User -UserAttr $UserAttr -DC $DC -Path_OU_Outstaff $Path_OU_Outstaff -Path_log $Path_log

            [string]$Color = $Create_Outstaff.Color
            [string]$Outcome = $Create_Outstaff.Outcome
            $Form_Outstaff_Info_User.Controls.Remove($Text_Outstaff_Info_User)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_1)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_LastName)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_2)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_FirstName)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_3)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_Initials)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_4)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_SamAccountName)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_5)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_Office)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_6)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_Department)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_7)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_Phone)
            $Form_Outstaff_Info_User.Controls.Remove($Label_ticket_8)
            $Form_Outstaff_Info_User.Controls.Remove($TextBox_Company)
            $Form_Outstaff_Info_User.Controls.Remove($Button_Outstaff_Create_User)
            $Text_Create_Outstaff = New-Object System.Windows.Forms.Label
            $Text_Create_Outstaff.Text = $Outcome
            $Text_Create_Outstaff.ForeColor = $Color
            $Text_Create_Outstaff.Location = New-Object System.Drawing.Point(10,20)
            $Text_Create_Outstaff.AutoSize = $true
            $Form_Outstaff_Info_User.Controls.Add($Text_Create_Outstaff)
        }
    )
            
    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$Back = Form_Ticket_Jira_Outstaff $Form_Outstaff_Info_User.Dispose(),$Back})
    $Button_Back.AutoSize = $true
                
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Label_Text
    $Label.Location  = New-Object System.Drawing.Point(110,450)
    $Label.AutoSize = $true

    $Button_Exit = New-Object System.Windows.Forms.Button
    $Button_Exit.Text = 'Exit'
    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Button_Exit.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Outstaff_Info_User = New-Object System.Windows.Forms.Form
    $Form_Outstaff_Info_User.Icon = $Icon
    $Form_Outstaff_Info_User.BackgroundImage = $Background
    $Form_Outstaff_Info_User.BackgroundImageLayout = "None"
    $Form_Outstaff_Info_User.Text ='Active Directory Account Management'
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
    $Form_Outstaff_Info_User.Controls.Add($Button_Outstaff_Create_User)
    $Form_Outstaff_Info_User.Controls.Add($Button_Back)
    $Form_Outstaff_Info_User.Controls.Add($Button_Exit)
    $Form_Outstaff_Info_User.Controls.Add($Label)
    $Form_Outstaff_Info_User.StartPosition = 'CenterScreen'
    $Form_Outstaff_Info_User.MaximizeBox = $false
    $Form_Outstaff_Info_User.FormBorderStyle = "FixedDialog"

    $Form_Ticket_Jira_Outstaff.Dispose()
    $Form_Outstaff_Info_User.ShowDialog()
}