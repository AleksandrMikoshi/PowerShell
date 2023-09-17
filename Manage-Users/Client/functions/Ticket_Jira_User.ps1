<#Menu function for specifying a request in Jira for employee exit#>
<#Функция меню для указания заявки в Jira на выход сотрудника#>
function Ticket_Jira_User {
    Add-Type -assembly System.Windows.Forms

    $Button_Ticket_Jira = New-Object System.Windows.Forms.Button
    $Button_Ticket_Jira.Location = New-Object System.Drawing.Size(393,50)
    $Button_Ticket_Jira.AutoSize = $true
    $Button_Ticket_Jira.Text = "Continue"
    $Button_Ticket_Jira.Add_click(
        {
            $Data = Data_From_Jira -PipeLine $TextBox_Ticket_Jira.Text -Fields $Fields
            Info_from_Jira -Data $Data
        }
    )
    $New_Label = New-Object System.Windows.Forms.Label
    $New_Label.Text = "Specify application number"
    $New_Label.Location  = New-Object System.Drawing.Point(10,20)
    $New_Label.AutoSize = $true

    $TextBox_Ticket_Jira = New-Object System.Windows.Forms.TextBox
    $TextBox_Ticket_Jira.Location  = New-Object System.Drawing.Point(10,50)
    $TextBox_Ticket_Jira.Text = ''
    $TextBox_Ticket_Jira.Size = New-Object System.Drawing.Size(270,30)

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Ticket_Jira.Dispose(),$AD})
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
    $Form_Ticket_Jira = New-Object System.Windows.Forms.Form
    $Form_Ticket_Jira.Icon = $Icon
    $Form_Ticket_Jira.BackgroundImage = $Background
    $Form_Ticket_Jira.BackgroundImageLayout = "None"
    $Form_Ticket_Jira.Text ='Active Directory Account Management'
    $Form_Ticket_Jira.Width = 500
    $Form_Ticket_Jira.Height = 500
    $Form_Ticket_Jira.AutoSize = $true
    $Form_Ticket_Jira.Controls.Add($Button_Ticket_Jira)
    $Form_Ticket_Jira.Controls.Add($New_Label)
    $Form_Ticket_Jira.Controls.Add($TextBox_Ticket_Jira)
    $Form_Ticket_Jira.Controls.Add($Label)
    $Form_Ticket_Jira.Controls.Add($Button_Back)
    $Form_Ticket_Jira.Controls.Add($Button_Exit)
    $Form_Ticket_Jira.StartPosition = 'CenterScreen'
    $Form_Ticket_Jira.MaximizeBox = $false
    $Form_Ticket_Jira.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Ticket_Jira.ShowDialog()
}