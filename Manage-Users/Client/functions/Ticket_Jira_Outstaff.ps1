<#Menu function for specifying a request in Jira for the exit of an external employee#>
<#Функция меню для указания заявки в Jira на выход внешнего сотрудника#>
function Ticket_Jira_Outstaff {
    Add-Type -assembly System.Windows.Forms

    $Button_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.Button
    $Button_Ticket_Jira_Outstaff.Location = New-Object System.Drawing.Size(393,50)
    $Button_Ticket_Jira_Outstaff.AutoSize = $true
    $Button_Ticket_Jira_Outstaff.Text = "Continue"
    $Button_Ticket_Jira_Outstaff.Add_click(
        {
            $Data = Data_From_Jira_Outstaff -PipeLine $TextBox_Ticket_Jira_Outstaff.Text
            Info_from_Jira_Outstaff -Data $Data
        }
    )

    $New_Label = New-Object System.Windows.Forms.Label
    $New_Label.Text = "Specify application number"
    $New_Label.Location  = New-Object System.Drawing.Point(10,20)
    $New_Label.AutoSize = $true

    $TextBox_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.TextBox
    $TextBox_Ticket_Jira_Outstaff.Location  = New-Object System.Drawing.Point(10,50)
    $TextBox_Ticket_Jira_Outstaff.Text = ''
    $TextBox_Ticket_Jira_Outstaff.Size = New-Object System.Drawing.Size(270,30)

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Ticket_Jira_Outstaff.Dispose(),$AD})
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
    $Form_Ticket_Jira_Outstaff = New-Object System.Windows.Forms.Form
    $Form_Ticket_Jira_Outstaff.Icon = $Icon
    $Form_Ticket_Jira_Outstaff.BackgroundImage = $Background
    $Form_Ticket_Jira_Outstaff.BackgroundImageLayout = "None"
    $Form_Ticket_Jira_Outstaff.Text ='Active Directory Account Management'
    $Form_Ticket_Jira_Outstaff.Width = 500
    $Form_Ticket_Jira_Outstaff.Height = 500
    $Form_Ticket_Jira_Outstaff.AutoSize = $true
    $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Ticket_Jira_Outstaff)
    $Form_Ticket_Jira_Outstaff.Controls.Add($New_Label)
    $Form_Ticket_Jira_Outstaff.Controls.Add($TextBox_Ticket_Jira_Outstaff)
    $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Back)
    $Form_Ticket_Jira_Outstaff.Controls.Add($Button_Exit)
    $Form_Ticket_Jira_Outstaff.Controls.Add($Label)
    $Form_Ticket_Jira_Outstaff.StartPosition = 'CenterScreen'
    $Form_Ticket_Jira_Outstaff.MaximizeBox = $false
    $Form_Ticket_Jira_Outstaff.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Ticket_Jira_Outstaff.ShowDialog()
}