<#Menu call function for managing accounts in Active Directory#>
<#Функция вызова меню для управления учетными записями в Active Directory#>
function ActiveDirectoryMenu {
    Add-Type -assembly System.Windows.Forms

    <#Button for further user creation#>
    <#Кнопка для дальнейшего создания пользователя#>
    $Button_Info_user = New-Object System.Windows.Forms.Button
    $Button_Info_user.Text = 'Create a new user'
    $Button_Info_user.Location = New-Object System.Drawing.Point(10,20)
    $Button_Info_user.Add_click({Ticket_Jira_User})
    $Button_Info_user.AutoSize = $true

    <#Button for further creation of an external user#>
    <#Кнопка для дальнейшего создания внешнего пользователя#>
    $Button_Info_user_outstaff = New-Object System.Windows.Forms.Button
    $Button_Info_user_outstaff.Text = 'Create a new Outstaff user'
    $Button_Info_user_outstaff.Location = New-Object System.Drawing.Point(10,50)
    $Button_Info_user_outstaff.Add_click({Ticket_Jira_Outstaff})
    $Button_Info_user_outstaff.AutoSize = $true

    <#Button for copying groups#>
    <#Кнопка для копирования групп#>
    $Button_Copy_group = New-Object System.Windows.Forms.Button
    $Button_Copy_group.Text = 'Copy groups from user to user'
    $Button_Copy_group.Location = New-Object System.Drawing.Point(10,80)
    $Button_Copy_group.Add_click({Copy_group})
    $Button_Copy_group.AutoSize = $true

    <#Password reset button#>
    <#Кнопка для сброса пароля#>
    $Button_Reset_Pass = New-Object System.Windows.Forms.Button
    $Button_Reset_Pass.Text = 'Reset account password'
    $Button_Reset_Pass.Location = New-Object System.Drawing.Point(10,110)
    $Button_Reset_Pass.Add_click({Reset_Pass})
    $Button_Reset_Pass.AutoSize = $true

    <#Button to unlock user#>
    <#Кнопка для разблокировки пользователя#>
    $Button_Unlock_user = New-Object System.Windows.Forms.Button
    $Button_Unlock_user.Text = 'Unlock account'
    $Button_Unlock_user.Location = New-Object System.Drawing.Point(10,140)
    $Button_Unlock_user.Add_click({Unlock_user})
    $Button_Unlock_user.AutoSize = $true

    <#Button to fire an employee#>
    <#Кнопка для увольнения сотрудника#>
    $Button_Fired_user = New-Object System.Windows.Forms.Button
    $Button_Fired_user.Text = 'Dismissing an employee'
    $Button_Fired_user.Location = New-Object System.Drawing.Point(10,170)
    $Button_Fired_user.Add_click({Fired_user})
    $Button_Fired_user.AutoSize = $true

    <#Button for deleting access groups#>
    <#Кнопка для удаления групп доступа#>
    $Button_Remove_group = New-Object System.Windows.Forms.Button
    $Button_Remove_group.Text = 'Removing access groups (access to mail will remain)'
    $Button_Remove_group.Location = New-Object System.Drawing.Point(10,200)
    $Button_Remove_group.Add_click({Remove_group})
    $Button_Remove_group.AutoSize = $true

    <#Button to exit to the main menu#>
    <#Кнопка для выхода на главное меню#>
    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Main menu'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$Main_Menu = main_menu $ActiveDirectoryMenu.Dispose(),$Main_Menu})
    $Button_Main_Menu.AutoSize = $true

    <#Signature at the bottom of the form#>
    <#Подпись снизу формы#>
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Label_Text
    $Label.Location  = New-Object System.Drawing.Point(110,450)
    $Label.AutoSize = $true

    <#Form close button#>
    <#Кнопка закрытия формы#>
    $Button_Exit = New-Object System.Windows.Forms.Button
    $Button_Exit.Text = 'Exit'
    $Button_Exit.Location = New-Object System.Drawing.Point(400,400)
    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Button_Exit.AutoSize = $true

    <#Adding a background#>
    <#Добавление фона#>
    $Background_AD = [system.drawing.image]::FromFile($Background_Image)
    <#Creating a form#>
    <#Создание формы#>
    $ActiveDirectoryMenu = New-Object System.Windows.Forms.Form
    $ActiveDirectoryMenu.Icon = $Icon
    $ActiveDirectoryMenu.BackgroundImage = $Background_AD
    $ActiveDirectoryMenu.BackgroundImageLayout = "None"
    $ActiveDirectoryMenu.Text ='Active Directory Account Management'
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
    $ActiveDirectoryMenu.StartPosition = 'CenterScreen'
    $ActiveDirectoryMenu.MaximizeBox = $false
    $ActiveDirectoryMenu.FormBorderStyle = "FixedDialog"

    $main_form.Dispose()
    $ActiveDirectoryMenu.ShowDialog()
}