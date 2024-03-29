<#Menu call function for managing accounts in Exchange#>
<#Функция вызова меню для управления учетными записями в Exchange#>
function ExchangeMenu {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    <#Button to create a mailbox#>
    <#Кнопка для создания почтового ящика#>
    $Button_Create_mail = New-Object System.Windows.Forms.Button
    $Button_Create_mail.Text = 'Create a mailbox'
    $Button_Create_mail.Location = New-Object System.Drawing.Point(10,20)
    $Button_Create_mail.Add_click({Create_mail})
    $Button_Create_mail.AutoSize = $true

    <#Button to grant "Send As" permissions#>
    <#Кнопка для предоставления прав "Отправить как"#>
    $Button_Send_as = New-Object System.Windows.Forms.Button
    $Button_Send_as.Text = 'Grant "Send As" rights'
    $Button_Send_as.Location = New-Object System.Drawing.Point(10,50)
    $Button_Send_as.Add_click({Sendas})
    $Button_Send_as.AutoSize = $true

    <#Button to disable "Send As" rights#>
    <#Кнопка для отключения прав "Отправить как"#>
    $Button_Remove_Send_as = New-Object System.Windows.Forms.Button
    $Button_Remove_Send_as.Text = 'Disable "Send as" rights'
    $Button_Remove_Send_as.Location = New-Object System.Drawing.Point(10,80)
    $Button_Remove_Send_as.Add_click({Remove_Send_as})
    $Button_Remove_Send_as.AutoSize = $true

    <#Button to grant "Full Access" rights#>
    <#Кнопка для предоставления прав "Полный доступ"#>
    $Button_Full_access.Text = 'Granting "Full Access" rights'
    $Button_Full_access.Location = New-Object System.Drawing.Point(10,110)
    $Button_Full_access.Add_click({Full_Access})
    $Button_Full_access.AutoSize = $true

    <#Button to disable "Full Access" rights#>
    <#Кнопка для отключения прав "Полный доступ"#>
    $Button_Remove_Full_access = New-Object System.Windows.Forms.Button
    $Button_Remove_Full_access.Text = 'Disable "Full Access" rights'
    $Button_Remove_Full_access.Location = New-Object System.Drawing.Point(10,140)
    $Button_Remove_Full_access.Add_click({Remove_Full_access})
    $Button_Remove_Full_access.AutoSize = $true

    <#Button to create a mailbox for an external employee#>
    <#Кнопка для создания почтового ящика для внешнего сотрудника#>
    $Button_Mail_for_Outstaff = New-Object System.Windows.Forms.Button
    $Button_Mail_for_Outstaff.Text = 'Connecting a mailbox for Outstaff'
    $Button_Mail_for_Outstaff.Location = New-Object System.Drawing.Point(10,170)
    $Button_Mail_for_Outstaff.Add_click({Mail_for_Outstaff})
    $Button_Mail_for_Outstaff.AutoSize = $true

    <#Button to exit to the main menu#>
    <#Кнопка для выхода на главное меню#>
    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Main menu'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$Main_menu = main_menu $ExchangeMenu.Dispose(),$Main_menu})
    $Button_Main_Menu.AutoSize = $true

    <#Adding a background#>
    <#Добавление фона#>
    $Background = [system.drawing.image]::FromFile($Background_Image)
    <#Creating a form#>
    <#Создание формы#>
    $ExchangeMenu = New-Object System.Windows.Forms.Form
    $ExchangeMenu.Icon = $Icon
    $ExchangeMenu.BackgroundImage = $Background
    $ExchangeMenu.BackgroundImageLayout = "None"
    $ExchangeMenu.Text ='Exchange account management'
    $ExchangeMenu.Width = 500
    $ExchangeMenu.Height = 500
    $ExchangeMenu.AutoSize = $true
    $ExchangeMenu.Controls.Add($Button_Create_mail)
    $ExchangeMenu.Controls.Add($Button_Send_as)
    $ExchangeMenu.Controls.Add($Button_Remove_Send_as)
    $ExchangeMenu.Controls.Add($Button_Full_access)
    $ExchangeMenu.Controls.Add($Button_Remove_Full_access)
    $ExchangeMenu.Controls.Add($Button_Mail_for_Outstaff)
    $ExchangeMenu.Controls.Add($Button_Main_Menu)
    $ExchangeMenu.Controls.Add($Button_Exit)
    $ExchangeMenu.Controls.Add($Label)
    $ExchangeMenu.StartPosition = 'CenterScreen'
    $ExchangeMenu.MaximizeBox = $false
    $ExchangeMenu.FormBorderStyle = "FixedDialog"

    $main_form.Dispose()
    $ExchangeMenu.ShowDialog()
}