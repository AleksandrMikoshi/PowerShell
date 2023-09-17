<#Menu call function for viewing information data#>
<#Функция вызова меню для просмотра информационных данных#>
function InfoMenu {
    Add-Type -assembly System.Windows.Forms

    <#Button to view when the password was last changed#>
    <#Кнопка для просмотра когда последний раз меняли пароль#>
    $Button_Info_Pass = New-Object System.Windows.Forms.Button
    $Button_Info_Pass.Text = 'See when the password was last changed'
    $Button_Info_Pass.Location = New-Object System.Drawing.Point(10,20)
    $Button_Info_Pass.Add_click({Info_Pass})
    $Button_Info_Pass.AutoSize = $true

    <#Button to exit to the main menu#>
    <#Кнопка для выхода на главное меню#>
    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Main menu'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$Main_menu = main_menu $InfoMenu.Dispose(),$Main_menu})
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
    $Background = [system.drawing.image]::FromFile($Background_Image)
    <#Creating a form#>
    <#Создание формы#>
    $InfoMenu = New-Object System.Windows.Forms.Form
    $InfoMenu.Icon = $Icon
    $InfoMenu.BackgroundImage = $Background
    $InfoMenu.BackgroundImageLayout = "None"
    $InfoMenu.Text ='Receiving the information'
    $InfoMenu.Width = 500
    $InfoMenu.Height = 500
    $InfoMenu.AutoSize = $true
    $InfoMenu.Controls.Add($Button_Main_Menu)
    $InfoMenu.Controls.Add($Button_Info_Pass)
    $InfoMenu.Controls.Add($Button_Exit)
    $InfoMenu.Controls.Add($Label)
    $InfoMenu.StartPosition = 'CenterScreen'
    $InfoMenu.MaximizeBox = $false
    $InfoMenu.FormBorderStyle = "FixedDialog"

    $main_form.Dispose()
    $InfoMenu.ShowDialog()
}