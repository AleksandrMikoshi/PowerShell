<#User Account Control main menu#>
<#Главное меню управления учетными записями пользователей#>
function main_menu {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $global:Button_ADMenu = New-Object System.Windows.Forms.Button
    $global:Button_ADMenu.Text = 'Active Directory Account Management'
    $global:Button_ADMenu.Location = New-Object System.Drawing.Point(10,20)
    $global:Button_ADMenu.Add_click({ActiveDirectoryMenu})
    $global:Button_ADMenu.AutoSize = $true
    
    $global:Button_ExchMenu = New-Object System.Windows.Forms.Button
    $global:Button_ExchMenu.Text = 'Exchange account management'
    $global:Button_ExchMenu.Location = New-Object System.Drawing.Point(10,50)
    $global:Button_ExchMenu.Add_click({ExchangeMenu})
    $global:Button_ExchMenu.AutoSize = $true

    $global:Button_Info = New-Object System.Windows.Forms.Button
    $global:Button_Info.Text = 'Receiving the information'
    $global:Button_Info.Location = New-Object System.Drawing.Point(10,80)
    $global:Button_Info.Add_click({InfoMenu})
    $global:Button_Info.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.Icon = $Icon
    $main_form.BackgroundImage = $Background
    $main_form.BackgroundImageLayout = "None"
    $main_form.Text ='Account Management'
    $main_form.Width = 500
    $main_form.Height = 500
    $main_form.AutoSize = $true
    $main_form.Controls.Add($Button_ADMenu)
    $main_form.Controls.Add($Button_ExchMenu)
    $main_form.Controls.Add($Button_Info)
    $main_form.Controls.Add($Button_Exit)
    $main_form.Controls.Add($Label)
    $main_form.StartPosition = 'CenterScreen'
    $main_form.MaximizeBox = $false
    $main_form.FormBorderStyle = "FixedDialog"

    $main_form.ShowDialog()
} 