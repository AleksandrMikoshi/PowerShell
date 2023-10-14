<#Menu function to view when the user last changed the password#>
<#Функция меню для просмотра когда пользователь последний раз менял пароль#>
function Info_Pass {
    Label
    Button_Exit 

    $Text_Info_Pass = New-Object System.Windows.Forms.Label
    $Text_Info_Pass.Text = "Select user"
    $Text_Info_Pass.Location = New-Object System.Drawing.Point(10,20)
    $Text_Info_Pass.AutoSize = $true

    $ComboBox_Info_Pass = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Info_Pass.Width = 250
    $ComboBox_Info_Pass.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users)
        {
            $ComboBox_Info_Pass.Items.Add($User.cn);
        }
    $ComboBox_Info_Pass.Location = New-Object System.Drawing.Point(10,50)

    $Text_Info_Pass_2 = New-Object System.Windows.Forms.Label
    $Text_Info_Pass_2.Text = "Last password change:"
    $Text_Info_Pass_2.Location = New-Object System.Drawing.Point(10,80)
    $Text_Info_Pass_2.AutoSize = $true

    $Text_Info_Pass_3 = New-Object System.Windows.Forms.Label
    $Text_Info_Pass_3.Text = ""
    $Text_Info_Pass_3.Location = New-Object System.Drawing.Point(10,110)
    $Text_Info_Pass_3.AutoSize = $true

    $Button_Check_Pass = New-Object System.Windows.Forms.Button
    $Button_Check_Pass.Location = New-Object System.Drawing.Size(394,50)
    $Button_Check_Pass.AutoSize = $true
    $Button_Check_Pass.Text = "Check"
    $Button_Check_Pass.Add_Click(
        {
            [string]$Name = $ComboBox_Info_Pass.selectedItem
            $Text_Info_Pass_3.Text = [datetime]::FromFileTime((Get-ADUser -Filter "cn -eq $Name" -Properties pwdLastSet).pwdLastSet).ToString('dd.MM.yyyy hh:mm')
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$IM = InfoMenu $Info_Pass.Dispose(),$IM})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Info_Pass = New-Object System.Windows.Forms.Form
    $Info_Pass.Icon = $Icon
    $Info_Pass.BackgroundImage = $Background
    $Info_Pass.BackgroundImageLayout = "None"
    $Info_Pass.Text ='Receiving the information'
    $Info_Pass.Width = 500
    $Info_Pass.Height = 500
    $Info_Pass.AutoSize = $true
    $Info_Pass.Controls.Add($Text_Info_Pass)
    $Info_Pass.Controls.Add($ComboBox_Info_Pass)
    $Info_Pass.Controls.Add($Text_Info_Pass_2)
    $Info_Pass.Controls.Add($Text_Info_Pass_3)
    $Info_Pass.Controls.Add($Button_Check_Pass)
    $Info_Pass.Controls.Add($Button_Back)
    $Info_Pass.Controls.Add($Button_Exit)
    $Info_Pass.Controls.Add($Label)
    $Info_Pass.StartPosition = 'CenterScreen'
    $Info_Pass.MaximizeBox = $false
    $Info_Pass.FormBorderStyle = "FixedDialog"

    $InfoMenu.Dispose()
    $Info_Pass.ShowDialog()
}