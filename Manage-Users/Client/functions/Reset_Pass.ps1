<#Password reset menu function#>
<#Функция меню для сброса пароля#>
function Reset_Pass {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $Text_Reset_Pass = New-Object System.Windows.Forms.Label
    $Text_Reset_Pass.Text = "Select user"
    $Text_Reset_Pass.Location = New-Object System.Drawing.Point(10,20)
    $Text_Reset_Pass.AutoSize = $true

    $ComboBox_Reset_Pass = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Reset_Pass.Width = 250
    $ComboBox_Reset_Pass.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Reset_Pass.Items.Add($User.cn);
    }
    $ComboBox_Reset_Pass.Location = New-Object System.Drawing.Point(10,50)

    $Button_Reset_Pass = New-Object System.Windows.Forms.Button
    $Button_Reset_Pass.Location = New-Object System.Drawing.Size(365,50)
    $Button_Reset_Pass.AutoSize = $true
    $Button_Reset_Pass.Text = "Reset the password"
    $Button_Reset_Pass.Add_Click(
        {
            $Form_Reset_Pass.Controls.Remove($Text_Reset_Pass)
            $Form_Reset_Pass.Controls.Remove($ComboBox_Reset_Pass)
            $Form_Reset_Pass.Controls.Remove($Button_Reset_Pass)
            [string]$User = $ComboBox_Reset_Pass.selectedItem
            $Password = Reset_Password -User $User
            [string]$Color = $Password.Color
            [string]$Outcome = $Password.Outcome
            $Text_Password_Done = New-Object System.Windows.Forms.Label
            $Text_Password_Done.Text = $Outcome
            $Text_Password_Done.ForeColor = $Color
            $Text_Password_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Password_Done.AutoSize = $true
            $Form_Reset_Pass.Controls.Add($Text_Password_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Reset_Pass.Dispose(),$AD})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Reset_Pass = New-Object System.Windows.Forms.Form
    $Form_Reset_Pass.Icon = $Icon
    $Form_Reset_Pass.BackgroundImage = $Background
    $Form_Reset_Pass.BackgroundImageLayout = "None"
    $Form_Reset_Pass.Text ='Active Directory Account Management'
    $Form_Reset_Pass.Width = 500
    $Form_Reset_Pass.Height = 500
    $Form_Reset_Pass.AutoSize = $true
    $Form_Reset_Pass.Controls.Add($Text_Reset_Pass)
    $Form_Reset_Pass.Controls.Add($ComboBox_Reset_Pass)
    $Form_Reset_Pass.Controls.Add($Button_Reset_Pass)
    $Form_Reset_Pass.Controls.Add($Button_Back)
    $Form_Reset_Pass.Controls.Add($Button_Exit)
    $Form_Reset_Pass.Controls.Add($Label)
    $Form_Reset_Pass.StartPosition = 'CenterScreen'
    $Form_Reset_Pass.MaximizeBox = $false
    $Form_Reset_Pass.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Reset_Pass.ShowDialog()
}