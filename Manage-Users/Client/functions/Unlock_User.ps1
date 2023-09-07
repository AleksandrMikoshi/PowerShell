<#Menu function to select user to unlock#>
<#Функция меню для выбора пользователя для разблокировки#>
function Unlock_user {
    Add-Type -assembly System.Windows.Forms

    $Text_Unlock_user = New-Object System.Windows.Forms.Label
    $Text_Unlock_user.Text = "Select user"
    $Text_Unlock_user.Location = New-Object System.Drawing.Point(10,20)
    $Text_Unlock_user.AutoSize = $true
            
    $ComboBox_Unlock_user = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Unlock_user.Width = 250
    $ComboBox_Unlock_user.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Unlock_user.Items.Add($User.cn);
    }
    $ComboBox_Unlock_user.Location = New-Object System.Drawing.Point(10,50)

    $Button_Unlock_user = New-Object System.Windows.Forms.Button
    $Button_Unlock_user.Location = New-Object System.Drawing.Size(300,50)
    $Button_Unlock_user.AutoSize = $true
    $Button_Unlock_user.Text = "Unblock user"
    $Button_Unlock_user.Add_Click(
        {
            [string]$User = $ComboBox_Unlock_user.selectedItem
            $Unlock = Unlock_Account -User $User
            $Form_Unlock_user.Controls.Remove($Text_Unlock_user)
            $Form_Unlock_user.Controls.Remove($ComboBox_Unlock_user)
            $Form_Unlock_user.Controls.Remove($Button_Unlock_user)
            [string]$Color = $Unlock.Color
            [string]$Outcome = $Unlock.Outcome
            $Text_Unlock_Done = New-Object System.Windows.Forms.Label
            $Text_Unlock_Done.Text = $Outcome
            $Text_Unlock_Done.ForeColor = $Color
            $Text_Unlock_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Unlock_Done.AutoSize = $true
            $Form_Unlock_user.Controls.Add($Text_Unlock_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Unlock_user.Dispose(),$AD})
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
    $Form_Unlock_user = New-Object System.Windows.Forms.Form
    $Form_Unlock_user.Icon = $Icon
    $Form_Unlock_user.BackgroundImage = $Background
    $Form_Unlock_user.BackgroundImageLayout = "None"
    $Form_Unlock_user.Text ='Active Directory Account Management'
    $Form_Unlock_user.Width = 500
    $Form_Unlock_user.Height = 500
    $Form_Unlock_user.AutoSize = $true
    $Form_Unlock_user.Controls.Add($Text_Unlock_user)
    $Form_Unlock_user.Controls.Add($ComboBox_Unlock_user)
    $Form_Unlock_user.Controls.Add($Button_Unlock_user)
    $Form_Unlock_user.Controls.Add($Button_Back)
    $Form_Unlock_user.Controls.Add($Button_Exit)
    $Form_Unlock_user.Controls.Add($Label)
    $Form_Unlock_user.StartPosition = 'CenterScreen'
    $Form_Unlock_user.MaximizeBox = $false
    $Form_Unlock_user.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Unlock_user.ShowDialog()
}