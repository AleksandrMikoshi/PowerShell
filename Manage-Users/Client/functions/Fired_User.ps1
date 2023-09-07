<#Function for the employee selection menu for dismissal#>
<#Функция для меню выбора сотрудника для увольнения#>
function Fired_user {
    Add-Type -assembly System.Windows.Forms

    $Text_Fired_user = New-Object System.Windows.Forms.Label
    $Text_Fired_user.Text = "Select user"
    $Text_Fired_user.Location = New-Object System.Drawing.Point(10,20)
    $Text_Fired_user.AutoSize = $true
            
    $ComboBox_Fired_user = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Fired_user.Width = 250
    $ComboBox_Fired_user.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Fired_user.Items.Add($User.cn);
    }
    $ComboBox_Fired_user.Location = New-Object System.Drawing.Point(10,50)

    $Button_Fired_user = New-Object System.Windows.Forms.Button
    $Button_Fired_user.Location = New-Object System.Drawing.Size(270,50)
    $Button_Fired_user.AutoSize = $true
    $Button_Fired_user.Text = "Carry out user dismissal"
    $Button_Fired_user.Add_Click(
        {
            [string]$User = $ComboBox_Fired_user.selectedItem
            $Fired = Fired_User_Jea -User $User -Mail_serv $Mail_serv -DC $DC -Path_fired $Path_fired -Path_log $Path_log
            $Form_Fired_user.Controls.Remove($Text_Fired_user)
            $Form_Fired_user.Controls.Remove($ComboBox_Fired_user)
            $Form_Fired_user.Controls.Remove($Button_Fired_user)
            [string]$Color = $Fired.Color
            [string]$Outcome = $Fired.Outcome
            $Text_Fired_Done = New-Object System.Windows.Forms.Label
            $Text_Fired_Done.Text = $Outcome
            $Text_Fired_Done.ForeColor = $Color
            $Text_Fired_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Fired_Done.AutoSize = $true
            $Form_Fired_user.Controls.Add($Text_Fired_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Fired_user.Dispose(),$AD})
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
    $Form_Fired_user = New-Object System.Windows.Forms.Form
    $Form_Fired_user.Icon = $Icon
    $Form_Fired_user.BackgroundImage = $Background
    $Form_Fired_user.BackgroundImageLayout = "None"
    $Form_Fired_user.Text ='Active Directory Account Management'
    $Form_Fired_user.Width = 500
    $Form_Fired_user.Height = 500
    $Form_Fired_user.AutoSize = $true
    $Form_Fired_user.Controls.Add($Text_Fired_user)
    $Form_Fired_user.Controls.Add($ComboBox_Fired_user)
    $Form_Fired_user.Controls.Add($Button_Fired_user)
    $Form_Fired_user.Controls.Add($Button_Back)
    $Form_Fired_user.Controls.Add($Button_Exit)
    $Form_Fired_user.Controls.Add($Label)
    $Form_Fired_user.StartPosition = 'CenterScreen'
    $Form_Fired_user.MaximizeBox = $false
    $Form_Fired_user.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Fired_user.ShowDialog()
}