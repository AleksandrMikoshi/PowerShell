<#Menu function for deleting access groups#>
<#Функция меню для удаления групп доступа#>
function Remove_Group {
    Add-Type -assembly System.Windows.Forms

    $Text_Remove_Group = New-Object System.Windows.Forms.Label
    $Text_Remove_Group.Text = "Select user"
    $Text_Remove_Group.Location = New-Object System.Drawing.Point(10,20)
    $Text_Remove_Group.AutoSize = $true
            
    $ComboBox_Remove_Group = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Remove_Group.Width = 250
    $ComboBox_Remove_Group.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Remove_Group.Items.Add($User.cn);
    }
    $ComboBox_Remove_Group.Location = New-Object System.Drawing.Point(10,50)

    $Button_Remove_Group = New-Object System.Windows.Forms.Button
    $Button_Remove_Group.Location = New-Object System.Drawing.Size(320,50)
    $Button_Remove_Group.AutoSize = $true
    $Button_Remove_Group.Text = "Disable access groups"
    $Button_Remove_Group.Add_Click(
        {
            [string]$User = $ComboBox_Remove_Group.selectedItem
            $Remove = Remove_Group_Jea -User $User
            $Form_Remove_Group.Controls.Remove($Text_Remove_Group)
            $Form_Remove_Group.Controls.Remove($ComboBox_Remove_Group)
            $Form_Remove_Group.Controls.Remove($Button_Remove_Group)
            [string]$Color = $Remove.Color
            [string]$Outcome = $Remove.Outcome
            $Text_Remove_Done.Text = $Outcome
            $Text_Remove_Done.ForeColor = $Color
            $Text_Remove_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Remove_Done.AutoSize = $true
            $Form_Remove_Group.Controls.Add($Text_Remove_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Remove_Group.Hide(),$AD})
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
    $Form_Remove_Group = New-Object System.Windows.Forms.Form
    $Form_Remove_Group.Icon = $Icon
    $Form_Remove_Group.BackgroundImage = $Background
    $Form_Remove_Group.BackgroundImageLayout = "None"
    $Form_Remove_Group.Text ='Active Directory Account Management'
    $Form_Remove_Group.Width = 500
    $Form_Remove_Group.Height = 500
    $Form_Remove_Group.AutoSize = $true
    $Form_Remove_Group.Controls.Add($Text_Remove_Group)
    $Form_Remove_Group.Controls.Add($ComboBox_Remove_Group)
    $Form_Remove_Group.Controls.Add($Button_Remove_Group)
    $Form_Remove_Group.Controls.Add($Button_Back)
    $Form_Remove_Group.Controls.Add($Button_Exit)
    $Form_Remove_Group.Controls.Add($Label)
    $Form_Remove_Group.StartPosition = 'CenterScreen'
    $Form_Remove_Group.MaximizeBox = $false
    $Form_Remove_Group.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Remove_Group.ShowDialog()
}