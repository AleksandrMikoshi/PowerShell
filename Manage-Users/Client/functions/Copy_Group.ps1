<#Menu function to select users who need to scorp groups#>
<#Функция меню для выбора пользователей, кому необходимо скорпировать группы#>
function Copy_group {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $Text_From_Copy_Group = New-Object System.Windows.Forms.Label
    $Text_From_Copy_Group.Text = "User FROM which we copy groups"
    $Text_From_Copy_Group.Location = New-Object System.Drawing.Point(10,30)
    $Text_From_Copy_Group.AutoSize = $true

    <#Menu for user selection#>
    <#Меню для выбора пользователя#>
    $ComboBox_From_Copy_Group = New-Object System.Windows.Forms.ComboBox
    $ComboBox_From_Copy_Group.Width = 250
    $ComboBox_From_Copy_Group.MaxDropDownItems = 20
    $FromUsers = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($FromUser in $FromUsers){
        $ComboBox_From_Copy_Group.Items.Add($FromUser.cn);
    }
    $ComboBox_From_Copy_Group.Location = New-Object System.Drawing.Point(10,50)

    $Text_To_Copy_Group = New-Object System.Windows.Forms.Label
    $Text_To_Copy_Group.Text = "User to WHOM we copy groups"
    $Text_To_Copy_Group.Location = New-Object System.Drawing.Point(10,80)
    $Text_To_Copy_Group.AutoSize = $true

    <#Menu for user selection#>
    <#Меню для выбора пользователя#>
    $ComboBox_To_Copy_Group = New-Object System.Windows.Forms.ComboBox
    $ComboBox_To_Copy_Group.Width = 250
    $ComboBox_To_Copy_Group.MaxDropDownItems = 20
    $ToUsers = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($ToUser in $ToUsers){
        $ComboBox_To_Copy_Group.Items.Add($ToUser.cn);
    }
    $ComboBox_To_Copy_Group.Location = New-Object System.Drawing.Point(10,100)

    $Button_Copy_Group_Cont = New-Object System.Windows.Forms.Button
    $Button_Copy_Group_Cont.Location = New-Object System.Drawing.Size(360,50)
    $Button_Copy_Group_Cont.AutoSize = $true
    $Button_Copy_Group_Cont.Text = "Copy groups"
    $Button_Copy_Group_Cont.Add_Click(
        {
            [string]$SourceUser = $ComboBox_From_Copy_Group.selectedItem
            [string]$TargetUser = $ComboBox_To_Copy_Group.selectedItem
            $Copy = Copy_Group_Jea -SourceUser $SourceUser -TargetUser $TargetUser
            $Form_Copy_Group.Controls.Remove($Text_From_Copy_Group)
            $Form_Copy_Group.Controls.Remove($ComboBox_From_Copy_Group)
            $Form_Copy_Group.Controls.Remove($Text_To_Copy_Group)
            $Form_Copy_Group.Controls.Remove($ComboBox_To_Copy_Group)
            $Form_Copy_Group.Controls.Remove($Button_Copy_Group_Cont)
            [string]$Color = $Copy.Color
            [string]$Outcome = $Copy.Outcome
            $Text_Copy_Done = New-Object System.Windows.Forms.Label
            $Text_Copy_Done.Text = $Outcome
            $Text_Copy_Done.ForeColor = $Color
            $Text_Copy_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Copy_Done.AutoSize = $true
            $Form_Copy_Group.Controls.Add($Text_Copy_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_Copy_Group.Dispose(),$AD})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Copy_Group = New-Object System.Windows.Forms.Form
    $Form_Copy_Group.Icon = $Icon
    $Form_Copy_Group.BackgroundImage = $Background
    $Form_Copy_Group.BackgroundImageLayout = "None"
    $Form_Copy_Group.Text ='Active Directory Account Management'
    $Form_Copy_Group.Width = 500
    $Form_Copy_Group.Height = 500
    $Form_Copy_Group.AutoSize = $true
    $Form_Copy_Group.Controls.Add($Text_From_Copy_Group)
    $Form_Copy_Group.Controls.Add($ComboBox_From_Copy_Group)
    $Form_Copy_Group.Controls.Add($Text_To_Copy_Group)
    $Form_Copy_Group.Controls.Add($ComboBox_To_Copy_Group)
    $Form_Copy_Group.Controls.Add($Button_Copy_Group_Cont)
    $Form_Copy_Group.Controls.Add($Button_Back)
    $Form_Copy_Group.Controls.Add($Button_Exit)
    $Form_Copy_Group.Controls.Add($Label)
    $Form_Copy_Group.StartPosition = 'CenterScreen'
    $Form_Copy_Group.MaximizeBox = $false
    $Form_Copy_Group.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_Copy_Group.ShowDialog()
}