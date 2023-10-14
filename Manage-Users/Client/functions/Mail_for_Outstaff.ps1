<#Menu function to create a mailbox for an external employee#>
<#Функция меню для создания почтового ящика для внешнего сотрудника#>
function Mail_for_Outstaff {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $Text_Mail_for_Outstaff = New-Object System.Windows.Forms.Label
    $Text_Mail_for_Outstaff.Text = "Select user"
    $Text_Mail_for_Outstaff.Location = New-Object System.Drawing.Point(10,20)
    $Text_Mail_for_Outstaff.AutoSize = $true

    $ComboBox_Mail_for_Outstaff = New-Object System.Windows.Forms.ComboBox
    $ComboBox_Mail_for_Outstaff.Width = 250
    $ComboBox_Mail_for_Outstaff.MaxDropDownItems = 20
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU_Outstaff | sort cn
    Foreach ($User in $Users){
        $ComboBox_Mail_for_Outstaff.Items.Add($User.cn);
    }
    $ComboBox_Mail_for_Outstaff.Location = New-Object System.Drawing.Point(10,50)

    $Button_Mail_for_Outstaff = New-Object System.Windows.Forms.Button
    $Button_Mail_for_Outstaff.Location = New-Object System.Drawing.Size(320,50)
    $Button_Mail_for_Outstaff.AutoSize = $true
    $Button_Mail_for_Outstaff.Text = "Create mailbox for Outstaff"
    $Button_Mail_for_Outstaff.Add_Click(
        {
            [string]$User = $ComboBox_Mail_for_Outstaff.selectedItem
            $Mail_Out = Create_Mail_for_Outstaff -User $User -DC $DC -Mail_serv $Mail_serv -Path_log $Path_log
            $Form_Mail_for_Outstaff.Controls.Remove($Text_Mail_for_Outstaff)
            $Form_Mail_for_Outstaff.Controls.Remove($ComboBox_Mail_for_Outstaff)
            $Form_Mail_for_Outstaff.Controls.Remove($Button_Mail_for_Outstaff)
            [string]$Color = $Mail_Out.Color
            [string]$Outcome = $Mail_Out.Outcome
            $Text_Mail_Out_Done = New-Object System.Windows.Forms.Label
            $Text_Mail_Out_Done.Text = $Outcome
            $Text_Mail_Out_Done.ForeColor = $Color
            $Text_Mail_Out_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Mail_Out_Done.AutoSize = $true
            $Form_Mail_for_Outstaff.Controls.Add($Text_Mail_Out_Done)
        }
    )

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$EM = ExchangeMenu $Form_Mail_for_Outstaff.Dispose(),$EM})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Mail_for_Outstaff = New-Object System.Windows.Forms.Form
    $Form_Mail_for_Outstaff.Icon = $Icon
    $Form_Mail_for_Outstaff.BackgroundImage = $Background
    $Form_Mail_for_Outstaff.BackgroundImageLayout = "None"
    $Form_Mail_for_Outstaff.Text ='Exchange account management'
    $Form_Mail_for_Outstaff.Width = 500
    $Form_Mail_for_Outstaff.Height = 500
    $Form_Mail_for_Outstaff.AutoSize = $true
    $Form_Mail_for_Outstaff.Controls.Add($Text_Mail_for_Outstaff)
    $Form_Mail_for_Outstaff.Controls.Add($ComboBox_Mail_for_Outstaff)
    $Form_Mail_for_Outstaff.Controls.Add($Button_Mail_for_Outstaff)
    $Form_Mail_for_Outstaff.Controls.Add($Button_Back)
    $Form_Mail_for_Outstaff.Controls.Add($Button_Exit)
    $Form_Mail_for_Outstaff.Controls.Add($Label)
    $Form_Mail_for_Outstaff.StartPosition = 'CenterScreen'
    $Form_Mail_for_Outstaff.MaximizeBox = $false
    $Form_Mail_for_Outstaff.FormBorderStyle = "FixedDialog"

    $ExchangeMenu.Dispose()
    $Form_Mail_for_Outstaff.ShowDialog()
}