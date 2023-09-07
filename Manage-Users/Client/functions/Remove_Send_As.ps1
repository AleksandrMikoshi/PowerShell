<#Menu function for specifying data for further disabling rights#>
<#Функция меню для указания данных для дальнейшего отключения прав#>
function Remove_Send_as {
    Add-Type -assembly System.Windows.Forms

    $ComboBox_Remove_Send_as = New-Object System.Windows.Forms.CheckedListBox
    $ComboBox_Remove_Send_as.Width = 250
    $ComboBox_Remove_Send_as.Height = 300
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Remove_Send_as.Items.Add($User.cn);
    }
    $ComboBox_Remove_Send_as.Location = New-Object System.Drawing.Point(10,90)

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Enter mailbox name"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,20)
    $Label_ticket_1.AutoSize = $true

    $TextBox_Remove_Send_as = New-Object System.Windows.Forms.TextBox
    $TextBox_Remove_Send_as.Location  = New-Object System.Drawing.Point(10,40)
    $TextBox_Remove_Send_as.Text = ""
    $TextBox_Remove_Send_as.Size = New-Object System.Drawing.Size(250,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "Select the users for whom you want to disable the 'Send as' right"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,70)
    $Label_ticket_2.AutoSize = $true

    $Button_Remove_Send_as = New-Object System.Windows.Forms.Button
    $Button_Remove_Send_as.Location = New-Object System.Drawing.Size(350,40)
    $Button_Remove_Send_as.AutoSize = $true
    $Button_Remove_Send_as.Text = "Disable rights"
    $Button_Remove_Send_as.Add_Click(
        { 
            $Mail = $TextBox_Remove_Send_as.Text
            $Users = $ComboBox_Remove_Send_as.CheckedItems
            $RemoveSendAs = Remove_SendAs -Mail $Mail -Users $Users -Mail_serv $Mail_serv -Path_log $Path_log
            $Menu_Remove_Send_as.Controls.Remove($Label_ticket_1)
            $Menu_Remove_Send_as.Controls.Remove($Label_ticket_2)
            $Menu_Remove_Send_as.Controls.Remove($TextBox_Remove_Send_as)
            $Menu_Remove_Send_as.Controls.Remove($ComboBox_Remove_Send_as)
            $Menu_Remove_Send_as.Controls.Remove($Button_Remove_Send_as)
            [string]$Color = $RemoveSendAs.Color
            [string]$Outcome = $RemoveSendAs.Outcome
            $Text_RSA_Done = New-Object System.Windows.Forms.Label
            $Text_RSA_Done.Text = $Outcome
            $Text_RSA_Done.ForeColor = $Color
            $Text_RSA_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_RSA_Done.AutoSize = $true
            $Menu_Remove_Send_as.Controls.Add($Text_RSA_Done)
        }
    )

    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Back'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$EM = ExchangeMenu $Menu_Remove_Send_as.Dispose(),$EM})
    $Button_Main_Menu.AutoSize = $true

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
    $Menu_Remove_Send_as = New-Object System.Windows.Forms.Form
    $Menu_Remove_Send_as.Icon = $Icon
    $Menu_Remove_Send_as.BackgroundImage = $Background
    $Menu_Remove_Send_as.BackgroundImageLayout = "None"
    $Menu_Remove_Send_as.Text ='Exchange account management'
    $Menu_Remove_Send_as.Width = 500
    $Menu_Remove_Send_as.Height = 500
    $Menu_Remove_Send_as.AutoSize = $true
    $Menu_Remove_Send_as.Controls.Add($Label_ticket_1)
    $Menu_Remove_Send_as.Controls.Add($Label_ticket_2)
    $Menu_Remove_Send_as.Controls.Add($TextBox_Remove_Send_as)
    $Menu_Remove_Send_as.Controls.Add($ComboBox_Remove_Send_as)
    $Menu_Remove_Send_as.Controls.Add($Button_Remove_Send_as)
    $Menu_Remove_Send_as.Controls.Add($Button_Main_Menu)
    $Menu_Remove_Send_as.Controls.Add($Button_Exit)
    $Menu_Remove_Send_as.Controls.Add($Label)
    $Menu_Remove_Send_as.StartPosition = 'CenterScreen'
    $Menu_Remove_Send_as.MaximizeBox = $false
    $Menu_Remove_Send_as.FormBorderStyle = "FixedDialog"

    $ExchangeMenu.Dispose()
    $Menu_Remove_Send_as.ShowDialog()
}