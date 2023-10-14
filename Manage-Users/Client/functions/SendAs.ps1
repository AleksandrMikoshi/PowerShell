<#Menu function for specifying data for further granting of rights#>
<#Функция меню для указания данных для дальнейшего предоставления прав#>
function Sendas{
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $ComboBox_Send_as = New-Object System.Windows.Forms.CheckedListBox
    $ComboBox_Send_as.Width = 250
    $ComboBox_Send_as.Height = 300
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Send_as.Items.Add($User.cn);
    }
    $ComboBox_Send_as.Location = New-Object System.Drawing.Point(10,90)

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Enter mailbox name"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,20)
    $Label_ticket_1.AutoSize = $true

    $TextBox_Send_as = New-Object System.Windows.Forms.TextBox
    $TextBox_Send_as.Location  = New-Object System.Drawing.Point(10,40)
    $TextBox_Send_as.Text = ""
    $TextBox_Send_as.Size = New-Object System.Drawing.Size(250,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "Select the users you want to grant the 'Send As' right to"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,70)
    $Label_ticket_2.AutoSize = $true


    $Button_Send_as = New-Object System.Windows.Forms.Button
    $Button_Send_as.Location = New-Object System.Drawing.Size(350,40)
    $Button_Send_as.AutoSize = $true
    $Button_Send_as.Text = "grant rights"
    $Button_Send_as.Add_Click(
        { 
            $Mail = $TextBox_Send_as.Text
            $Users = $ComboBox_Send_as.CheckedItems
            $SendAs = Send_As -Mail $Mail -Users $Users -Mail_serv $Mail_serv -Path_log $Path_log
            [string]$Color = $SendAs.Color
            [string]$Outcome = $SendAs.Outcome
            $Menu_Send_as.Controls.Remove($Label_ticket_1)
            $Menu_Send_as.Controls.Remove($Label_ticket_2)
            $Menu_Send_as.Controls.Remove($TextBox_Send_as)
            $Menu_Send_as.Controls.Remove($ComboBox_Send_as)
            $Menu_Send_as.Controls.Remove($Button_Send_as)
            $Text_SendAs_Done = New-Object System.Windows.Forms.Label
            $Text_SendAs_Done.Text = $Outcome
            $Text_SendAs_Done.ForeColor = $Color
            $Text_SendAs_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_SendAs_Done.AutoSize = $true
            $Menu_Send_as.Controls.Add($Text_SendAs_Done)
        }
    )

    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Back'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$EM = ExchangeMenu $Menu_Send_as.Dispose(),$EM})
    $Button_Main_Menu.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Menu_Send_as = New-Object System.Windows.Forms.Form
    $Menu_Send_as.Icon = $Icon
    $Menu_Send_as.BackgroundImage = $Background
    $Menu_Send_as.BackgroundImageLayout = "None"
    $Menu_Send_as.Text ='Exchange account management'
    $Menu_Send_as.Width = 500
    $Menu_Send_as.Height = 500
    $Menu_Send_as.AutoSize = $true
    $Menu_Send_as.Controls.Add($Label_ticket_1)
    $Menu_Send_as.Controls.Add($Label_ticket_2)
    $Menu_Send_as.Controls.Add($TextBox_Send_as)
    $Menu_Send_as.Controls.Add($ComboBox_Send_as)
    $Menu_Send_as.Controls.Add($Button_Send_as)
    $Menu_Send_as.Controls.Add($Button_Main_Menu)
    $Menu_Send_as.Controls.Add($Button_Exit)
    $Menu_Send_as.Controls.Add($Label)
    $Menu_Send_as.StartPosition = 'CenterScreen'
    $Menu_Send_as.MaximizeBox = $false
    $Menu_Send_as.FormBorderStyle = "FixedDialog"

    $ExchangeMenu.Dispose()
    $Menu_Send_as.ShowDialog()
}