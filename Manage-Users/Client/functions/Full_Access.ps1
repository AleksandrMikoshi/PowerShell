<#Menu function for specifying data for further granting of rights#>
<#Функция меню для указания данных для дальнейшего предоставления прав#>
function Full_Access {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $ComboBox_Full_Access = New-Object System.Windows.Forms.CheckedListBox
    $ComboBox_Full_Access.Width = 250
    $ComboBox_Full_Access.Height = 300
    $Users = get-aduser -filter * -Properties cn -SearchBase $Path_OU | sort cn
    Foreach ($User in $Users){
        $ComboBox_Full_Access.Items.Add($User.cn);
    }
    $ComboBox_Full_Access.Location = New-Object System.Drawing.Point(10,90)

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Enter mailbox name"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,20)
    $Label_ticket_1.AutoSize = $true

    $TextBox_Full_Access = New-Object System.Windows.Forms.TextBox
    $TextBox_Full_Access.Location  = New-Object System.Drawing.Point(10,40)
    $TextBox_Full_Access.Text = ""
    $TextBox_Full_Access.Size = New-Object System.Drawing.Size(250,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "Select the users to whom you want to grant the 'Full Access' right"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,70)
    $Label_ticket_2.AutoSize = $true


    $Button_Full_Access = New-Object System.Windows.Forms.Button
    $Button_Full_Access.Location = New-Object System.Drawing.Size(350,40)
    $Button_Full_Access.AutoSize = $true
    $Button_Full_Access.Text = "grant rights"
    $Button_Full_Access.Add_Click(
        { 
            $Mail = $TextBox_Full_Access.Text
            $Users = $ComboBox_Full_Access.CheckedItems
            $FA = Full_Access_JEA -Mail $Mail -Users $Users -Mail_serv $Mail_serv -Path_log $Path_log
            $Menu_Full_Access.Controls.Remove($Label_ticket_1)
            $Menu_Full_Access.Controls.Remove($Label_ticket_2)
            $Menu_Full_Access.Controls.Remove($TextBox_Full_Access)
            $Menu_Full_Access.Controls.Remove($ComboBox_Full_Access)
            $Menu_Full_Access.Controls.Remove($Button_Full_Access)
            [string]$Color = $FA.Color
            [string]$Outcome = $FA.Outcome
            $Text_Full_Access_Done = New-Object System.Windows.Forms.Label
            $Text_Full_Access_Done.Text = $Outcome
            $Text_Full_Access_Done.ForeColor = $Color
            $Text_Full_Access_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_Full_Access_Done.AutoSize = $true
            $Menu_Full_Access.Controls.Add($Text_Full_Access_Done)
        }
    )

    $Button_Main_Menu = New-Object System.Windows.Forms.Button
    $Button_Main_Menu.Text = 'Back'
    $Button_Main_Menu.Location = New-Object System.Drawing.Point(10,400)
    $Button_Main_Menu.Add_click({$EM = ExchangeMenu $Menu_Full_Access.Dispose(),$EM})
    $Button_Main_Menu.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Menu_Full_Access = New-Object System.Windows.Forms.Form
    $Menu_Full_Access.Icon = $Icon
    $Menu_Full_Access.BackgroundImage = $Background
    $Menu_Full_Access.BackgroundImageLayout = "None"
    $Menu_Full_Access.Text ='Exchange account management'
    $Menu_Full_Access.Width = 500
    $Menu_Full_Access.Height = 500
    $Menu_Full_Access.AutoSize = $true
    $Menu_Full_Access.Controls.Add($Label_ticket_1)
    $Menu_Full_Access.Controls.Add($Label_ticket_2)
    $Menu_Full_Access.Controls.Add($TextBox_Full_Access)
    $Menu_Full_Access.Controls.Add($ComboBox_Full_Access)
    $Menu_Full_Access.Controls.Add($Button_Full_Access)
    $Menu_Full_Access.Controls.Add($Button_Main_Menu)
    $Menu_Full_Access.Controls.Add($Button_Exit)
    $Menu_Full_Access.Controls.Add($Label)
    $Menu_Full_Access.StartPosition = 'CenterScreen'
    $Menu_Full_Access.MaximizeBox = $false
    $Menu_Full_Access.FormBorderStyle = "FixedDialog"

    $ExchangeMenu.Dispose()
    $Menu_Full_Access.ShowDialog()
}