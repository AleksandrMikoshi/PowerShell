<#Menu function for specifying data for further mailbox creation#>
<#Функция меню для указания данных для дальнейшего создания почтового ящика#>
function Create_mail {
    Add-Type -assembly System.Windows.Forms
    Label
    Button_Exit 

    $Label_Create_mail = New-Object System.Windows.Forms.Label
    $Label_Create_mail.Text = "Enter information to create a mailbox"
    $Label_Create_mail.Location = New-Object System.Drawing.Point(10,20)
    $Label_Create_mail.AutoSize = $true

    $Label_ticket_1 = New-Object System.Windows.Forms.Label
    $Label_ticket_1.Text = "Enter mailbox name"
    $Label_ticket_1.Location = New-Object System.Drawing.Point(10,50)
    $Label_ticket_1.AutoSize = $true

    $TextBox_mail = New-Object System.Windows.Forms.TextBox
    $TextBox_mail.Location  = New-Object System.Drawing.Point(220,50)
    $TextBox_mail.Text = ""
    $TextBox_mail.Size = New-Object System.Drawing.Size(250,30)

    $Label_ticket_2 = New-Object System.Windows.Forms.Label
    $Label_ticket_2.Text = "Enter the mailbox address`nwith the domain name"
    $Label_ticket_2.Location = New-Object System.Drawing.Point(10,75)
    $Label_ticket_2.AutoSize = $true

    $TextBox_addr = New-Object System.Windows.Forms.TextBox
    $TextBox_addr.Location  = New-Object System.Drawing.Point(220,75)
    $TextBox_addr.Text = ""
    $TextBox_addr.Size = New-Object System.Drawing.Size(250,30)

    $Button_Create_Mail = New-Object System.Windows.Forms.Button
    $Button_Create_Mail.Location = New-Object System.Drawing.Size(330,20)
    $Button_Create_Mail.AutoSize = $true
    $Button_Create_Mail.Text = "Create mailbox"
    $Button_Create_Mail.Add_click(
        {
            $leght = Get-Random -Minimum 8 -Maximum 12
            $generate = [System.Web.Security.Membership]::GeneratePassword($leght,2)
            $MailAttr = @{
                Name = $TextBox_mail.Text
                UserPrincipalName = $TextBox_addr.Text
                Initials = $TextBox_Initials.Text
                SamAccountName = $TextBox_SamAccountName.Text
                password_mail = ConvertTo-SecureString $generate -AsPlainText -Force
            }
            $MailBox = Create_MailBox -MailAttr $MailAttr -DC $DC -Path_OU $Path_Mail -Mail_serv $Mail_serv -Path_log $Path_log -URL_auth $URL_auth -API_key_Passwork $API_key_Passwork -URL_data_passwords $URL_data_passwords -Case_Passwork $Case_Passwork
            $Form_Mail_create.Controls.Remove($Label_Create_mail)
            $Form_Mail_create.Controls.Remove($Label_ticket_1)
            $Form_Mail_create.Controls.Remove($TextBox_mail)
            $Form_Mail_create.Controls.Remove($Label_ticket_2)
            $Form_Mail_create.Controls.Remove($TextBox_addr)
            $Form_Mail_create.Controls.Remove($Button_Create_Mail)
            [string]$Color = $MailBox.Color
            [string]$Outcome = $MailBox.Outcome
            $Text_MailBox_Done = New-Object System.Windows.Forms.Label
            $Text_MailBox_Done.Text = $Outcome
            $Text_MailBox_Done.ForeColor = $Color
            $Text_MailBox_Done.Location = New-Object System.Drawing.Point(10,20)
            $Text_MailBox_Done.AutoSize = $true
            $Form_Mail_create.Controls.Add($Text_MailBox_Done)
        }
    )
            
    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$EM = ExchangeMenu $Form_Mail_create.Dispose(),$EM})
    $Button_Back.AutoSize = $true

    $Background = [system.drawing.image]::FromFile($Background_Image)
    $Form_Mail_create = New-Object System.Windows.Forms.Form
    $Form_Mail_create.Icon = $Icon
    $Form_Mail_create.BackgroundImage = $Background
    $Form_Mail_create.BackgroundImageLayout = "None"
    $Form_Mail_create.Text ='Active Directory Account Management'
    $Form_Mail_create.Width = 500
    $Form_Mail_create.Height = 500
    $Form_Mail_create.AutoSize = $true
    $Form_Mail_create.Controls.Add($Label_Create_mail)
    $Form_Mail_create.Controls.Add($Label_ticket_1)
    $Form_Mail_create.Controls.Add($TextBox_mail)
    $Form_Mail_create.Controls.Add($Label_ticket_2)
    $Form_Mail_create.Controls.Add($TextBox_addr)
    $Form_Mail_create.Controls.Add($Button_Create_Mail)
    $Form_Mail_create.Controls.Add($Button_Back)
    $Form_Mail_create.Controls.Add($Button_Exit)
    $Form_Mail_create.Controls.Add($Label)
    $Form_Mail_create.StartPosition = 'CenterScreen'
    $Form_Mail_create.MaximizeBox = $false
    $Form_Mail_create.FormBorderStyle = "FixedDialog"

    $ExchangeMenu.Dispose()
    $Form_Mail_create.ShowDialog()
}