function Cat {
    Add-Type -assembly System.Windows.Forms

    $PictureBox_For_Cat = New-Object System.Windows.Forms.PictureBox
    $PictureBox_For_Cat.Load($Cat)
    $PictureBox_For_Cat.Location = New-Object System.Drawing.Point(10,10)
    $PictureBox_For_Cat.AutoSize = $true

    $Button_Back = New-Object System.Windows.Forms.Button
    $Button_Back.Text = 'Back'
    $Button_Back.Location = New-Object System.Drawing.Point(10,400)
    $Button_Back.Add_click({$AD = ActiveDirectoryMenu $Form_For_Kuzia.Dispose(),$AD})
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

    $Form_For_Cat = New-Object System.Windows.Forms.Form
    $Form_For_Cat.Icon = $Icon
    $Form_For_Cat.Text ='Kittens go!'
    $Form_For_Cat.Width = 500
    $Form_For_Cat.Height = 500
    $Form_For_Cat.AutoSize = $true
    $Form_For_Cat.Controls.Add($Label)
    $Form_For_Cat.Controls.add($PictureBox_For_Cat)
    $Form_For_Cat.Controls.Add($Button_Exit)
    $Form_For_Cat.Controls.Add($Button_Back)
    $Form_For_Cat.StartPosition = 'CenterScreen'
    $Form_For_Cat.MaximizeBox = $false
    $Form_For_Cat.FormBorderStyle = "FixedDialog"

    $ActiveDirectoryMenu.Dispose()
    $Form_For_Cat.ShowDialog()
}