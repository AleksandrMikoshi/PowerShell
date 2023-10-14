function Button_Exit {
    $global:Button_Exit = New-Object System.Windows.Forms.Button
    $global:Button_Exit.Text = 'Exit'
    $global:Button_Exit.Location = New-Object System.Drawing.Point(400,400)
    $global:Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $global:Button_Exit.AutoSize = $true
    $global:Button_Exit.BackColor = "$Button_Color"
    $global:Button_Exit.ForeColor = "$Text_Color"
} 