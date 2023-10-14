function Label{
    $global:Label = New-Object System.Windows.Forms.Label
    $global:Label.Text = $Label_Text
    $global:Label.Location  = New-Object System.Drawing.Point(110,450)
    $global:Label.AutoSize = $true
    $global:Label.BackColor = $Label_Color
    $global:Label.ForeColor = $Text_Label
}