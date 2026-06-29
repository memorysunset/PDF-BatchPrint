Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "测试输入"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

$txt = New-Object System.Windows.Forms.TextBox
$txt.Location = New-Object System.Drawing.Point(50, 50)
$txt.Size = New-Object System.Drawing.Size(300, 30)
$txt.Font = New-Object System.Drawing.Font("Microsoft YaHei", 12)
$form.Controls.Add($txt)

$btn = New-Object System.Windows.Forms.Button
$btn.Location = New-Object System.Drawing.Point(150, 100)
$btn.Size = New-Object System.Drawing.Size(100, 30)
$btn.Text = "确定"
$btn.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("你输入了: " + $txt.Text)
})
$form.Controls.Add($btn)

[System.Windows.Forms.Application]::Run($form)
