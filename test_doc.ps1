Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "测试公文输入"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 20)
$label1.Size = New-Object System.Drawing.Size(80, 23)
$label1.Text = "发文机关"
$form.Controls.Add($label1)

$txt1 = New-Object System.Windows.Forms.TextBox
$txt1.Location = New-Object System.Drawing.Point(100, 17)
$txt1.Size = New-Object System.Drawing.Size(350, 23)
$txt1.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10)
$form.Controls.Add($txt1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20, 60)
$label2.Size = New-Object System.Drawing.Size(80, 23)
$label2.Text = "主送机关"
$form.Controls.Add($label2)

$txt2 = New-Object System.Windows.Forms.TextBox
$txt2.Location = New-Object System.Drawing.Point(100, 57)
$txt2.Size = New-Object System.Drawing.Size(350, 23)
$txt2.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10)
$form.Controls.Add($txt2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(20, 100)
$label3.Size = New-Object System.Drawing.Size(80, 23)
$label3.Text = "标题"
$form.Controls.Add($label3)

$txt3 = New-Object System.Windows.Forms.TextBox
$txt3.Location = New-Object System.Drawing.Point(100, 97)
$txt3.Size = New-Object System.Drawing.Size(350, 23)
$txt3.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10)
$form.Controls.Add($txt3)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(20, 140)
$label4.Size = New-Object System.Drawing.Size(80, 23)
$label4.Text = "正文"
$form.Controls.Add($label4)

$txt4 = New-Object System.Windows.Forms.RichTextBox
$txt4.Location = New-Object System.Drawing.Point(100, 137)
$txt4.Size = New-Object System.Drawing.Size(350, 150)
$txt4.Font = New-Object System.Drawing.Font("FangSong", 12)
$txt4.ScrollBars = "Vertical"
$form.Controls.Add($txt4)

$btn = New-Object System.Windows.Forms.Button
$btn.Location = New-Object System.Drawing.Point(200, 310)
$btn.Size = New-Object System.Drawing.Size(100, 30)
$btn.Text = "测试"
$btn.Add_Click({
    $msg = "发文机关: " + $txt1.Text + "`n"
    $msg += "主送机关: " + $txt2.Text + "`n"
    $msg += "标题: " + $txt3.Text + "`n"
    $msg += "正文: " + $txt4.Text
    [System.Windows.Forms.MessageBox]::Show($msg, "测试结果")
})
$form.Controls.Add($btn)

[System.Windows.Forms.Application]::Run($form)
