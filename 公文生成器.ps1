Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============ 主窗口 ============
$form = New-Object System.Windows.Forms.Form
$form.Text = "中国公文生成器"
$form.Size = New-Object System.Drawing.Size(700, 610)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

# ============ 发文字号 ============
$lblDocPrefix = New-Object System.Windows.Forms.Label
$lblDocPrefix.Location = New-Object System.Drawing.Point(15, 15)
$lblDocPrefix.Size = New-Object System.Drawing.Size(60, 23)
$lblDocPrefix.Text = "发文字号"
$form.Controls.Add($lblDocPrefix)

$txtDocPrefix = New-Object System.Windows.Forms.TextBox
$txtDocPrefix.Location = New-Object System.Drawing.Point(80, 12)
$txtDocPrefix.Size = New-Object System.Drawing.Size(180, 23)
$form.Controls.Add($txtDocPrefix)

$lblLeft = New-Object System.Windows.Forms.Label
$lblLeft.Location = New-Object System.Drawing.Point(265, 15)
$lblLeft.Size = New-Object System.Drawing.Size(15, 23)
$lblLeft.Text = "["
$form.Controls.Add($lblLeft)

$txtDocYear = New-Object System.Windows.Forms.TextBox
$txtDocYear.Location = New-Object System.Drawing.Point(280, 12)
$txtDocYear.Size = New-Object System.Drawing.Size(50, 23)
$txtDocYear.Text = (Get-Date).ToString("yyyy")
$form.Controls.Add($txtDocYear)

$lblRight = New-Object System.Windows.Forms.Label
$lblRight.Location = New-Object System.Drawing.Point(333, 15)
$lblRight.Size = New-Object System.Drawing.Size(15, 23)
$lblRight.Text = "]"
$form.Controls.Add($lblRight)

$txtDocNum = New-Object System.Windows.Forms.TextBox
$txtDocNum.Location = New-Object System.Drawing.Point(348, 12)
$txtDocNum.Size = New-Object System.Drawing.Size(50, 23)
$form.Controls.Add($txtDocNum)

$lblNum = New-Object System.Windows.Forms.Label
$lblNum.Location = New-Object System.Drawing.Point(401, 15)
$lblNum.Size = New-Object System.Drawing.Size(20, 23)
$lblNum.Text = "号"
$form.Controls.Add($lblNum)

$lblSecret = New-Object System.Windows.Forms.Label
$lblSecret.Location = New-Object System.Drawing.Point(460, 15)
$lblSecret.Size = New-Object System.Drawing.Size(40, 23)
$lblSecret.Text = "密级"
$form.Controls.Add($lblSecret)

$comboSecret = New-Object System.Windows.Forms.ComboBox
$comboSecret.Location = New-Object System.Drawing.Point(505, 12)
$comboSecret.Size = New-Object System.Drawing.Size(80, 23)
$comboSecret.DropDownStyle = "DropDownList"
[void]$comboSecret.Items.Add("无")
[void]$comboSecret.Items.Add("秘密")
[void]$comboSecret.Items.Add("机密")
[void]$comboSecret.Items.Add("绝密")
$comboSecret.SelectedIndex = 0
$form.Controls.Add($comboSecret)

# ============ 紧急程度 ============
$lblUrgent = New-Object System.Windows.Forms.Label
$lblUrgent.Location = New-Object System.Drawing.Point(15, 45)
$lblUrgent.Size = New-Object System.Drawing.Size(60, 23)
$lblUrgent.Text = "紧急程度"
$form.Controls.Add($lblUrgent)

$comboUrgent = New-Object System.Windows.Forms.ComboBox
$comboUrgent.Location = New-Object System.Drawing.Point(80, 42)
$comboUrgent.Size = New-Object System.Drawing.Size(80, 23)
$comboUrgent.DropDownStyle = "DropDownList"
[void]$comboUrgent.Items.Add("无")
[void]$comboUrgent.Items.Add("急")
[void]$comboUrgent.Items.Add("特急")
$comboUrgent.SelectedIndex = 0
$form.Controls.Add($comboUrgent)

# ============ 发文机关 ============
$lblSender = New-Object System.Windows.Forms.Label
$lblSender.Location = New-Object System.Drawing.Point(15, 80)
$lblSender.Size = New-Object System.Drawing.Size(60, 23)
$lblSender.Text = "发文机关"
$form.Controls.Add($lblSender)

$txtSender = New-Object System.Windows.Forms.TextBox
$txtSender.Location = New-Object System.Drawing.Point(80, 77)
$txtSender.Size = New-Object System.Drawing.Size(590, 23)
$form.Controls.Add($txtSender)

# ============ 主送机关 ============
$lblReceiver = New-Object System.Windows.Forms.Label
$lblReceiver.Location = New-Object System.Drawing.Point(15, 110)
$lblReceiver.Size = New-Object System.Drawing.Size(60, 23)
$lblReceiver.Text = "主送机关"
$form.Controls.Add($lblReceiver)

$txtReceiver = New-Object System.Windows.Forms.TextBox
$txtReceiver.Location = New-Object System.Drawing.Point(80, 107)
$txtReceiver.Size = New-Object System.Drawing.Size(590, 23)
$form.Controls.Add($txtReceiver)

# ============ 标题 ============
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(15, 145)
$lblTitle.Size = New-Object System.Drawing.Size(60, 23)
$lblTitle.Text = "标题"
$form.Controls.Add($lblTitle)

$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Location = New-Object System.Drawing.Point(80, 142)
$txtTitle.Size = New-Object System.Drawing.Size(590, 23)
$form.Controls.Add($txtTitle)

# ============ 正文 ============
$lblBody = New-Object System.Windows.Forms.Label
$lblBody.Location = New-Object System.Drawing.Point(15, 180)
$lblBody.Size = New-Object System.Drawing.Size(60, 23)
$lblBody.Text = "正文"
$form.Controls.Add($lblBody)

$txtBody = New-Object System.Windows.Forms.RichTextBox
$txtBody.Location = New-Object System.Drawing.Point(80, 177)
$txtBody.Size = New-Object System.Drawing.Size(590, 200)
$txtBody.Font = New-Object System.Drawing.Font("FangSong", 12)
$txtBody.ScrollBars = "Vertical"
$form.Controls.Add($txtBody)

# ============ 落款机关 ============
$lblSigner = New-Object System.Windows.Forms.Label
$lblSigner.Location = New-Object System.Drawing.Point(15, 390)
$lblSigner.Size = New-Object System.Drawing.Size(60, 23)
$lblSigner.Text = "落款机关"
$form.Controls.Add($lblSigner)

$txtSigner = New-Object System.Windows.Forms.TextBox
$txtSigner.Location = New-Object System.Drawing.Point(80, 387)
$txtSigner.Size = New-Object System.Drawing.Size(200, 23)
$form.Controls.Add($txtSigner)

# ============ 发文日期 ============
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Location = New-Object System.Drawing.Point(310, 390)
$lblDate.Size = New-Object System.Drawing.Size(60, 23)
$lblDate.Text = "发文日期"
$form.Controls.Add($lblDate)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(375, 387)
$datePicker.Size = New-Object System.Drawing.Size(200, 23)
$datePicker.Format = "Custom"
$datePicker.CustomFormat = "yyyy年MM月dd日"
$form.Controls.Add($datePicker)

# ============ 附件说明 ============
$lblAttach = New-Object System.Windows.Forms.Label
$lblAttach.Location = New-Object System.Drawing.Point(15, 420)
$lblAttach.Size = New-Object System.Drawing.Size(60, 23)
$lblAttach.Text = "附件说明"
$form.Controls.Add($lblAttach)

$txtAttach = New-Object System.Windows.Forms.TextBox
$txtAttach.Location = New-Object System.Drawing.Point(80, 417)
$txtAttach.Size = New-Object System.Drawing.Size(590, 23)
$form.Controls.Add($txtAttach)

# ============ 抄送机关 ============
$lblContact = New-Object System.Windows.Forms.Label
$lblContact.Location = New-Object System.Drawing.Point(15, 450)
$lblContact.Size = New-Object System.Drawing.Size(60, 23)
$lblContact.Text = "联系人"
$form.Controls.Add($lblContact)

$txtContact = New-Object System.Windows.Forms.TextBox
$txtContact.Location = New-Object System.Drawing.Point(80, 447)
$txtContact.Size = New-Object System.Drawing.Size(200, 23)
$form.Controls.Add($txtContact)

$lblPhone = New-Object System.Windows.Forms.Label
$lblPhone.Location = New-Object System.Drawing.Point(310, 450)
$lblPhone.Size = New-Object System.Drawing.Size(60, 23)
$lblPhone.Text = "联系电话"
$form.Controls.Add($lblPhone)

$txtPhone = New-Object System.Windows.Forms.TextBox
$txtPhone.Location = New-Object System.Drawing.Point(375, 447)
$txtPhone.Size = New-Object System.Drawing.Size(200, 23)
$form.Controls.Add($txtPhone)

$lblCC = New-Object System.Windows.Forms.Label
$lblCC.Location = New-Object System.Drawing.Point(15, 480)
$lblCC.Size = New-Object System.Drawing.Size(60, 23)
$lblCC.Text = "抄送机关"
$form.Controls.Add($lblCC)

$txtCC = New-Object System.Windows.Forms.TextBox
$txtCC.Location = New-Object System.Drawing.Point(80, 477)
$txtCC.Size = New-Object System.Drawing.Size(590, 23)
$form.Controls.Add($txtCC)

# ============ 按钮 ============
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Location = New-Object System.Drawing.Point(220, 520)
$btnGenerate.Size = New-Object System.Drawing.Size(120, 35)
$btnGenerate.Text = "生成公文"
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(200, 30, 30)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = "Flat"
$form.Controls.Add($btnGenerate)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(360, 520)
$btnClear.Size = New-Object System.Drawing.Size(120, 35)
$btnClear.Text = "清空重填"
$btnClear.FlatStyle = "Flat"
$form.Controls.Add($btnClear)

# ============ 生成公文函数 ============
function Generate-Document {
    $docPrefix = $txtDocPrefix.Text
    $docYear = $txtDocYear.Text
    $docNum = $txtDocNum.Text
    $secret = $comboSecret.SelectedItem
    $urgent = $comboUrgent.SelectedItem
    $sender = $txtSender.Text
    $receiver = $txtReceiver.Text
    $title = $txtTitle.Text
    $body = $txtBody.Text
    $signer = $txtSigner.Text
    $dateStr = $datePicker.Value.ToString("yyyy年MM月dd日")
    $attach = $txtAttach.Text
    $contact = $txtContact.Text
    $phone = $txtPhone.Text
    $cc = $txtCC.Text

    # 验证必填项
    if ([string]::IsNullOrWhiteSpace($sender)) {
        [System.Windows.Forms.MessageBox]::Show("请填写发文机关！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($receiver)) {
        [System.Windows.Forms.MessageBox]::Show("请填写主送机关！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($title)) {
        [System.Windows.Forms.MessageBox]::Show("请填写公文标题！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($body)) {
        [System.Windows.Forms.MessageBox]::Show("请填写公文正文！", "提示", "OK", "Warning")
        return
    }

    # 构建文号
    $docNumStr = ""
    if (-not [string]::IsNullOrWhiteSpace($docPrefix) -and -not [string]::IsNullOrWhiteSpace($docNum)) {
        $docNumStr = "$docPrefix[$docYear]$docNum号"
    }

    # 处理正文换行和段落缩进
    $bodyHtml = $body -replace "`r`n", "`n"
    $paragraphs = $bodyHtml -split "`n"
    $bodyHtml = ""
    foreach ($para in $paragraphs) {
        if (-not [string]::IsNullOrWhiteSpace($para)) {
            $bodyHtml += "<p>$para</p>`n"
        }
    }

    # 生成HTML
    $html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>$title</title>
<style>
@page { size: A4; margin: 0; }
body { font-family: FangSong, 'FangSong_GB2312', SimSun, serif; font-size: 16pt; line-height: 28pt; color: #000; margin: 0; padding: 0; }
.page { width: 210mm; min-height: 297mm; margin: 0; padding: 37mm 26mm 35mm 28mm; box-sizing: border-box; }
@media print {
    body { margin: 0; padding: 0; }
    .page { margin: 0; padding: 37mm 26mm 35mm 28mm; }
}
.header { text-align: center; margin-bottom: 20pt; }
.secret-urgent { text-align: left; font-size: 16pt; margin-bottom: 10pt; }
.secret-tag { color: #000; font-weight: bold; margin-right: 20pt; }
.urgent-tag { color: #000; font-weight: bold; }
.sender-name { font-family: 'FZXiaoBiaoSong-B05S', 'SimSun', serif; font-size: 36pt; color: #FF0000; font-weight: bold; letter-spacing: 2pt; margin: 20pt 0; white-space: nowrap; width: fit-content; max-width: 100%; transform-origin: center; }
.red-line { border: none; margin: 15pt 0; position: relative; height: 4px; }
.red-line::before { content: ''; position: absolute; top: 0; left: 0; right: 0; border-top: 1px solid #FF0000; }
.red-line::after { content: ''; position: absolute; top: 3px; left: 0; right: 0; border-top: 3px solid #FF0000; }
.doc-number { font-size: 16pt; text-align: right; margin: 10pt 0; }
.receiver { font-size: 16pt; margin: 20pt 0 10pt 0; }
.title { font-family: 'FZXiaoBiaoSong-B05S', 'SimSun', serif; font-size: 22pt; text-align: center; margin: 20pt 0; font-weight: bold; }
.body-text { font-size: 16pt; margin: 10pt 0; line-height: 28pt; }
.body-text p { text-indent: 2em; margin: 0; }
.footer { margin-top: 40pt; }
.signer { text-align: right; font-size: 16pt; margin: 5pt 0; }
.date { text-align: right; font-size: 16pt; margin: 5pt 0; }
.attachment { font-size: 14pt; margin: 20pt 0 10pt 0; border-top: 1px solid #000; padding-top: 10pt; font-family: FangSong, 'FangSong_GB2312', SimSun, serif; }
.cc { font-size: 14pt; margin: 10pt 0; border-top: 1px solid #000; padding-top: 10pt; font-family: FangSong, 'FangSong_GB2312', SimSun, serif; }
.contact-info { font-size: 14pt; margin: 10pt 0; font-family: FangSong, 'FangSong_GB2312', SimSun, serif; }
</style>
</head>
<body>
<div class="page">
    <div class="header">
"@

    if ($secret -ne "无" -or $urgent -ne "无") {
        $html += "`n        <div class=`"secret-urgent`">`n"
        if ($secret -ne "无") {
            $html += "            <span class=`"secret-tag`">★ $secret</span>`n"
        }
        if ($urgent -ne "无") {
            $html += "            <span class=`"urgent-tag`">[$urgent]</span>`n"
        }
        $html += "        </div>`n"
    }

    $html += @"
        <div class="sender-name" id="senderName">$sender</div>
        <script>
        (function(){
            var el = document.getElementById('senderName');
            var maxWidth = el.parentElement.offsetWidth - 20;
            var fontSize = 36;
            el.style.fontSize = fontSize + 'pt';
            while (el.scrollWidth > maxWidth && fontSize > 12) {
                fontSize -= 0.5;
                el.style.fontSize = fontSize + 'pt';
            }
        })();
        </script>
        <hr class="red-line">
"@

    if ($docNumStr) {
        $html += "        <div class=`"doc-number`">$docNumStr</div>`n"
    }

    $html += "    </div>`n`n"
    $html += "    <div class=`"title`">$title</div>`n`n"
    $html += "    <div class=`"receiver`">$receiver：</div>`n`n"
    $html += "    <div class=`"body-text`">$bodyHtml</div>`n`n"
    $html += "    <div class=`"footer`">`n"
    $html += "        <div class=`"signer`">$signer</div>`n"
    $html += "        <div class=`"date`">$dateStr</div>`n"

    if (-not [string]::IsNullOrWhiteSpace($attach)) {
        $html += "        <div class=`"attachment`">附件：$attach</div>`n"
    }
    if (-not [string]::IsNullOrWhiteSpace($contact) -or -not [string]::IsNullOrWhiteSpace($phone)) {
        $html += "        <div class=`"contact-info`">`n"
        if (-not [string]::IsNullOrWhiteSpace($contact)) {
            $html += "            <span>联系人：$contact</span>`n"
        }
        if (-not [string]::IsNullOrWhiteSpace($phone)) {
            $html += "            <span style=`"margin-left: 30pt;`">联系电话：$phone</span>`n"
        }
        $html += "        </div>`n"
    }
    if (-not [string]::IsNullOrWhiteSpace($cc)) {
        $html += "        <div class=`"cc`">抄送：$cc</div>`n"
    }

    $html += "    </div>`n</div>`n</body>`n</html>"

    # 保存文件
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "HTML文件|*.html|所有文件|*.*"
    $saveDialog.FileName = "公文_$title"
    $saveDialog.DefaultExt = "html"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $html | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("公文已生成！`n`n文件位置：$($saveDialog.FileName)", "完成", "OK", "Information")
        Start-Process $saveDialog.FileName
    }
}

# ============ 清空表单 ============
function Clear-Form {
    $txtDocPrefix.Text = ""
    $txtDocYear.Text = (Get-Date).ToString("yyyy")
    $txtDocNum.Text = ""
    $comboSecret.SelectedIndex = 0
    $comboUrgent.SelectedIndex = 0
    $txtSender.Text = ""
    $txtReceiver.Text = ""
    $txtTitle.Text = ""
    $txtBody.Text = ""
    $txtSigner.Text = ""
    $datePicker.Value = (Get-Date)
    $txtAttach.Text = ""
    $txtContact.Text = ""
    $txtPhone.Text = ""
    $txtCC.Text = ""
}

# ============ 事件绑定 ============
$btnGenerate.Add_Click({ Generate-Document })
$btnClear.Add_Click({ Clear-Form })

# ============ 启动 ============
[System.Windows.Forms.Application]::Run($form)
