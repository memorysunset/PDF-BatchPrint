Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============ 主窗口 ============
$form = New-Object System.Windows.Forms.Form
$form.Text = "中国公文生成器"
$form.Size = New-Object System.Drawing.Size(700, 820)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

# ============ UI辅助函数 ============
function Add-UIControl {
    param([string]$Type, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text = "")
    $ctrl = New-Object "System.Windows.Forms.$Type"
    $ctrl.Location = New-Object System.Drawing.Point($X, $Y)
    $ctrl.Size = New-Object System.Drawing.Size($W, $H)
    if ($Text) { $ctrl.Text = $Text }
    $form.Controls.Add($ctrl)
    return $ctrl
}

# ============ 基本信息区域 ============
$groupBasic = Add-UIControl -Type GroupBox -X 15 -Y 10 -W 655 -H 90 -Text "基本信息"

$labelDocNum = Add-UIControl -Type Label -X 25 -Y 30 -W 35 -H 23 -Text "文号"
$txtDocPrefix = Add-UIControl -Type TextBox -X 65 -Y 27 -W 150 -H 23
$txtDocPrefix.Text = ""
$lblYear = Add-UIControl -Type Label -X 220 -Y 30 -W 20 -H 23 -Text "〔"
$txtDocYear = Add-UIControl -Type TextBox -X 240 -Y 27 -W 50 -H 23
$txtDocYear.Text = (Get-Date).ToString("yyyy")
$lblYearEnd = Add-UIControl -Type Label -X 293 -Y 30 -W 15 -H 23 -Text "〕"
$txtDocNum = Add-UIControl -Type TextBox -X 310 -Y 27 -W 50 -H 23
$lblNum = Add-UIControl -Type Label -X 363 -Y 30 -W 20 -H 23 -Text "号"

$labelSecret = Add-UIControl -Type Label -X 420 -Y 30 -W 40 -H 23 -Text "密级"
$comboSecret = Add-UIControl -Type ComboBox -X 465 -Y 27 -W 80 -H 23
$comboSecret.DropDownStyle = "DropDownList"
$comboSecret.Items.AddRange(@("无", "秘密", "机密", "绝密"))
$comboSecret.SelectedIndex = 0

$labelUrgent = Add-UIControl -Type Label -X 25 -Y 60 -W 60 -H 23 -Text "紧急程度"
$comboUrgent = Add-UIControl -Type ComboBox -X 90 -Y 57 -W 80 -H 23
$comboUrgent.DropDownStyle = "DropDownList"
$comboUrgent.Items.AddRange(@("无", "急", "特急"))
$comboUrgent.SelectedIndex = 0

# ============ 发文信息区域 ============
$groupSender = Add-UIControl -Type GroupBox -X 15 -Y 105 -W 655 -H 90 -Text "发文信息"

$labelSender = Add-UIControl -Type Label -X 25 -Y 128 -W 60 -H 23 -Text "发文机关"
$txtSender = Add-UIControl -Type TextBox -X 90 -Y 125 -W 555 -H 23
$txtSender.Text = ""

$labelReceiver = Add-UIControl -Type Label -X 25 -Y 158 -W 60 -H 23 -Text "主送机关"
$txtReceiver = Add-UIControl -Type TextBox -X 90 -Y 155 -W 555 -H 23
$txtReceiver.Text = ""

# ============ 公文内容区域 ============
$groupContent = Add-UIControl -Type GroupBox -X 15 -Y 200 -W 655 -H 330 -Text "公文内容"

$labelTitle = Add-UIControl -Type Label -X 25 -Y 223 -W 40 -H 23 -Text "标题"
$txtTitle = Add-UIControl -Type TextBox -X 70 -Y 220 -W 575 -H 23
$txtTitle.Text = ""

$labelBody = Add-UIControl -Type Label -X 25 -Y 253 -W 40 -H 23 -Text "正文"
$txtBody = Add-UIControl -Type RichTextBox -X 70 -Y 250 -W 575 -H 265
$txtBody.Text = ""
$txtBody.Font = New-Object System.Drawing.Font("FangSong", 12)
$txtBody.ScrollBars = "Vertical"

# ============ 落款信息区域 ============
$groupFooter = Add-UIControl -Type GroupBox -X 15 -Y 535 -W 655 -H 90 -Text "落款信息"

$labelSigner = Add-UIControl -Type Label -X 25 -Y 558 -W 60 -H 23 -Text "落款机关"
$txtSigner = Add-UIControl -Type TextBox -X 90 -Y 555 -W 200 -H 23
$txtSigner.Text = ""

$labelDate = Add-UIControl -Type Label -X 320 -Y 558 -W 60 -H 23 -Text "发文日期"
$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(385, 555)
$datePicker.Size = New-Object System.Drawing.Size(200, 23)
$datePicker.Format = "Custom"
$datePicker.CustomFormat = "yyyy年MM月dd日"
$form.Controls.Add($datePicker)

$labelAttach = Add-UIControl -Type Label -X 25 -Y 588 -W 60 -H 23 -Text "附件说明"
$txtAttach = Add-UIControl -Type TextBox -X 90 -Y 585 -W 555 -H 23
$txtAttach.Text = ""

# ============ 抄送区域 ============
$groupCC = Add-UIControl -Type GroupBox -X 15 -Y 630 -W 655 -H 60 -Text "抄送"

$labelCC = Add-UIControl -Type Label -X 25 -Y 653 -W 60 -H 23 -Text "抄送机关"
$txtCC = Add-UIControl -Type TextBox -X 90 -Y 650 -W 555 -H 23
$txtCC.Text = ""

# ============ 按钮区域 ============
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Location = New-Object System.Drawing.Point(220, 700)
$btnGenerate.Size = New-Object System.Drawing.Size(120, 35)
$btnGenerate.Text = "生成公文"
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(200, 30, 30)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = "Flat"
$form.Controls.Add($btnGenerate)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(360, 700)
$btnClear.Size = New-Object System.Drawing.Size(120, 35)
$btnClear.Text = "清空重填"
$btnClear.FlatStyle = "Flat"
$form.Controls.Add($btnClear)

# ============ 生成公文函数 ============
function Generate-Document {
    $data = @{
        DocPrefix  = $txtDocPrefix.Text
        DocYear    = $txtDocYear.Text
        DocNum     = $txtDocNum.Text
        Secret     = $comboSecret.SelectedItem
        Urgent     = $comboUrgent.SelectedItem
        Sender     = $txtSender.Text
        Receiver   = $txtReceiver.Text
        Title      = $txtTitle.Text
        Body       = $txtBody.Text
        Signer     = $txtSigner.Text
        Date       = $datePicker.Value.ToString("yyyy年MM月dd日")
        Attachment = $txtAttach.Text
        CC         = $txtCC.Text
    }

    # 验证必填项
    if ([string]::IsNullOrWhiteSpace($data.Sender)) {
        [System.Windows.Forms.MessageBox]::Show("请填写发文机关！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($data.Receiver)) {
        [System.Windows.Forms.MessageBox]::Show("请填写主送机关！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($data.Title)) {
        [System.Windows.Forms.MessageBox]::Show("请填写公文标题！", "提示", "OK", "Warning")
        return
    }
    if ([string]::IsNullOrWhiteSpace($data.Body)) {
        [System.Windows.Forms.MessageBox]::Show("请填写公文正文！", "提示", "OK", "Warning")
        return
    }

    # 构建文号
    $docNumStr = ""
    if (-not [string]::IsNullOrWhiteSpace($data.DocPrefix) -and -not [string]::IsNullOrWhiteSpace($data.DocNum)) {
        $docNumStr = "$($data.DocPrefix)〔$($data.DocYear)〕$($data.DocNum)号"
    }

    # 处理正文换行
    $bodyHtml = $data.Body -replace "`r`n", "<br>" -replace "`n", "<br>"

    # 生成HTML公文
    $html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>$($data.Title)</title>
<style>
@page {
    size: A4;
    margin: 37mm 26mm 35mm 28mm;
}
body {
    font-family: FangSong, 仿宋, SimSun, 宋体;
    font-size: 16pt;
    line-height: 28pt;
    color: #000000;
    margin: 0;
    padding: 0;
}
.page {
    width: 210mm;
    min-height: 297mm;
    margin: 0 auto;
    padding: 37mm 26mm 35mm 28mm;
    box-sizing: border-box;
}
/* 版头区域 */
.header {
    text-align: center;
    margin-bottom: 20pt;
    position: relative;
}
.secret-urgent {
    text-align: left;
    font-size: 16pt;
    margin-bottom: 10pt;
}
.secret-tag {
    color: #000;
    font-weight: bold;
    margin-right: 20pt;
}
.urgent-tag {
    color: #000;
    font-weight: bold;
}
.sender-name {
    font-family: SimHei, 黑体;
    font-size: 36pt;
    color: #FF0000;
    font-weight: bold;
    letter-spacing: 2pt;
    margin: 20pt 0;
}
.red-line {
    border: none;
    border-top: 2px solid #FF0000;
    margin: 15pt 0;
}
.doc-number {
    font-size: 16pt;
    text-align: center;
    margin: 10pt 0;
}
/* 正文区域 */
.receiver {
    font-size: 16pt;
    margin: 20pt 0 10pt 0;
}
.title {
    font-family: SimHei, 黑体;
    font-size: 22pt;
    text-align: center;
    margin: 20pt 0;
    font-weight: bold;
}
.body-text {
    font-size: 16pt;
    text-indent: 2em;
    margin: 10pt 0;
    line-height: 28pt;
}
/* 版记区域 */
.footer {
    margin-top: 40pt;
}
.signer {
    text-align: right;
    font-size: 16pt;
    margin: 5pt 0;
}
.date {
    text-align: right;
    font-size: 16pt;
    margin: 5pt 0;
}
.attachment {
    font-size: 16pt;
    margin: 20pt 0 10pt 0;
    border-top: 1px solid #000;
    padding-top: 10pt;
}
.cc {
    font-size: 16pt;
    margin: 10pt 0;
    border-top: 1px solid #000;
    padding-top: 10pt;
}
</style>
</head>
<body>
<div class="page">
    <!-- 版头 -->
    <div class="header">
"@

    # 密级和紧急程度
    if ($data.Secret -ne "无" -or $data.Urgent -ne "无") {
        $html += "`n        <div class=`"secret-urgent`">`n"
        if ($data.Secret -ne "无") {
            $html += "            <span class=`"secret-tag`">★ $data.Secret</span>`n"
        }
        if ($data.Urgent -ne "无") {
            $html += "            <span class=`"urgent-tag`">[$data.Urgent]</span>`n"
        }
        $html += "        </div>`n"
    }

    # 发文机关标志
    $html += @"
        <div class="sender-name">$($data.Sender)</div>
        <hr class="red-line">
"@

    # 文号
    if ($docNumStr) {
        $html += "        <div class=`"doc-number`">$docNumStr</div>`n"
    }

    $html += "    </div>`n`n"

    # 主送机关
    $html += "    <div class=`"receiver`">$($data.Receiver)：</div>`n`n"

    # 标题
    $html += "    <div class=`"title`">$($data.Title)</div>`n`n"

    # 正文
    $html += "    <div class=`"body-text`">$bodyHtml</div>`n`n"

    # 版记
    $html += "    <div class=`"footer`">`n"

    # 发文机关和日期
    $html += "        <div class=`"signer`">$($data.Signer)</div>`n"
    $html += "        <div class=`"date`">$($data.Date)</div>`n"

    # 附件
    if (-not [string]::IsNullOrWhiteSpace($data.Attachment)) {
        $html += "        <div class=`"attachment`">附件：$($data.Attachment)</div>`n"
    }

    # 抄送
    if (-not [string]::IsNullOrWhiteSpace($data.CC)) {
        $html += "        <div class=`"cc`">抄送：$($data.CC)</div>`n"
    }

    $html += "    </div>`n"
    $html += "</div>`n</body>`n</html>"

    # 保存文件
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "HTML文件|*.html|所有文件|*.*"
    $saveDialog.FileName = "公文_$($data.Title)"
    $saveDialog.DefaultExt = "html"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $html | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("公文已生成！`n`n文件位置：$($saveDialog.FileName)", "完成", "OK", "Information")

        # 打开文件
        Start-Process $saveDialog.FileName
    }
}

# ============ 清空表单函数 ============
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
    $txtCC.Text = ""
}

# ============ 事件绑定 ============
$btnGenerate.Add_Click({ Generate-Document })
$btnClear.Add_Click({ Clear-Form })

# ============ 启动 ============
[System.Windows.Forms.Application]::Run($form)
