Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============ 配置常量 ============
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:SumatraPath = Join-Path $script:ScriptDir "SumatraPDF\SumatraPDF.exe"
$script:JobDelayMs = 500

# ============ 主窗口 ============
$form = New-Object System.Windows.Forms.Form
$form.Text = "PDF 批量打印工具"
$form.Size = New-Object System.Drawing.Size(600, 520)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

# ============ UI 辅助函数 ============
function Add-UIControl {
    param([string]$Type, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text = "")
    $ctrl = New-Object "System.Windows.Forms.$Type"
    $ctrl.Location = New-Object System.Drawing.Point($X, $Y)
    $ctrl.Size = New-Object System.Drawing.Size($W, $H)
    if ($Text) { $ctrl.Text = $Text }
    $form.Controls.Add($ctrl)
    return $ctrl
}

function Set-UIForPrinting {
    param([bool]$Printing)
    $btnPrint.Enabled = -not $Printing
    $btnCancel.Enabled = $Printing
    $btnBrowse.Enabled = -not $Printing
    $comboPrinter.Enabled = -not $Printing
    $radioColor.Enabled = -not $Printing
    $radioBW.Enabled = -not $Printing
}

function Set-AllChecked {
    param([bool]$Checked)
    foreach ($item in $listView.Items) { $item.Checked = $Checked }
    Update-FileCount
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    else { return "{0:N1} KB" -f ($Bytes / 1KB) }
}

# ============ 界面控件 ============
$labelFolder = Add-UIControl -Type Label -X 15 -Y 18 -W 80 -H 23 -Text "文件夹路径："
$textFolder = Add-UIControl -Type TextBox -X 95 -Y 15 -W 380 -H 23
$textFolder.ReadOnly = $true
$btnBrowse = Add-UIControl -Type Button -X 485 -Y 13 -W 90 -H 27 -Text "选择文件夹"

$labelPrinter = Add-UIControl -Type Label -X 15 -Y 50 -W 60 -H 23 -Text "打印机："
$comboPrinter = Add-UIControl -Type ComboBox -X 95 -Y 47 -W 380 -H 23
$comboPrinter.DropDownStyle = "DropDownList"

# 加载系统打印机
$printers = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
$comboPrinter.Items.AddRange($printers)
if ($comboPrinter.Items.Count -gt 0) {
    $defaultPrinter = (New-Object System.Drawing.Printing.PrinterSettings).PrinterName
    $idx = $comboPrinter.Items.IndexOf($defaultPrinter)
    $comboPrinter.SelectedIndex = [Math]::Max($idx, 0)
}

$btnSelectAll = Add-UIControl -Type Button -X 15 -Y 80 -W 60 -H 25 -Text "全选"
$btnDeselectAll = Add-UIControl -Type Button -X 80 -Y 80 -W 70 -H 25 -Text "取消全选"
$labelCount = Add-UIControl -Type Label -X 165 -Y 84 -W 100 -H 23 -Text "已选 0 个文件"
$labelCount.ForeColor = [System.Drawing.Color]::Gray

$labelColorMode = Add-UIControl -Type Label -X 280 -Y 84 -W 45 -H 23 -Text "颜色："
$radioColor = Add-UIControl -Type RadioButton -X 325 -Y 82 -W 55 -H 25 -Text "彩色"
$radioColor.Checked = $true
$radioBW = Add-UIControl -Type RadioButton -X 385 -Y 82 -W 55 -H 25 -Text "黑白"

$listView = Add-UIControl -Type ListView -X 15 -Y 110 -W 560 -H 280
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.CheckBoxes = $true
$listView.GridLines = $true
$listView.Columns.Add("文件名", 300)
$listView.Columns.Add("大小", 100)
$listView.Columns.Add("状态", 140)

$progressBar = Add-UIControl -Type ProgressBar -X 15 -Y 398 -W 560 -H 23
$progressBar.Style = "Continuous"

$labelStatus = Add-UIControl -Type Label -X 15 -Y 428 -W 560 -H 23 -Text "请选择包含 PDF 文件的文件夹"
$labelStatus.ForeColor = [System.Drawing.Color]::DarkBlue

$btnPrint = Add-UIControl -Type Button -X 370 -Y 455 -W 100 -H 30 -Text "开始打印"
$btnPrint.Enabled = $false
$btnPrint.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnPrint.ForeColor = [System.Drawing.Color]::White
$btnPrint.FlatStyle = "Flat"

$btnCancel = Add-UIControl -Type Button -X 480 -Y 455 -W 95 -H 30 -Text "取消"
$btnCancel.Enabled = $false

# ============ 全局变量 ============
$script:isPrinting = $false

# ============ 函数 ============

function Update-FileCount {
    $checked = 0
    foreach ($item in $listView.Items) {
        if ($item.Checked) { $checked++ }
    }
    $labelCount.Text = "已选 $checked 个文件"
    $btnPrint.Enabled = ($checked -gt 0 -and -not $script:isPrinting)
}

function Load-PdfFiles {
    param([string]$folderPath)

    $listView.Items.Clear()

    $files = @(Get-ChildItem -Path $folderPath -Filter "*.pdf" -File | Sort-Object Name)

    if ($files.Count -eq 0) {
        $labelStatus.Text = "该文件夹中没有找到 PDF 文件"
        $labelStatus.ForeColor = [System.Drawing.Color]::Red
        $btnPrint.Enabled = $false
        return
    }

    foreach ($file in $files) {
        $item = New-Object System.Windows.Forms.ListViewItem($file.Name)
        $item.SubItems.Add((Format-FileSize $file.Length))
        $item.SubItems.Add("等待打印")
        $item.Checked = $true
        $item.Tag = $file.FullName
        $listView.Items.Add($item)
    }

    $labelStatus.Text = "找到 $($files.Count) 个 PDF 文件，已全部勾选"
    $labelStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    $progressBar.Value = 0
    Update-FileCount
}

# ============ 事件处理 ============

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择包含 PDF 文件的文件夹"
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -eq "OK") {
        $textFolder.Text = $dialog.SelectedPath
        Load-PdfFiles -folderPath $dialog.SelectedPath
    }
})

$btnSelectAll.Add_Click({ Set-AllChecked -Checked $true })
$btnDeselectAll.Add_Click({ Set-AllChecked -Checked $false })

$listView.Add_ItemCheck({
    $form.BeginInvoke([Action]{ Update-FileCount }) | Out-Null
})

# 开始打印
$btnPrint.Add_Click({
    $selectedPrinter = $comboPrinter.SelectedItem
    if (-not $selectedPrinter) {
        [System.Windows.Forms.MessageBox]::Show("请选择打印机！", "提示", "OK", "Warning")
        return
    }

    # 预检查 SumatraPDF
    if (-not (Test-Path $script:SumatraPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "未找到 SumatraPDF：`n$($script:SumatraPath)`n`n请确认文件存在。",
            "错误", "OK", "Error"
        )
        return
    }

    # 收集勾选项（提取数据，不引用 UI 控件）
    $printJobs = [System.Collections.ArrayList]::new()
    foreach ($item in $listView.Items) {
        if ($item.Checked) {
            [void]$printJobs.Add(@{
                Index    = $item.Index
                FileName = $item.Text
                FilePath = $item.Tag
            })
        }
    }

    if ($printJobs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请至少勾选一个文件！", "提示", "OK", "Warning")
        return
    }

    # 确认打印
    $colorMode = if ($radioBW.Checked) { "黑白" } else { "彩色" }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "即将发送 $($printJobs.Count) 个文件到打印机：`n$selectedPrinter`n颜色模式：$colorMode`n`n确认开始打印？",
        "确认", "YesNo", "Question"
    )
    if ($confirm -ne "Yes") { return }

    # 禁用 UI
    $script:isPrinting = $true
    Set-UIForPrinting -Printing $true
    $progressBar.Maximum = $printJobs.Count
    $progressBar.Value = 0

    # 共享取消标志（引用类型，跨 runspace 可见）
    $cancelFlag = @{ Cancelled = $false }

    # 构建打印参数
    $useBW = $radioBW.Checked
    $printerName = $selectedPrinter
    $sumatra = $script:SumatraPath
    $delayMs = $script:JobDelayMs

    # 后台线程
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace

    $ps.AddScript({
        param($jobs, $form, $listView, $progressBar, $labelStatus, $cancelFlag, $printer, $bw, $sumatraExe, $delay)

        $total = $jobs.Count
        $completed = 0
        $failed = 0

        foreach ($job in $jobs) {
            # 检查取消
            if ($cancelFlag.Cancelled) {
                $form.Invoke([Action]{
                    $labelStatus.Text = "打印已取消（已完成 $completed / $total）"
                    $labelStatus.ForeColor = [System.Drawing.Color]::Orange
                })
                break
            }

            # 更新状态：打印中
            $idx = $job.Index
            $form.Invoke([Action]{
                $listView.Items[$idx].SubItems[2].Text = "打印中..."
                $listView.Items[$idx].ForeColor = [System.Drawing.Color]::Blue
                $labelStatus.Text = "正在打印: $($job.FileName)"
                $labelStatus.ForeColor = [System.Drawing.Color]::DarkBlue
            })

            # 执行打印
            $printSettings = if ($bw) { "1x,monochrome" } else { "1x,color" }
            $arguments = "-print-to `"$printer`" -print-settings `"$printSettings`" `"$($job.FilePath)`""
            $errorMsg = ""

            try {
                $proc = Start-Process -FilePath $sumatraExe -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
                if ($proc.ExitCode -eq 0) {
                    $success = $true
                } else {
                    $success = $false
                    $errorMsg = "退出码: $($proc.ExitCode)"
                }
            }
            catch {
                $success = $false
                $_.Exception.Message
            }

            $completed++
            if (-not $success) { $failed++ }

            # 更新状态：结果
            $form.Invoke([Action]{
                if ($success) {
                    $listView.Items[$idx].SubItems[2].Text = "已发送"
                    $listView.Items[$idx].ForeColor = [System.Drawing.Color]::DarkGreen
                } else {
                    $listView.Items[$idx].SubItems[2].Text = "失败"
                    $listView.Items[$idx].ForeColor = [System.Drawing.Color]::Red
                }
                $progressBar.Value = $completed
                $labelStatus.Text = "进度: $completed / $total"
            })

            Start-Sleep -Milliseconds $delay
        }

        # 完成
        $form.Invoke([Action]{
            $script:isPrinting = $false
            Set-UIForPrinting -Printing $false

            if (-not $cancelFlag.Cancelled) {
                if ($failed -eq 0) {
                    $labelStatus.Text = "打印完成！共 $completed 个文件已发送到打印机"
                    $labelStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                    [System.Windows.Forms.MessageBox]::Show(
                        "打印完成！`n`n共 $completed 个文件已发送到打印机。",
                        "完成", "OK", "Information"
                    )
                } else {
                    $labelStatus.Text = "打印完成，$failed 个文件失败"
                    $labelStatus.ForeColor = [System.Drawing.Color]::OrangeRed
                    [System.Windows.Forms.MessageBox]::Show(
                        "打印完成，但有 $failed 个文件失败。`n成功: $($completed - $failed) / $total",
                        "完成", "OK", "Warning"
                    )
                }
            }
        })

        # 清理资源
        $psInner = [powershell]::Create()
        $psInner.Runspace = $runspace
        $psInner.Dispose()
    }).AddArgument($printJobs).AddArgument($form).AddArgument($listView).AddArgument($progressBar).AddArgument($labelStatus).AddArgument($cancelFlag).AddArgument($printerName).AddArgument($useBW).AddArgument($sumatra).AddArgument($delayMs)

    # 启动异步执行
    $handle = $ps.BeginInvoke()

    # 完成后清理 (用定时器检查)
    $cleanupTimer = New-Object System.Windows.Forms.Timer
    $cleanupTimer.Interval = 500
    $cleanupTimer.Add_Tick({
        if ($handle.IsCompleted) {
            $cleanupTimer.Stop()
            $cleanupTimer.Dispose()
            try { $ps.EndInvoke($handle) } catch {}
            $ps.Dispose()
            $runspace.Dispose()
        }
    })
    $cleanupTimer.Start()

    # 保存引用以便取消时使用
    $script:_cancelFlag = $cancelFlag
})

# 取消打印
$btnCancel.Add_Click({
    if ($script:_cancelFlag) {
        $script:_cancelFlag.Cancelled = $true
    }
    $labelStatus.Text = "正在取消..."
    $labelStatus.ForeColor = [System.Drawing.Color]::Orange
})

# ============ 启动 ============
[System.Windows.Forms.Application]::Run($form)
