# アセンブリの読み込み
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")


# グローバル変数
$global:IsFileSelect = $false
$global:IsFileGenerate = $false
$global:rotation_data = ""
$global:csv_data = ""
$global:OutFileData = ""


# CSVファイルのみ処理対象にする
function IsTargetFile($filename) {
    if ([IO.Path]::GetExtension($filename) -eq ".csv") {
        return $true
    } else {
        return $false
    }
}


# 対象のファイルをセット
function SetTargetFile($filename) {
    $global:IsFileSelect = $true
    $txtFileName.Text = $filename
    WriteOutput("ファイル選択 : " + $filename)
    Write-Host "Load :"$filename
}


# テキスト欄に追記
function WriteOutput($text) {
    $txtOutput.AppendText((Get-Date).ToString("[HH:mm:ss] ") + $text + [System.Environment]::NewLine)
    Write-Host (Get-Date).ToString("[HH:mm:ss] ") + $text
}


# ファイルを開くダイアログ
function FileOpen() {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV ファイル(*.CSV)|*.CSV"
    #$dialog.InitialDirectory = "C:\"
    $dialog.Title = "開く"
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        SetTargetFile($dialog.FileName)
    }
}


# ファイルを保存するダイアログ
function FileSave($filename) {
    $dialog = New-Object System.Windows.Forms.SaveFileDialog 
    $dialog.Filter = "CSV ファイル(*.CSV)|*.CSV"
    $dialog.InitialDirectory = Split-Path (Split-Path $filename -Parent) -Parent
    $dialog.Title = "名前をつけて保存"
    $dialog.FileName = Split-Path $filename -Leaf
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        # PowerShell v5 is Not Support -Encoding utf8NoBOM
        try {
            $global:OutFileData | Out-String `
            | % { [Text.Encoding]::UTF8.GetBytes($_) } `
            | Set-Content -Path ($dialog.FileName) -Encoding Byte
        } catch { 
            WriteOutput("中断：ファイルが保存できません")
            return
        }
        WriteOutput("ファイル書き込み完了")
    }
}


# コントロールを有効化する
function EnableControl() {
    $btnLoad.Enabled = $true
    $btnRun.Enabled = $true
    $txtFileName.Enabled = $true
}


# Rotation データ読み込み
function GetRotation() {
    # Rotationファイルを読み込む
    try {
        $rotation = Get-Content ($PSScriptRoot + "\Rotation.md") -Encoding UTF8 -ErrorAction Stop
    } catch { 
        WriteOutput("中断：Rotationファイルが開けません")
        return
    }

    # テーブル部分の位置を探す
    $rotation_table_line_start = 0
    for($i=0; $i -lt $rotation.Count; $i++){
        if($rotation[$i].Contains("|")) {
            $rotation_table_line_start = $i
            break
        }
    }
    $rotation_table_line_end = 0
    for($i=$rotation.Count-1; $i -gt $rotation_table_line_start-1; $i--){
        if($rotation[$i].Contains("|")) {
            $rotation_table_line_end = $i
            break
        }
    }

    # テーブルが無い
    if (($rotation_table_line_end - $rotation_table_line_start) -lt 2) {
        WriteOutput("中断：Rotationデータがありません")
        return
    }

    # 読み込み処理
    $rotation_table_index = $rotation[$rotation_table_line_start].Trim().SubString(1, $rotation[$rotation_table_line_start].Trim().Length-2).Trim().Replace("`t", "")
    $rotation_table=""
    for($i=$rotation_table_line_start+2; $i -lt $rotation_table_line_end; $i++){
        $rotation_table += ($rotation[$i].Trim().SubString(1, $rotation[$i].Length-2).Split("|").Trim() -join ",") + [System.Environment]::NewLine
    }
    $global:rotation_data = ConvertFrom-Csv -InputObject $rotation_table -Header ($rotation_table_index.Split("|").Trim())

    WriteOutput("Rotationデータ読み込み完了")

    #$global:rotation_data | ogv

}


# メイン処理
function ConvertFile() {
    if (!$global:IsFileSelect) {
        WriteOutput("中断：ファイルが選択されていません")
        return
    }
    # ボタン類を無効に
    $btnLoad.Enabled = $false
    $btnRun.Enabled = $false
    $txtFileName.Enabled = $false

    WriteOutput("処理開始")

    # そのままファイルを読み込んでヘッダーの処理に備える
    try {
        $lines = Get-Content $txtFileName.Text -Encoding UTF8 -ErrorAction Stop
    } catch { 
        WriteOutput("中断：ファイルが開けません")
        EnableControl
        return
    }

    # 1行目でPnPファイルかを判定する
    if(!$lines[0].Contains("Altium Designer Pick and Place Locations")){
        WriteOutput("中断：Altium Designer Pick and Place Locations ファイルではありません")
        EnableControl
        return
    }

    # ヘッダー部分を探す
    $header_line = 0
    for($i=0; $i -lt $lines.Count; $i++){
        if($lines[$i].Contains("Designator")) {
            $header_line = $i
            WriteOutput("CSVヘッダ行："+($header_line+1)+"行目")
            WriteOutput("ヘッダ内容："+$lines[$header_line])
            break
        }
    }

    # ヘッダー部分が無い場合
    if ($header_line -eq 0) {
        WriteOutput("中断：CSV ヘッダ行が見つかりません")
        EnableControl
        return
    }

    # ヘッダー2行目のファイル名フルパスを削除する
    $lines[1] = Split-Path $lines[1] -Leaf

    # ヘッダー部分を読み飛ばして再度ファイルを読み込む
    try {
        $csv = Get-Content $txtFileName.Text -Encoding UTF8 -ErrorAction Stop | Select-Object -Skip ($header_line + 1)
    } catch { 
        WriteOutput("中断：ファイルが開けません")
        EnableControl
        return
    }

    # CSV としてデータを読み込む
    $global:csv_data = ConvertFrom-Csv -InputObject $csv -Header ($lines[$header_line].Replace("`"", "").Split(","))

    #$global:csv_data | ogv
    #$global:rotation_data | ogv

    # メインループ（Pick and Place）
    foreach ($item in $global:csv_data) {
        $item_rotation = ($global:rotation_data | Where-Object {($_.Comment -match $item.Comment) -and ($_.Footprint -match $item.Footprint)} | Select-Object -First 1).Rotation
        $output_text = $item.Designator+" : "
        if ($item_rotation) {
            # ローテーションが宣言されている場合
            [Int32]$item.Rotation = [Int32]$item.Rotation + [Int32]$item_rotation
            if ([Int32]$item_rotation -eq 0) {
                $output_text += " [NO CHANGE] "
            } else {
                $output_text += " [EDIT] "
            }
        } else {
            # ローテーションが宣言されていない場合
            $output_text += " [NO ROTATION DATA!!!] "
        }
        # 角度チェック（360度以上になった場合に引く）
        if ([Int32]$item.Rotation -ge 360) {
            [Int32]$item.Rotation = [Int32]$item.Rotation % 360
        }
        $output_text += $item.Comment + " / " + $item.Footprint + " @ " + $item.Rotation + " degrees"
        WriteOutput($output_text)
    }

    # ファイル出力処理
    $OutFileHeader = ""
    for($i=0; $i -lt $header_line; $i++){
        $OutFileHeader += $lines[$i] + [System.Environment]::NewLine
    }
    $global:OutFileData = $OutFileHeader + (($global:csv_data | ConvertTo-Csv -NoTypeInformation) -join [System.Environment]::NewLine)

    # 出力フラグ
    $global:IsFileGenerate = $true

    # ボタン類を有効に戻す
    $btnLoad.Enabled = $true
    $btnLoad.Text = "ファイル保存"
    $btnRun.Text = "処理終了"
    $screen.AllowDrop = $false

    WriteOutput("処理終了")

    # ファイル保存ダイアログ
    FileSave($txtFileName.Text)

}


# ウィンドウの作成・表示
$screen = New-Object System.Windows.Forms.Form
$screen.Width = 640
$screen.Height = 480
$screen.MinimumSize = '400,200'
$screen.AutoScaleMode = 2
$screen.Text = "Centroid File Converter"
$screen.AllowDrop = $true
$screen.Add_DragDrop({
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        if (IsTargetFile($filename)) {
            SetTargetFile($filename)
        }
    }
})
$screen.Add_DragOver({
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        if (IsTargetFile($filename)) {
            $_.Effect = [Windows.Forms.DragDropEffects]::All
        }
    }
})
$screen.Show()
$screen.Activate()

# ボタンの作成・表示
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "ファイル選択"
$btnLoad.Location = New-Object System.Drawing.Point(5, 5)
$btnLoad.Size = New-Object System.Drawing.Point(100, 26)
$btnLoad.Add_Click({ if ($global:IsFileGenerate) { FileSave($txtFileName.Text) } else { FileOpen } })
$screen.Controls.Add($btnLoad)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "処理実行"
$btnRun.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
            -bor [System.Windows.Forms.AnchorStyles]::Right
$btnRun.Location = New-Object System.Drawing.Point(520, 5)
$btnRun.Size = New-Object System.Drawing.Point(100, 26)
$btnRun.Add_Click({ ConvertFile })
$screen.Controls.Add($btnRun)

# 入力ボックスの作成・表示
$txtFileName = New-Object System.Windows.Forms.TextBox
$txtFileName.Location = New-Object System.Drawing.Point(110, 9)
$txtFileName.Size = New-Object System.Drawing.Point(405, 26)
$txtFileName.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
                 -bor [System.Windows.Forms.AnchorStyles]::Left `
                 -bor [System.Windows.Forms.AnchorStyles]::Right
#$txtFileName.Add_TextChanged({param($s,$e) inputhere_textchanged $s $e})
$txtFileName.Text = "ファイルパスを入力・ファイルをドラッグ・左のボタンをクリック"
$screen.Controls.Add($txtFileName)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtOutput.ReadOnly = $true
$txtOutput.Location = New-Object System.Drawing.Point(5, 35)
$txtOutput.Size = New-Object System.Drawing.Point(615, 400)
$txtOutput.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
               -bor [System.Windows.Forms.AnchorStyles]::Bottom `
               -bor [System.Windows.Forms.AnchorStyles]::Left `
               -bor [System.Windows.Forms.AnchorStyles]::Right
$txtOutput.Text = ""
$screen.Controls.Add($txtOutput)

WriteOutput("起動")

# コマンドライン引数の処理
foreach ($arg in $args) {
    if (IsTargetFile($arg)) {
        SetTargetFile($arg)
    }
}

# Rotationデータ読み込み
GetRotation


# イベントループ開始
[System.Windows.Forms.Application]::Run($screen)
#Write-Host "イベントループ脱出"
#exit


