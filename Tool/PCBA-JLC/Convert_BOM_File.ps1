# アセンブリの読み込み
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# グローバル変数
$global:partsListDirPath = $PSScriptRoot + "`\" + "Parts"
$global:partsList = @{}
$global:IsFileSelect = $false
$global:IsFileGenerate = $false
$global:partsNoValue = "#NOVALUE#"
$global:partsListColumnName = "LCSC Part #"
$global:partsListCSV = [PSCustomObject]@{}
$global:OutFileData = ""



# デバッグ用
function WriteOutput($text) {
    Write-Host (Get-Date).ToString("[HH:mm:ss] ") + $text

    # ツールバーに表示
    $statusLabel.text = $text
}



function GenerateCSV {
    $global:OutFileData = ""

    # 表からヘッダーを取得する
    foreach ($column in $txtOutput.Columns) {
        if ($column.name -ne "Output") {
            if ($column.name -ne $global:partsListColumnName) {
                $global:OutFileData += $column.name + ","
            } else {
                $global:OutFileData += $column.name
            }
        }
    }
    $global:OutFileData += [System.Environment]::NewLine


    # 表からデータを取得する
    foreach($row in $txtOutput.Rows){
        $line = ""
        foreach($cell in $row.Cells){
            # チェックが入っている場合に
            if (($cell.ColumnIndex -ne 0) -and $row.Cells["Output"].Value) {
                if ($cell -ne $row.Cells[$global:partsListColumnName]) {
                    # LCSCパーツナンバー以外のセル
                    $line = $line + "`"" + $cell.Value + "`","
                } else {
                    # LCSCパーツナンバーのセルは最初の要素だけを抜き出す。それ以外は空
                    if ($cell.Value) {
                        $line = $line + "`"" + $cell.Value.Split(" ", 2)[0].Trim() + "`""
                    } else {
                        $line = $line + "`"`""
                    }
                }
            }
        }
        # データが存在した場合に追記
        if ($line) {
            $global:OutFileData += [System.Environment]::NewLine + $line
        }
    }

    # Write-Host $global:OutFileData
    
    $global:IsFileGenerate = $true

    # 保存ボタンを有効にする
    $btnSave.Enabled = $true
}



function GetPartsCount {
    # メイン処理
    foreach($row in $txtOutput.Rows){
        # 処理名取得
        $cellPartsName = ($row.cells["LibRef"].value + " @ " + $row.cells["Footprint"].value).Trim()
        $cellPartsValue = $global:partsNoValue
        if ($row.cells["Value"].value.Trim() -ne "") {
            $cellPartsValue = $row.cells["Value"].value.Trim()
        }

        # パーツリストファイルにコンテンツが存在するかどうか？
        if ($global:partsList.ContainsKey($cellPartsName)) {
            # パーツが存在する
            if ($global:partsList.$cellPartsName.ContainsKey($cellPartsValue)) {
                # 対応するValueも存在する
                #Write-Host $cellPartsName " / " $cellPartsValue
                #Write-Host $global:partsList.$cellPartsName.$cellPartsValue.total
                $cellPartsTotal = $global:partsList.$cellPartsName.$cellPartsValue.total
                # ストックを初期化する
                $row.cells[$global:partsListColumnName].value = ""
                [void]$row.cells[$global:partsListColumnName].Items.Clear()
                # Itemごとの処理
                for($i=0; $i -lt $cellPartsTotal; $i++){
                    $statusLabel.text = "Loading JLCPCB Stock..... [" + $cellPartsName + " / " + $cellPartsValue + "] ( " + ($i+1) + " / " + $cellPartsTotal + " )"
                    if ($global:partsList.$cellPartsName.$cellPartsValue.Item[$i].status -lt 0) {
                        # Statusがマイナスなので処理対象
                        $status = GetJLCPartsCount($global:partsList.$cellPartsName.$cellPartsValue.Item[$i].url)
                        if ($status -lt 0) {
                            # 取得結果がマイナス（エラー発生の場合）
                            $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].status = $status
                            $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].stock = 0
                        } else {
                            # 取得結果がプラス（正常：在庫数）
                            $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].status = 1
                            $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].stock = $status
                        }
                        #Write-Host $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].url
                        #Write-Host $status
                    }
                    # リストに追加する
                    $partsNumber = Split-Path $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].url -leaf
                    [void]$row.cells[$global:partsListColumnName].Items.Add(($partsNumber + " [Stock: " + $global:partsList.$cellPartsName.$cellPartsValue.Item[$i].stock + " pcs]"))
                }
                # ストックの1個目を選択状態にする
                $row.cells[$global:partsListColumnName].value = $row.cells[$global:partsListColumnName].Items[0]
                # DNP用に何もない文字列を追加しておく
                [void]$row.cells[$global:partsListColumnName].Items.Add(" ")
            }
        }
        # Outputにチェックを入れる
        $row.cells["Output"].value = $true

    }

    # CSV生成
    WriteOutput("Generate CSV Data.....")
    GenerateCSV

    # 完了
    WriteOutput("Done !!!!!")

}



function ReadBOMFile($filePath) {
    if (!$global:IsFileSelect) {
        WriteOutput("中断：ファイルが選択されていません")
        return
    }
    # ボタン類を無効に
    $txtFileName.Enabled = $false

    $txtOutput.Columns.Clear()

    $statusLabel.text = "Loading BOM File..."

    # そのままファイルを読み込んでヘッダーの処理に備える
    try {
        $lines = Get-Content $filePath -Encoding UTF8 -ErrorAction Stop
    } catch { 
        WriteOutput("中断：ファイルが開けません")
        return
    }

    # 1行目でBOMファイルかを判定する
    if ($lines[0].Contains("Designator") -and $lines[0].Contains("Footprint") -and $lines[0].Contains("LibRef")) {
        WriteOutput("ファイルを開きました")
    } else {
        WriteOutput("中断：ファイルが開けません")
        return
    }
    
    # ファイルにValue項目があるか確認する
    if ($lines[0].Contains("Value")) {
        WriteOutput("ファイルチェック完了")
    } else {
        WriteOutput("中断：パラメータが足りません。Value要素を追加してBOMを生成してください。")
        return
    }

    # ファイルにLCSC項目があるか確認する
    if ($lines[0].Contains("LCSC")) {
        WriteOutput("中断：すでに LCSC パラメータが含まれています。")
        return
    }

    # 1つ目のカラムにチェックボックスを追加する
    $outputColumnCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{
        Name = "Output"
        HeaderText = "出力"
        AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    }
    [void]$txtOutput.Columns.Add($outputColumnCheck)

    try {
        $global:partsListCSV = Import-Csv $filePath -Encoding UTF8 -ErrorAction Stop
    } catch { 
        WriteOutput("中断：ファイルが開けません")
        return
    }

    $dataTable = New-Object System.Data.DataTable
    $global:partsListCSV[0].PSObject.Properties | ForEach-Object {
        [void]$dataTable.Columns.Add($_.Name)
    }
    $global:partsListCSV | ForEach-Object {
        $row = $dataTable.NewRow()
        $_.PSObject.Properties | ForEach-Object {
            $row[$_.Name] = $_.Value
        }
        [void]$dataTable.Rows.Add($row)
    }

    $txtOutput.DataSource = $dataTable

    #$global:partsListCSV | ogv

    # Read-Only化
    foreach ($column in $txtOutput.Columns) {
        if ($column.name -ne "Output") {
            $column.ReadOnly = $true
        }
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }


    # 最後のカラムに $global:partsListColumnName が存在するかチェック
    if ($txtOutput.Columns[[int]$txtOutput.Columns.Count - 1].Name -eq $global:partsListColumnName) {
        # すでに$global:partsListColumnNameがある
    } else {
        # $global:partsListColumnName カラムを追加する
            $outputColumnParts = New-Object System.Windows.Forms.DataGridViewComboBoxColumn -Property @{
            Name = $global:partsListColumnName
            HeaderText = $global:partsListColumnName
            DisplayMember = "Item"
            ValueMember = "Value"
            AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        }
        [void]$txtOutput.Columns.Add($outputColumnParts)
    }

    $statusLabel.text = "Loading BOM File... done!"

}



# パーツリストファイル読み込み関数
function ReadPartsListFile($filePath) {
    try {
        $partsListFile = Get-Content ($filePath) -Encoding UTF8 -ErrorAction Stop
    } catch { 
        WriteOutput("中断：パーツリストファイルが開けません")
        return
    }
    # 読み込みメインループ
    $partsTitle = ""
    $partsValue = $global:partsNoValue
    $partsUrl = ""
    for($i=0; $i -lt $partsListFile.Count; $i++){
        if($partsListFile[$i].Trim() -ne "") {
            # Markdownの行頭が「# 」だった場合に$partsTitleに読み取り、連想配列に要素（空配列）を追加
            # 例：$global:partsList."CAP_1608 @ C1608_N"
            if ($partsListFile[$i].Substring(0, 2) -eq "# ") {
                $partsTitle = [string]$partsListFile[$i].Substring(2, $partsListFile[$i].Length-2).Trim()
                $global:partsList.Add($partsTitle, @{})
                # Valueを初期値に変更する
                $partsValue = $global:partsNoValue
            # Markdownの行頭が「## 」だった場合に$partsValueに読み取り、連想配列に要素（total=0/item空配列）を追加
            # 例：$global:partsList."CAP_1608 @ C1608_N"."0.1uF"
            } elseif ($partsListFile[$i].Substring(0, 3) -eq "## ") {
                # Valueが空白かどうか
                if ([string]$partsListFile[$i].Substring(3, $partsListFile[$i].Length-3).Trim() -ne "") {
                    $partsValue = [string]$partsListFile[$i].Substring(3, $partsListFile[$i].Length-3).Trim()
                } else {
                    $partsValue = $global:partsNoValue
                }
                $global:partsList.$partsTitle.Add($partsValue, @{"total"=[int]"0"
                 "item"=@{}
                 })
            # Markdownの行頭が「 - 」だった場合に$partsUrlに読み取り、連想配列に要素（total<=status,url,stock）を追加
            # totalに1を加算して戻す。totalが5だった場合は、item[0]～item[4]が存在。
            # 例：$global:partsList."CAP_1608 @ C1608_N"."0.1uF".total
            # 例：$global:partsList."CAP_1608 @ C1608_N"."0.1uF".item[0]
            } elseif ($partsListFile[$i].Substring(0, 3) -eq " - ") {
                # Valueキーがあるかどうか確認する、ない場合は作成する
                if (!$global:partsList.$partsTitle.ContainsKey($partsValue)) {
                    $global:partsList.$partsTitle.Add($partsValue, @{"total"=[int]"0"
                    "item"=@{}
                    })
                }
                $partsUrl = [string]$partsListFile[$i].Substring(3, $partsListFile[$i].Length-3).Trim()
                $count = [int]$global:partsList.$partsTitle.$partsValue.total
                $global:partsList.$partsTitle.$partsValue.item.Add($count, @{"status"=[int]-10
                    "url"=[string]$partsUrl
                    "stock"=[int]0
                    })
                $global:partsList.$partsTitle.$partsValue.total = [int]$count + 1
            }
            # Write-Host $partsListFile[$i]
        }
    }
    WriteOutput ("Read :" + $filePath)
}



# JLCPCB Parts Check
function GetJLCPartsCount($url, $timeout=5, $wait=1000) {
    Start-Sleep -m $wait
    if ($url.IndexOf("https://jlcpcb.com/partdetail/") -ne 0) {
        return [int]-1  # URLエラー
    }
    try {
        $response = Invoke-WebRequest -Method GET -Uri $url -TimeoutSec $timeout
    } catch {
        return [int]-2  # 接続エラー
    }
    $countString = $response.ParsedHtml.body.GetElementsByClassName("smt-count-component")[0].innerText -split '\r?\n' | 
                Select-String -Pattern "^In Stock:".Trim() | Out-String
    $count = $countString.Substring($countString.IndexOf(":") + 1, $countString.Length - $countString.IndexOf(":") - 1).Trim()
    if ($count -eq "") {
        return [int]-3  # 文字列無しエラー
    } elseif (![int]::TryParse($count, [ref]$null)) {
        return [int]-4  # 数字以外の文字列エラー
    } else {
        return [int]$count  # 個数を戻す
    }
}



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
    #Write-Host "Load :"$filename
    ReadBOMFile($txtFileName.Text)
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

    # 保存直前にCSVを生成する
    GenerateCSV

    $dialog = New-Object System.Windows.Forms.SaveFileDialog 
    $dialog.Filter = "CSV ファイル(*.CSV)|*.CSV"
    $dialog.InitialDirectory = Split-Path (Split-Path $filename -Parent) -Parent
    $dialog.Title = "名前をつけて保存"
    $dialog.FileName = "BOM for " + (Split-Path $filename -Leaf)
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



# ウィンドウの作成・表示
$screen = New-Object System.Windows.Forms.Form
$screen.Width = 1024
$screen.Height = 768
$screen.MinimumSize = '400,200'
$screen.AutoScaleMode = 2
$screen.Text = "BOM File Converter"
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

# ステータスバー・ラベル
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.ShowItemToolTips = $true
$screen.Controls.Add($statusStrip)

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.text = "Starting..."
[void]$statusStrip.Items.add($statusLabel)

# ボタンの作成・表示
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "ファイル選択"
$btnLoad.Location = New-Object System.Drawing.Point(5, 5)
$btnLoad.Size = New-Object System.Drawing.Point(100, 26)
$btnLoad.Add_Click({  FileOpen })
$screen.Controls.Add($btnLoad)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "処理実行"
$btnRun.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
            -bor [System.Windows.Forms.AnchorStyles]::Right
$btnRun.Location = New-Object System.Drawing.Point(799, 5)
$btnRun.Size = New-Object System.Drawing.Point(100, 26)
$btnRun.Add_Click({ if ($global:IsFileSelect) { GetPartsCount } })
$screen.Controls.Add($btnRun)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "ファイル保存"
$btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
            -bor [System.Windows.Forms.AnchorStyles]::Right
$btnSave.Location = New-Object System.Drawing.Point(904, 5)
$btnSave.Size = New-Object System.Drawing.Point(100, 26)
$btnSave.Add_Click({ if ($global:IsFileGenerate) { FileSave($txtFileName.Text) } })
$btnSave.Enabled = $false
$screen.Controls.Add($btnSave)

# 入力ボックスの作成・表示
$txtFileName = New-Object System.Windows.Forms.TextBox
$txtFileName.Location = New-Object System.Drawing.Point(110, 9)
$txtFileName.Size = New-Object System.Drawing.Point(684, 26)
$txtFileName.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
                 -bor [System.Windows.Forms.AnchorStyles]::Left `
                 -bor [System.Windows.Forms.AnchorStyles]::Right
$txtFileName.Text = "ファイルをドラッグ・左のボタンをクリック"
$screen.Controls.Add($txtFileName)

$txtOutput = New-Object System.Windows.Forms.DataGridView
#$txtOutput.ReadOnly = $true
$txtOutput.AllowUserToAddRows = $false
$txtOutput.AllowUserToDeleteRows = $false
$txtOutput.MultiSelect = $false
$txtOutput.Location = New-Object System.Drawing.Point(5, 35)
$txtOutput.Size = New-Object System.Drawing.Point(998, 666)
$txtOutput.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
               -bor [System.Windows.Forms.AnchorStyles]::Bottom `
               -bor [System.Windows.Forms.AnchorStyles]::Left `
               -bor [System.Windows.Forms.AnchorStyles]::Right
$screen.Controls.Add($txtOutput)


# コマンドライン引数の処理
foreach ($arg in $args) {
    if (IsTargetFile($arg)) {
        SetTargetFile($arg)
    }
}



# パーツリストファイル一覧取得
$statusLabel.text = "Loading Parts List Files..."
WriteOutput "Loading Parts List Files"
if((Test-Path $global:partsListDirPath) -eq "True"){
    Get-ChildItem -File -Path ($global:partsListDirPath + "`\*") -Exclude README* | ForEach-Object { ReadPartsListFile($_.FullName) }
} else {
    WriteOutput "Path Not Found"
}
$statusLabel.text = "Loading Parts List Files... done!"



# イベントループ開始
[System.Windows.Forms.Application]::Run($screen)
#Write-Host "イベントループ脱出"
#exit


