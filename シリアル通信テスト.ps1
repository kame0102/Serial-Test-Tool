##################################################
# シリアル通信テスト
################################################## 初期設定
Set-StrictMode -Version Latest    #コーディング規則を設定
#---+---+---+---+---+---+---+---+---+---+ アセンブリのロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#---+---+---+---+---+---+---+---+---+---+ 変数宣言
$ScriptTitle = "シリアル通信テスト Ver.250513"  #スクリプト名
$IniFilePath = ".\シリアル通信テスト.ini" #設定ファイルパス
$OpenFlag = $false    #シリアルポートのオープンフラグ
[byte[]]$ByteBuf = new-object byte[] 4096  #データバッファ byte
$ByteBuf_Len = 0        #格納データのサイズ
$RecvChkInterval = 100  #受信チェック間隔 ms
$RecvTimeout = 500      #受信タイムアウト時間 ms 0の場合は終端文字まで無限待ち（受信データが途切れる場合は大きめに設定）
$TerminateChr = [byte]0x0d  #受信データの終端文字 ""でタイムアウトするまで受信
$AutoSendId = -1         #自動送信用データNo -1：停止中
$Encoding = "shift_jis"  #エンコーディング "shift_jis", "utf-8", "utf-16", "utf-32"
$IntervalTime = Get-Date #送受信間隔の計測用（送信・受信完了後の時刻をセット）

################################################## 関数定義
#---+---+---+---+---+---+---+---+---+---+ シリアルポート オープン
Function SerialOpen($ComParam) {
    # COMポート生成、パラメータ設定
    $Script:ComPortObj = New-Object System.IO.Ports.SerialPort $ComParam    #COM番号、ボーレート、データ長、パリティ、ストップ
    $ComPortObj.DtrEnable = $CheckBox_Dtr.Checked    #DTR設定
    $ComPortObj.RtsEnable = $CheckBox_Rts.Checked    #RTS設定
    $ComPortObj.Handshake = $listBox_Flow.SelectedItem   # ハンドシェイク設定 (None、XOnXOff、RequestToSend、RequestToSendXOnXOff)
    $ComPortObj.ReadTimeout = 500   #受信タイムアウト ms
    $ComPortObj.WriteTimeout = 500  #送信タイムアウト ms
    $ComPortObj.NewLine = "`r"   #改行文字設定（WriteLineやReadLineメソッドに適用）
    $ComPortObj.Encoding = [System.Text.Encoding]::GetEncoding($Encoding)    # 文字コード設定

    # シリアルポートエラーイベント（エラー発生時に$trueを返す）通信でパリティエラーが発生した場合など
    $Script:ComErrEventObj = Register-ObjectEvent -InputObject $ComPortObj -EventName "ErrorReceived" -Action {$true}

    # COMポートを開く
    try {
        $ComPortObj.Open()
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("シリアルポートのオープンが出来ません。", "エラー", "OK", "Error")
        Return
    }
    # バッファデータクリア
    Start-Sleep -m 500  #ポートオープン時の不要データ削除用に待機
    $ComPortObj.DiscardInBuffer()    #受信バッファーのクリア
    $ComPortObj.DiscardOutBuffer()   #送信バッファーのクリア
    $Script:ByteBuf_Len = 0

    $TextBox_Log.Text = "シリアルポート　オープン（$ComParam）`r`n"  #ログ表示クリア
    $TextBox_LogBin.Text = ""  #ログ表示クリア
    $Script:LogLineNo = 0
    $Script:IntervalTime = Get-Date  #現在日時セット
    $Script:OpenFlag = $true
}

#---+---+---+---+---+---+---+---+---+---+ シリアルポート クローズ
Function SerialClose() {
    If ($OpenFlag) {
        If ($ComPortObj.IsOpen) {  #COMポートがオープンされているか確認
            $ComPortObj.Close()    #COMポートのクローズ
            $TextBox_Log.AppendText("シリアルポート　クローズ`r`n")
            $Script:OpenFlag = $false
        }
    }
}

#---+---+---+---+---+---+---+---+---+---+ データ送信
Function SendSub($Str) {
    If ($Str -eq "") { Return }

    If (!$OpenFlag) {
        [void][System.Windows.Forms.MessageBox]::Show("シリアルポートがオープンされていません。", "警告", "OK", "Warning")
        Return $false
    }

    #機器接続チェック DSR/CTS共にOFF時は警告
    If (-not ($ComPortObj.CtsHolding -or $ComPortObj.DsrHolding)) {
        If($CheckBox_Warning.Checked){
            [void][System.Windows.Forms.MessageBox]::Show("制御線CTS/DSRがOFFしています。機器の電源および接続を確認してください。", "情報", "OK", "Information")
        }
    }

    If ($Str -ceq "Send_BinFile") {
        $TextBox_Log.AppendText("バイナリファイル送信`r`n")
        ComLogBinDisp "送"
    } Else {
        $Script:ByteBuf_Len = ChrCodeConv $Str  #文字列 -> バイト配列
        ComLogDisp "送" $Str  #ログ表示
    }

    $ComPortObj.Write($ByteBuf, 0, $ByteBuf_Len)  #送信（データバッファ, オフセット, 送信バイト数）

    $Script:IntervalTime = Get-Date  #現在日時セット

    #通信エラー確認
    If (Receive-job -job $ComErrEventObj) {
        [void][System.Windows.Forms.MessageBox]::Show("送信エラーが発生しました。機器の電源を入れ直してください。", "エラー", "OK", "Error")
    }

    Return $true
}

#---+---+---+---+---+---+---+---+---+---+ データ受信
#戻り値：ステータスコードまたは、読取データ  終了またはエラー時は""
Function RecvSub {
    $Script:ByteBuf_Len = 0
    $Time = Get-Date  #現在日時セット
    If ($RecvTimeout -eq 0) {  #無限待ちの場合サブフォームを表示
        $p = Start-Process powershell.exe -NoNewWindow -ArgumentList "-command $cmd" -PassThru  #サブフォーム表示プロセスを起動
    }
    Do {
        If ($ComPortObj.BytesToRead -gt 0) {    #受信データ有無確認
            $Script:ByteBuf_Len += $ComPortObj.Read($ByteBuf, $ByteBuf_Len, $ComPortObj.BytesToRead)  #受信（データバッファ, オフセット, 受信バイト数）
            If ($Combo_Term.Text -ne "") {
                If ($ByteBuf[$ByteBuf_Len - 1] -eq $TerminateChr) { Break }    #データ末尾が終端文字なら受信終了
            }
            $Time = Get-Date  #現在日時を再セット
        }
        If ($RecvTimeout -eq 0) {
            If ($p.HasExited) {  #サブフォームが閉じられたか確認
                $Ms = 1  #タイムアウトで終了させる
            } Else {
                $Ms = 0  #無限待ち
                Start-Sleep -m 1    #処理負荷低減のためWait 1ms
            }
        } Else {
            If ($RecvTimeout -gt 1000) {
                Start-Sleep -m 1    #処理負荷低減のためWait 1ms
            }
            $Ms = ((Get-Date) - $Time).TotalMilliseconds  #経過時間 ms
        }
    } While ($Ms -le $RecvTimeout)    #受信タイムアウト確認

    If ($RecvTimeout -eq 0) {
        If (!$p.HasExited) { $p.kill() }  #サブフォームを閉じる
    }

    If ($Encoding -eq "shift_jis") {
        $RecvStr = [System.Text.Encoding]::Default.GetString($ByteBuf, 0, $ByteBuf_Len)  #バイナリ -> SJIS
    } ElseIf ($Encoding -eq "utf-8") {
        $RecvStr = [System.Text.Encoding]::UTF8.GetString($ByteBuf, 0, $ByteBuf_Len)
    } ElseIf ($Encoding -eq "utf-16") {
        $RecvStr = [System.Text.Encoding]::Unicode.GetString($ByteBuf, 0, $ByteBuf_Len)
    } Else {               #"utf-32"
        $RecvStr = [System.Text.Encoding]::UTF32.GetString($ByteBuf, 0, $ByteBuf_Len)
    }
    $Str = CodeChrConv $RecvStr
    ComLogDisp "受" $Str  #ログ表示

    $Script:IntervalTime = Get-Date  #現在日時セット

    #通信エラー確認
    If (Receive-job -job $ComErrEventObj) {
        [void][System.Windows.Forms.MessageBox]::Show("受信エラーが発生しました。機器の電源を入れ直してください。", "エラー", "OK", "Error")
        Return ""
    }

    #終端文字確認
    If ($Combo_Term.Text -ne "") {
        If ($ByteBuf[$ByteBuf_Len - 1] -ne $TerminateChr) {
            [void][System.Windows.Forms.MessageBox]::Show("データの終端文字が受信できませんでした。", "警告", "OK", "Warning")
        #} Else {
        #    $RecvStr = $RecvStr.Substring(0, $RecvStr.Length - 1)    #終端文字を削除
        }
    }

    Return $Str
    #Return $RecvStr
}

#---+---+---+---+---+---+---+---+---+---+ 通信ログ表示
Function ComLogDisp($Str1, $Str2) {
    $Script:LogLineNo++
    $Str = "$LogLineNo"
    $Time = ((Get-Date) - $IntervalTime).Totalseconds
    If ($Time -lt 100) {  #100秒以下の場合のみ時間表示
        $Str += " [" + ($Time).ToString("00.000") + "] " + $Str1 + " "  #行ヘッダ
    } Else {
        $Str += " [--.---] " + $Str1 + " "  #行ヘッダ
    }
    $TextBox_Log.AppendText($Str + $Str2 + "`r`n")  #テキストログ　追加後キャレットを最下行にする
    ComLogBinDisp $Str
}

Function ComLogBinDisp($Str1) {
    $ByteStr = ""
    For ($i = 0; $i -lt $ByteBuf_Len; $i++) {
        $ByteStr += $ByteBuf[$i].ToString("X2") + " "    #16進数変換 大文字
    }
    $TextBox_LogBin.AppendText($Str1 + $ByteStr + "[$ByteBuf_Len]`r`n")  #バイナリログ
}

#---+---+---+---+---+---+---+---+---+---+ キャラクタコード変換
Function ChrCodeConv($Str) {
    $i = 0
    #<NUL>～<SPC>,<DEL>変換
    Foreach ($c in @(`
            "<NUL>", "<SOH>" ,"<STX>" ,"<ETX>" ,"<EOT>" ,"<ENQ>" ,"<ACK>" ,"<BEL>" ,"<BS>" ,"<HT>" ,"<LF>",`
            "<VT>", "<FF>", "<CR>", "<SO>", "<SI>", "<DLE>", "<DC1>" ,"<DC2>", "<DC3>", "<DC4>", "<NAK>",`
            "<SYN>", "<ETB>", "<CAN>", "<EM>", "<SUB>", "<ESC>", "<FS>", "<GS>", "<RS>", "<US>", "<SPC>")) {
        $Str = $Str.Replace($c, $([char]$i))
        $i++
    }
    $Str = $Str.Replace("<DEL>", $([char]0x7f))
    
    #<01>～<FF>変換
    If ($Encoding -eq "shift_jis") {
        $Ary = [system.text.encoding]::Default.GetBytes("$Str")  #SJIS -> バイト配列
    } ElseIf ($Encoding -eq "utf-8") {
        $Ary = [system.text.encoding]::UTF8.GetBytes("$Str")
    } ElseIf ($Encoding -eq "utf-16") {
        $Ary = [system.text.encoding]::Unicode.GetBytes("$Str")
    } Else {               #"utf-32"
        $Ary = [system.text.encoding]::UTF32.GetBytes("$Str")
    }
    $i = 0
    $ByteStr = ""
    Foreach ($B in $Ary) {
        If ([char]$B -eq "<") {
            $ByteStr = "0x"
        } Else {
            If ($ByteStr -ne "") {
                If ([char]$B -match '[0-9]' -or [char]$B -match '[a-f]') {
                    If ($ByteStr.Length -lt 4) {
                        $ByteStr += [char]$B
                    } Else {
                        $ByteStr = ""
                    }
                } ElseIf ([char]$B -eq ">" -and $ByteStr.Length -eq 4) {
                    $i -= 3
                    $B = [byte]$ByteStr
                    $ByteStr = ""
                } Else {
                    $ByteStr = ""
                }
            }
        }
        $Script:ByteBuf[$i] = $B
        $i++
    }
    Return $i
}

Function CodeChrConv($Str) {
    #<NUL>～<SPC>,<DEL>変換
    $ary = `
        "<NUL>", "<SOH>" ,"<STX>" ,"<ETX>" ,"<EOT>" ,"<ENQ>" ,"<ACK>" ,"<BEL>" ,"<BS>" ,"<HT>" ,"<LF>",`
        "<VT>", "<FF>", "<CR>", "<SO>", "<SI>", "<DLE>", "<DC1>" ,"<DC2>", "<DC3>", "<DC4>", "<NAK>",`
        "<SYN>", "<ETB>", "<CAN>", "<EM>", "<SUB>", "<ESC>", "<FS>", "<GS>", "<RS>", "<US>"  #<SPC>は除外
    For ($i = 0; $i -lt $ary.Length; $i++) {
        $Str = $Str.Replace([string][char]$i, $ary[$i])
    }
    $Str = $Str.Replace([string][char]0x7f, "<DEL>")

    #SJISコード範囲外の<fd>～<ff>変換
    If ($Encoding -eq "shift_jis") {
        For ($i = 253; $i -lt 256; $i++) {
            $EncStr = [System.Text.Encoding]::Default.GetString($i)  #バイナリをSJIS文字列に変更
            $Str = $Str.Replace($EncStr, "<" + $i.ToString("x2") + ">")
        }
    }
    Return $Str
}

#---+---+---+---+---+---+---+---+---+---+ 自動送信
Function AutoSend() {
    $Flag = $true
    If ($AutoSendId -ge $Combo_Send_S.Count) { 
        $Script:AutoSendId = 0
    }
    $i = $AutoSendId
    While ($Combo_Send_S[$AutoSendId].Text -eq "") {
        $Script:AutoSendId++
        If ($AutoSendId -ge $Combo_Send_S.Count) {
            $Script:AutoSendId = 0
        }
        if ($AutoSendId -eq $i) {  #送信データがない
            [void][System.Windows.Forms.MessageBox]::Show("送信データを設定してください。", "警告", "OK", "Warning")
            $Flag = $false
            Break
        }
    }
    If ($AutoSendId -gt -1 -and $Flag) {
        Start-Sleep -m $Combo_Auto.Text  #Delay
        $Flag = SendSub $Combo_Send_S[$AutoSendId].Text   #データ送信
    }
    If (-not $Flag) {
        $TextBox_Log.AppendText("自動送信 停止`r`n")
        $Button_Auto.Text = "自動送信"
        $Script:AutoSendId = -1  #自動送信停止
    }
}

#---+---+---+---+---+---+---+---+---+---+ 設定ファイル読み込み
Function IniFileLoad($FilePath) {
    If (Test-Path $FilePath) {  #ファイル存在確認
        #アイテム全て削除
        foreach ($ComboBox in $Combo_Send_L) {
            $ComboBox.Items.Clear()
        }
        foreach ($ComboBox in $Combo_Send_S) {
            $ComboBox.Items.Clear()
        }
        #$TextBox_Stop.Text = ""

        #アイテム追加
        $ItemTextAry = Get-Content $FilePath
        foreach ($ItemText in $ItemTextAry) {
            For ($i = 0; $i -lt 3; $i++) {
                If ($ItemText -match "^Send_Item_L$i=") {
                    $ItemText = $ItemText.Substring(13, $ItemText.Length - 13)
                    [void] $Combo_Send_L[$i].Items.Add($ItemText)  #アイテム追加
                    If ($Combo_Send_L[$i].Items.Count -eq 1) { $Combo_Send_L[$i].SelectedItem = $ItemText }
                }
            }
            For ($i = 0; $i -lt 10; $i++) {
                If ($ItemText -match "^Send_Item_S$i=") {
                    $ItemText = $ItemText.Substring(13, $ItemText.Length - 13)
                    [void] $Combo_Send_S[$i].Items.Add($ItemText)  #アイテム追加
                    If ($Combo_Send_S[$i].Items.Count -eq 1) { $Combo_Send_S[$i].SelectedItem = $ItemText }
                }
            }
            If ($ItemText -match "^Stop_Item=") {
                $ItemText = $ItemText.Substring(10, $ItemText.Length - 10)
                $TextBox_Stop.Text = $ItemText  #テキスト変更
            }
        }
    }
}

################################################## イニシャル処理
Set-Location -Path $PSScriptRoot    #カレントディレクトリをスクリプトのディレクトリへ変更

#---+---+---+---+---+---+---+---+---+---+ メインフォーム
$WindowSize_X = 600; $WindowSize_Y = 450
#フォーム作成
$Form = New-Object System.Windows.Forms.Form
$Form.Text = $ScriptTitle
$Form.Size = "$WindowSize_X, $WindowSize_Y"
#$Form.MaximumSize = "$WindowSize_X, $WindowSize_Y"    #最大サイズ
$Form.MinimumSize = "$WindowSize_X, $WindowSize_Y"    #最小サイズ
$Form.StartPosition = 'Manual'  #フォーム表示位置 中央：CenterScreen
$Form.Top = 0   #各コンテナ基点
$Form.Left = 0  #各コンテナ基点
#$Form.MaximizeBox = $false  #最大化ボタン
#$Form.MinimizeBox = $false  #最小化ボタン
#$Form.TopMost = $true  #常に手前に表示

#---+---+---+---+---+---+---+---+---+---+ COMポート番号設定
$Pos_X = 10; $Pos_Y = 10  #表示位置
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = 'COM番号'
$Form.Controls.Add($Label)

$Pos_X += 60
#コンボボックス
$Combo_ComPort = New-Object System.Windows.Forms.Combobox
$Combo_ComPort.Location = "$Pos_X, $Pos_Y"
$Combo_ComPort.size = "70, 30"
$Combo_ComPort.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
$Combo_ComPort.FlatStyle = "standard"
#Foreach ($i in @(1..255)) { [void] $Combo_ComPort.Items.Add("COM$i") }  #コンボボックスに項目を追加
$Form.Controls.Add($Combo_ComPort) 
#ドロップダウン時イベント
$Combo_ComPort.Add_DropDown({
    $TextBox_Log.Text = "シリアルポート情報取得`r`n"  #ログ表示クリア
    $ComAry = [System.IO.Ports.SerialPort]::GetPortNames()
    #$ComAry = Get-WmiObject -Class Win32_PnPSignedDriver -Filter "FriendlyName LIKE '%COM%'" | Select-Object -Property FriendlyName  #デバイス情報からシリアル情報取得
    If ($ComAry.Count -eq 0) {
        $TextBox_Log.Text += "デバイスが見つかりません。`r`n"
    } Else {
        Foreach ($Com in $ComAry) { $TextBox_Log.Text += $Com + "`r`n" }
        #Foreach ($Com in $ComAry) { $TextBox_Log.Text += $Com.FriendlyName + "`r`n" }
    }
    $Combo_ComPort.Items.Clear()
    $Com = [System.IO.Ports.SerialPort]::GetPortNames()  #COMポート番号一覧取得
    If ($Com.Count -gt 0) { [void] $Combo_ComPort.Items.AddRange($Com) }  #アイテム追加
})

#---+---+---+---+---+---+---+---+---+---+ ボーレート設定
$Pos_X += 80
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = 'ボーレート'
$Form.Controls.Add($Label)

$Pos_X += 60
#コンボボックス
$Combo_Baud = New-Object System.Windows.Forms.Combobox
$Combo_Baud.Location = "$Pos_X, $Pos_Y"
$Combo_Baud.size = "60, 30"
$Combo_Baud.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
$Combo_Baud.FlatStyle = "standard"
[void] $Combo_Baud.Items.AddRange(@("4800", "9600", "19200", "38400", "57600", "115200"))  #コンボボックスに項目を追加
$Combo_Baud.SelectedItem = "38400"
$Form.Controls.Add($Combo_Baud) 

#---+---+---+---+---+---+---+---+---+---+ その他設定
$Pos_X += 70
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = '他'
$Form.Controls.Add($Label)

$Pos_X += 20
#リストボックス
$listBox_Etc = New-Object System.Windows.Forms.ListBox
$listBox_Etc = New-Object System.Windows.Forms.Combobox
$listBox_Etc.Location = "$Pos_X, $Pos_Y"
$listBox_Etc.size = "50, 30"
[void] $listBox_Etc.Items.AddRange(@("8N1", "8E1", "8O1", "7N1", "7E1", "7O1"))  #リストボックスに項目を追加
$listBox_Etc.SelectedItem = "8N1"
$Form.Controls.Add($listBox_Etc) 

#---+---+---+---+---+---+---+---+---+---+ フロー制御設定
$Pos_X += 60
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = 'フロー制御'
$Form.Controls.Add($Label)

$Pos_X += 60
#リストボックス
$listBox_Flow = New-Object System.Windows.Forms.ListBox
$listBox_Flow = New-Object System.Windows.Forms.Combobox
$listBox_Flow.Location = "$Pos_X, $Pos_Y"
$listBox_Flow.size = "150, 30"
[void] $listBox_Flow.Items.AddRange(@("None", "XOnXOff", "RequestToSend", "RequestToSendXOnXOff"))  #リストボックスに項目を追加
$listBox_Flow.SelectedItem = "None"
$Form.Controls.Add($listBox_Flow) 

#---+---+---+---+---+---+---+---+---+---+ 線
$Pos_X = 10; $Pos_Y += 30  #表示位置
#ラベル
$Label_Line = New-Object System.Windows.Forms.Label
$Label_Line.Location = "$Pos_X, $Pos_Y"
$Label_Line.Size = "560, 1"
$Label_Line.AutoSize = $false
$Label_Line.Text = ""
$Label_Line.BorderStyle = "FixedSingle"  #境界線スタイル FixedSingle　Fixed3D　None
$Form.Controls.Add($Label_Line)

#---+---+---+---+---+---+---+---+---+---+ 受信データ
$Pos_X = 10; $Pos_Y += 10  #表示位置
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "受信データ"
$Label.BorderStyle = "Fixed3D"
$Label.BackColor = "GreenYellow"
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ 受信チェック間隔
$Pos_X += 70
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "チェック間隔`n　　　　ms"
$Form.Controls.Add($Label)

$Pos_X += 70
#数値UpDownコントロール
$Numeric_Int = New-Object System.Windows.Forms.NumericUpDown
$Numeric_Int.location = "$Pos_X, $Pos_Y"
$Numeric_Int.Size = "50, 30"
#$Numeric_Int.TextAlign = "Right"
#$Numeric_Int.UpDownAlign = "Right"
$Numeric_Int.Increment = 10
$Numeric_Int.Maximum = "999"
$Numeric_Int.Minimum = "0"
$Numeric_Int.Text = "100"
$Numeric_Int.InterceptArrowKeys = $True
$Form.Controls.Add($Numeric_Int) 
#値の変更時イベント
$Numeric_Int.Add_TextChanged({
    $Script:RecvChkInterval = [int]$Numeric_Int.Text
    If ($RecvChkInterval -lt 1) {
        $Script:RecvChkInterval = 1
    }
    $Timer.Interval = $RecvChkInterval
})

#---+---+---+---+---+---+---+---+---+---+ 受信ターミネータ
$Pos_X += 60
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "終端文字"
$Form.Controls.Add($Label)

$Pos_X += 60
#コンボボックス
$Combo_Term = New-Object System.Windows.Forms.Combobox
$Combo_Term.Location = "$Pos_X, $Pos_Y"
$Combo_Term.size = "60, 30"
#コンボボックスに項目を追加
[void] $Combo_Term.Items.AddRange(@(`
    "", "<NUL>", "<SOH>" ,"<STX>" ,"<ETX>" ,"<EOT>" ,"<ENQ>" ,"<ACK>" ,"<BEL>" ,"<BS>" ,"<HT>" ,"<LF>",`
    "<VT>", "<FF>", "<CR>", "<SO>", "<SI>", "<DLE>", "<DC1>" ,"<DC2>", "<DC3>", "<DC4>", "<NAK>",`
    "<SYN>", "<ETB>", "<CAN>", "<EM>", "<SUB>", "<ESC>", "<FS>", "<GS>", "<RS>", "<US>", "<SPC>", "<DEL>"))
$Combo_Term.SelectedItem = "<CR>"
$Form.Controls.Add($Combo_Term) 
#値の変更時イベント
$Combo_Term.Add_TextChanged({
    If ($Combo_Term.Text -ne "") {
        $Script:ByteBuf_Len = ChrCodeConv $Combo_Term.Text
        If ($ByteBuf_Len -eq 1) {
            $Script:TerminateChr = $ByteBuf[0]
        } Else {
            [void][System.Windows.Forms.MessageBox]::Show("終端文字は1文字までです。", "警告", "OK", "Warning")
        }
    }
})

#---+---+---+---+---+---+---+---+---+---+ 受信タイムアウト
$Pos_X += 70
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "タイムアウト`n　　　　ms"
$Form.Controls.Add($Label)

$Pos_X += 65
#数値UpDownコントロール
$Numeric_Tout = New-Object System.Windows.Forms.NumericUpDown
$Numeric_Tout.location = "$Pos_X, $Pos_Y"
$Numeric_Tout.Size = "60, 30"
#$Numeric_Tout.TextAlign = "Right"
#$Numeric_Tout.UpDownAlign = "Right"
$Numeric_Tout.Increment = 10
$Numeric_Tout.Maximum = "99999"
$Numeric_Tout.Minimum = "0"
$Numeric_Tout.Text = "$RecvTimeout"
$Numeric_Tout.InterceptArrowKeys = $True
$Form.Controls.Add($Numeric_Tout) 
#値の変更時イベント
$Numeric_Tout.Add_TextChanged({
    $Script:RecvTimeout = [int]$Numeric_Tout.Text
})

$Pos_X += 65
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "0 で無限待ち"
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ 線
$Pos_X = 10; $Pos_Y += 30  #表示位置
#ラベル
$Label_Line = New-Object System.Windows.Forms.Label
$Label_Line.Location = "$Pos_X, $Pos_Y"
$Label_Line.Size = "560, 1"
$Label_Line.AutoSize = $false
$Label_Line.Text = ""
$Label_Line.BorderStyle = "FixedSingle"  #境界線スタイル FixedSingle　Fixed3D　None
$Form.Controls.Add($Label_Line)

#---+---+---+---+---+---+---+---+---+---+ 送信データ 短
$Pos_X = 10; $Pos_Y += 10  #表示位置
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "送信データ"
$Label.BorderStyle = "Fixed3D"
$Label.BackColor = "LightPink"
$Form.Controls.Add($Label)
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Pos_X += 80
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = 'バイナリ変換：<NUL>～<SPC>、<DEL>、<00>～<ff>　　自動送信は下段データのみ'
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ 送信データ 長
$Pos_X = 10; $Pos_Y += 20
[object[]]$Combo_Send_L = new-object object[] 3
[object[]]$Button_Send_L = new-object object[] 3
For ($i = 0; $i -lt 3; $i++) {
    #コンボボックス
    $Combo_Send_L[$i] = New-Object System.Windows.Forms.Combobox
    $Combo_Send_L[$i].Location = "$Pos_X, $Pos_Y"
    $Combo_Send_L[$i].size = "420, 30"
    $Combo_Send_L[$i].DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
    $Combo_Send_L[$i].FlatStyle = "standard"
    #[void] $Combo_Send_L[$i].Items.Add("Item")  #コンボボックスに項目を追加
    $Form.Controls.Add($Combo_Send_L[$i]) 
    
    $Pos_X += 425
    #ボタン
    $Button_Send_L[$i] = New-Object System.Windows.Forms.Button
    $Button_Send_L[$i].Location = "$Pos_X, $Pos_Y"
    $Button_Send_L[$i].size = "45, 23"
    $Button_Send_L[$i].text = "送信"
    $Form.Controls.Add($Button_Send_L[$i])
    $Pos_X += 55
}
# ボタンのクリックイベント
$Button_Send_L[0].Add_Click({
    SendSub $Combo_Send_L[0].Text   #データ送信
    If (-not $Combo_Send_L[0].Items.Contains($Combo_Send_L[0].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_L[0].Items.Insert(0, $Combo_Send_L[0].Text)  #アイテム先頭に追加
    }
})
$Button_Send_L[1].Add_Click({
    SendSub $Combo_Send_L[1].Text   #データ送信
    If (-not $Combo_Send_L[1].Items.Contains($Combo_Send_L[1].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_L[1].Items.Insert(0, $Combo_Send_L[1].Text)  #アイテム先頭に追加
    }
})
$Button_Send_L[2].Add_Click({
    SendSub $Combo_Send_L[2].Text   #データ送信
    If (-not $Combo_Send_L[2].Items.Contains($Combo_Send_L[2].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_L[2].Items.Insert(0, $Combo_Send_L[2].Text)  #アイテム先頭に追加
    }
})

#---+---+---+---+---+---+---+---+---+---+ 送信データ 短
$Pos_X = 10; $Pos_Y += 30
[object[]]$Combo_Send_S = new-object object[] 10
[object[]]$Button_Send_S = new-object object[] 10
For ($i = 0; $i -lt 10; $i++) {
    #コンボボックス
    $Combo_Send_S[$i] = New-Object System.Windows.Forms.Combobox
    $Combo_Send_S[$i].Location = "$Pos_X, $Pos_Y"
    $Combo_Send_S[$i].size = "85, 30"
    $Combo_Send_S[$i].DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
    $Combo_Send_S[$i].FlatStyle = "standard"
    #[void] $Combo_Send_S.Items.AddRange(@("Item1", "Item2", "Item3"))  #コンボボックスに項目を追加
    $Form.Controls.Add($Combo_Send_S[$i]) 
    
    $Pos_X += 90
    #ボタン
    $Button_Send_S[$i] = New-Object System.Windows.Forms.Button
    $Button_Send_S[$i].Location = "$Pos_X, $Pos_Y"
    $Button_Send_S[$i].size = "45, 23"
    $Button_Send_S[$i].text = "送信"
    $Form.Controls.Add($Button_Send_S[$i])
    $Pos_X += 55
}
#ボタンのクリックイベント
$Button_Send_S[0].Add_Click({
    SendSub $Combo_Send_S[0].Text   #データ送信
    If (-not $Combo_Send_S[0].Items.Contains($Combo_Send_S[0].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[0].Items.Insert(0, $Combo_Send_S[0].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[1].Add_Click({
    SendSub $Combo_Send_S[1].Text   #データ送信
    If (-not $Combo_Send_S[1].Items.Contains($Combo_Send_S[1].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[1].Items.Insert(0, $Combo_Send_S[1].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[2].Add_Click({
    SendSub $Combo_Send_S[2].Text   #データ送信
    If (-not $Combo_Send_S[2].Items.Contains($Combo_Send_S[2].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[2].Items.Insert(0, $Combo_Send_S[2].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[3].Add_Click({
    SendSub $Combo_Send_S[3].Text   #データ送信
    If (-not $Combo_Send_S[3].Items.Contains($Combo_Send_S[3].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[3].Items.Insert(0, $Combo_Send_S[3].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[4].Add_Click({
    SendSub $Combo_Send_S[4].Text   #データ送信
    If (-not $Combo_Send_S[4].Items.Contains($Combo_Send_S[4].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[4].Items.Insert(0, $Combo_Send_S[4].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[5].Add_Click({
    SendSub $Combo_Send_S[5].Text   #データ送信
    If (-not $Combo_Send_S[5].Items.Contains($Combo_Send_S[5].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[5].Items.Insert(0, $Combo_Send_S[5].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[6].Add_Click({
    SendSub $Combo_Send_S[6].Text   #データ送信
    If (-not $Combo_Send_S[6].Items.Contains($Combo_Send_S[6].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[6].Items.Insert(0, $Combo_Send_S[6].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[7].Add_Click({
    SendSub $Combo_Send_S[7].Text   #データ送信
    If (-not $Combo_Send_S[7].Items.Contains($Combo_Send_S[7].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[7].Items.Insert(0, $Combo_Send_S[7].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[8].Add_Click({
    SendSub $Combo_Send_S[8].Text   #データ送信
    If (-not $Combo_Send_S[8].Items.Contains($Combo_Send_S[8].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[8].Items.Insert(0, $Combo_Send_S[8].Text)  #アイテム先頭に追加
    }
})
$Button_Send_S[9].Add_Click({
    SendSub $Combo_Send_S[9].Text   #データ送信
    If (-not $Combo_Send_S[9].Items.Contains($Combo_Send_S[9].Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_Send_S[9].Items.Insert(0, $Combo_Send_S[9].Text)  #アイテム先頭に追加
    }
})

#---+---+---+---+---+---+---+---+---+---+ 自動送信
$Pos_X = 10; $Pos_Y += 30
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "自動送信`n　設定"
$Label.BorderStyle = "Fixed3D"
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ Delay
$Pos_X += 60
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "Delay`n　ms"
$Form.Controls.Add($Label)

$Pos_X += 40
#コンボボックス
$Combo_Auto = New-Object System.Windows.Forms.Combobox
$Combo_Auto.Location = "$Pos_X, $Pos_Y"
$Combo_Auto.size = "60, 30"
$Combo_Auto.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
$Combo_Auto.FlatStyle = "standard"
[void] $Combo_Auto.Items.AddRange(@(0, 10, 50, 100, 200, 500, 1000, 2000, 5000, 10000))  #コンボボックスに項目を追加
$Combo_Auto.Text = "100"
$Form.Controls.Add($Combo_Auto) 
# 値の変更時イベント
$Combo_Auto.Add_TextChanged({
    If (-not [int]::TryParse($Combo_Auto.Text, [ref]$null)) {  #数値に変換できるか確認
        [void][System.Windows.Forms.MessageBox]::Show("数値を入力してください。", "警告", "OK", "Warning")
        $Combo_Auto.Text = "100"
    }
})

#---+---+---+---+---+---+---+---+---+---+ 停止受信データ
$Pos_X += 70
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "停止受信データ`n　　　　-clike"
$Form.Controls.Add($Label)

$Pos_X += 90
#テキストボックス
$TextBox_Stop = New-Object System.Windows.Forms.TextBox
$TextBox_Stop.Location = "$Pos_X, $Pos_Y"
$TextBox_Stop.size = "300, 30"
#$TextBox_Stop.MaxLength = 128        #最大入力文字数
$TextBox_Stop.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
                 -bor ([System.Windows.Forms.AnchorStyles]::Top)`
                 -bor ([System.Windows.Forms.AnchorStyles]::Right))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($TextBox_Stop) 

#---+---+---+---+---+---+---+---+---+---+ ファイル送信
$Pos_X = 10; $Pos_Y += 35
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "ファイル送信"
$Label.BorderStyle = "Fixed3D"
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ ファイル選択
$Pos_X += 70; $Pos_Y -= 5
#ボタン
$Button_Fselect = New-Object System.Windows.Forms.Button
$Button_Fselect.Location = "$Pos_X, $Pos_Y"
$Button_Fselect.Size = "80, 23"
$Button_Fselect.Text = "ファイル選択"
$Form.Controls.Add($Button_Fselect)
# ボタンのクリックイベント
$Button_Fselect.Add_Click({
    $DialogFile = New-Object System.Windows.Forms.OpenFileDialog    #ファイル選択ダイアログのオブジェクト取得
    $DialogFile.Filter = 'テキスト・バイナリファイル(*.txt;*.bin)|*.txt;*.bin|全てのファイル(*.*)|*.*'    #フィルタ条件の設定
    $DialogFile.Title = "送信データファイルを選択してください"    #タイトルの設定
    $DialogFile.InitialDirectory = $PSScriptRoot    #デフォルト選択ディレクトリの設定
    if ($DialogFile.ShowDialog() -eq "OK") {
        $Combo_SendF.Text = $DialogFile.FileName  #設定ファイルセット
    }
})

$Pos_X += 90
#コンボボックス
$Combo_SendF = New-Object System.Windows.Forms.Combobox
$Combo_SendF.Location = "$Pos_X, $Pos_Y"
$Combo_SendF.size = "350, 30"
$Combo_SendF.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
$Combo_SendF.FlatStyle = "standard"
$Combo_SendF.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
                 -bor ([System.Windows.Forms.AnchorStyles]::Top)`
                 -bor ([System.Windows.Forms.AnchorStyles]::Right))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Combo_SendF) 

$Pos_X += 355
#ボタン
$Button_SendF = New-Object System.Windows.Forms.Button
$Button_SendF.Location = "$Pos_X, $Pos_Y"
$Button_SendF.Size = "45, 23"
$Button_SendF.Text = "送信"
$Button_SendF.Anchor = (([System.Windows.Forms.AnchorStyles]::Top)`
                 -bor ([System.Windows.Forms.AnchorStyles]::Right))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_SendF)
#ボタンのクリックイベント
$Button_SendF.Add_Click({
    $FilePath = $Combo_SendF.Text
    If (-not (Test-Path $FilePath)) {  #ファイル存在確認
        [void][System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません。", "警告", "OK", "Warning")
        Return
    }
    
    If ([System.IO.Path]::GetExtension($FilePath) -eq ".txt") {  #拡張子txt確認
        #テキストファイル送信
        foreach ($Str in Get-Content $FilePath) {  #1行ずつ読み込み
            SendSub $Str  #データ送信
        }
    } Else {
        #テキストファイル以外は、バイナリファイルとして送信
        $Script:ByteBuf = Get-Content $FilePath -Encoding Byte  #読み込み 数MBのファイルはNG
        #$Script:ByteBuf = [System.IO.File]::ReadAllBytes($FilePath)  #.Netクラス使用　読み込み
        $Script:ByteBuf_Len = (Get-Item $FilePath).Length  #ファイルサイズ取得
        SendSub "Send_BinFile"
    }
    #ファイルパスをコンボアイテムに追加
    If (-not $Combo_SendF.Items.Contains($Combo_SendF.Text)) {  #コンテンツ内の同一アイテムを検索
        $Combo_SendF.Items.Insert(0, $Combo_SendF.Text)  #アイテム先頭に追加
    }
})

#---+---+---+---+---+---+---+---+---+---+ 線
$Pos_X = 10; $Pos_Y += 30  #表示位置
#ラベル
$Label_Line = New-Object System.Windows.Forms.Label
$Label_Line.Location = "$Pos_X, $Pos_Y"
$Label_Line.Size = "560, 1"
$Label_Line.AutoSize = $false
$Label_Line.Text = ""
$Label_Line.BorderStyle = "FixedSingle"  #境界線スタイル FixedSingle　Fixed3D　None
$Form.Controls.Add($Label_Line)

#---+---+---+---+---+---+---+---+---+---+ 通信ログ
$Pos_X = 10; $Pos_Y += 10  #表示位置
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = '通信ログ'
$Label.BorderStyle = "Fixed3D"
$Label.BackColor = "Yellow"
$Form.Controls.Add($Label)
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Pos_X += 60
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = '制御線'
$Form.Controls.Add($Label)

#---+---+---+---+---+---+---+---+---+---+ DTR、RTS
$Pos_X += 50; $Pos_Y -= 8
#チェックボックス
$CheckBox_Dtr = New-Object System.Windows.Forms.CheckBox
$CheckBox_Dtr.Location = "$Pos_X, $Pos_Y"
$CheckBox_Dtr.Size = "50, 30"
$CheckBox_Dtr.Text = "DTR"
$CheckBox_Dtr.Checked = $true
$Form.Controls.Add($CheckBox_Dtr)
#クリックイベント
$CheckBox_Dtr.Add_Click({
    If ($OpenFlag) {
        $Script:ComPortObj.DtrEnable = $CheckBox_Dtr.Checked    #DTR設定
    }
})

$Pos_X += 60
# チェックボックス
$CheckBox_Rts = New-Object System.Windows.Forms.CheckBox
$CheckBox_Rts.Location = "$Pos_X, $Pos_Y"
$CheckBox_Rts.Size = "50, 30"
$CheckBox_Rts.Text = "RTS"
$CheckBox_Rts.Checked = $true
$Form.Controls.Add($CheckBox_Rts)
#クリックイベント
$CheckBox_Rts.Add_Click({
    If ($OpenFlag) {
        $Script:ComPortObj.RtsEnable = $CheckBox_Rts.Checked    #RTS設定
    }
})

#---+---+---+---+---+---+---+---+---+---+ DSR、CTS
$Pos_X += 60
#チェックボックス
$CheckBox_Dsr = New-Object System.Windows.Forms.CheckBox
$CheckBox_Dsr.Location = "$Pos_X, $Pos_Y"
$CheckBox_Dsr.Size = "50, 30"
$CheckBox_Dsr.Text = "DSR"
#$CheckBox_Dsr.Checked = $true
$CheckBox_Dsr.Enabled = $false      #編集不可
$Form.Controls.Add($CheckBox_Dsr)

$Pos_X += 60
#チェックボックス
$CheckBox_Cts = New-Object System.Windows.Forms.CheckBox
$CheckBox_Cts.Location = "$Pos_X, $Pos_Y"
$CheckBox_Cts.Size = "50, 30"
$CheckBox_Cts.Text = "CTS"
#$CheckBox_Cts.Checked = $true
$CheckBox_Cts.Enabled = $false      #編集不可
$Form.Controls.Add($CheckBox_Cts)

$Pos_X += 60
#チェックボックス
$CheckBox_Warning = New-Object System.Windows.Forms.CheckBox
$CheckBox_Warning.Location = "$Pos_X, $Pos_Y"
$CheckBox_Warning.Size = "50, 30"
$CheckBox_Warning.Text = "接続警告"
$CheckBox_Warning.Checked = $true
$Form.Controls.Add($CheckBox_Warning)

#---+---+---+---+---+---+---+---+---+---+ エンコード
$Pos_X += 70; $Pos_Y += 8
#ラベル
$Label = New-Object System.Windows.Forms.Label
$Label.Location = "$Pos_X, $Pos_Y"
$Label.AutoSize = $true            #文字の長さに合わせ自動調整
$Label.Text = "エンコード"
$Form.Controls.Add($Label)

$Pos_X += 60
#コンボボックス
$Combo_Enc = New-Object System.Windows.Forms.Combobox
$Combo_Enc.Location = "$Pos_X, $Pos_Y"
$Combo_Enc.size = "80, 30"
$Combo_Enc.DropDownStyle = "DropDown"    #DropDown、DropDownList、Simleから選択し変更可能
$Combo_Enc.FlatStyle = "standard"
[void] $Combo_Enc.Items.AddRange(@("shift_jis", "utf-8", "utf-16", "utf-32"))  #コンボボックスに項目を追加
$Combo_Enc.Text = $Encoding
$Form.Controls.Add($Combo_Enc) 
#値の変更時イベント
$Combo_Enc.Add_TextChanged({
    $Script:Encoding = $Combo_Enc.Text
    If ($OpenFlag) {
        $Script:ComPortObj.Encoding = [System.Text.Encoding]::GetEncoding($Encoding)    # シリアル通信 文字コード設定
    }
})

#---+---+---+---+---+---+---+---+---+---+ 通信ログ（テキスト）
$Pos_X = 10; $Pos_Y += 25  #表示位置
#テキストボックス
$TextBox_Log = New-Object System.Windows.Forms.TextBox
$TextBox_Log.Location = "$Pos_X, $Pos_Y"
$TextBox_Log.Multiline = $true      #複数行
$TextBox_Log.AcceptsReturn = $true  #改行
$TextBox_Log.WordWrap = $true       #折り返し
#$TextBox_Log.MaxLength = 128        #最大入力文字数
$TextBox_Log.ReadOnly = $true       #編集不可
$TextBox_Log.BackColor = "white"    #背景色
$TextBox_Log.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical    #スクロールバー None:なし Horizontal:水平 Vertical:垂直 Both:水平と垂直
$TextBox_Log.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
              -bor ([System.Windows.Forms.AnchorStyles]::Top)`
              -bor ([System.Windows.Forms.AnchorStyles]::Right)`
              -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$TextBox_Log.Size = "$($WindowSize_X - 35), 54"
$Form.Controls.Add($TextBox_Log)

#---+---+---+---+---+---+---+---+---+---+ 通信ログ（バイナリ）
$Pos_Y += 60
#テキストボックス
$TextBox_LogBin = New-Object System.Windows.Forms.TextBox
$TextBox_LogBin.Location = "$Pos_X, $Pos_Y"
$TextBox_LogBin.Multiline = $true      #複数行
$TextBox_LogBin.AcceptsReturn = $true  #改行
$TextBox_LogBin.WordWrap = $true       #折り返し
#$TextBox_LogBin.MaxLength = 128        #最大入力文字数
$TextBox_LogBin.ReadOnly = $true       #編集不可
$TextBox_LogBin.BackColor = "white"    #背景色
$TextBox_LogBin.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical    #スクロールバー None:なし Horizontal:水平 Vertical:垂直 Both:水平と垂直
$TextBox_LogBin.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
              -bor ([System.Windows.Forms.AnchorStyles]::Right)`
              -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$TextBox_LogBin.Size = "$($WindowSize_X - 35), 54"
$Form.Controls.Add($TextBox_LogBin)

#---+---+---+---+---+---+---+---+---+---+ オープン ボタン
$Pos_X = 10; $Pos_Y = $WindowSize_Y - 70  #表示位置
#ボタン
$Button_ComOpen = New-Object System.Windows.Forms.Button
$Button_ComOpen.Location = "$Pos_X, $Pos_Y"
$Button_ComOpen.Size = "70, 23"
$Button_ComOpen.Text = "オープン"
$Button_ComOpen.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_ComOpen)
#ボタンのクリックイベント
$Button_ComOpen.Add_Click({
    If ($Combo_ComPort.Text -ne "" -and $Combo_Baud.Text -ne "") {
        $Parity = "None"
        If ($listBox_Etc.SelectedItem.Substring(1, 1) -eq "E") {
            $Parity = "Even"
        } ElseIf ($listBox_Etc.SelectedItem.Substring(1, 1) -eq "O") {
            $Parity = "Odd"
        }
        $DataLen = $listBox_Etc.SelectedItem.Substring(0, 1)
        $ComParam = $Combo_ComPort.Text, $Combo_Baud.Text, $Parity, $DataLen, "One"    #COM番号、ボーレート、パリティ（None, Odd, Even, Mark, Space)、データ長、ストップ
        SerialOpen $ComParam  #ポートオープン

        $Combo_ComPort.Enabled = $false
        $Combo_Baud.Enabled = $false
        $listBox_Etc.Enabled = $false
        $listBox_Flow.Enabled = $false
        $Button_ComOpen.Enabled = $false
    } Else {
        [void][System.Windows.Forms.MessageBox]::Show("COM番号を入力してください。", "警告", "OK", "Warning")
    }
})

#---+---+---+---+---+---+---+---+---+---+ クローズ ボタン
$Pos_X += 75
#ボタン
$Button_ComClose = New-Object System.Windows.Forms.Button
$Button_ComClose.Location = "$Pos_X, $Pos_Y"
$Button_ComClose.Size = "70, 23"
$Button_ComClose.Text = "クローズ"
$Button_ComClose.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_ComClose)
#ボタンのクリックイベント
$Button_ComClose.Add_Click({
    SerialClose  #ポートクローズ
    $Combo_ComPort.Enabled = $true
    $Combo_Baud.Enabled = $true
    $listBox_Etc.Enabled = $true
    $listBox_Flow.Enabled = $true
    $Button_ComOpen.Enabled = $true
})

#---+---+---+---+---+---+---+---+---+---+ 自動送信 ボタン
$Pos_X += 80
#ボタン
$Button_Auto = New-Object System.Windows.Forms.Button
$Button_Auto.Location = "$Pos_X, $Pos_Y"
$Button_Auto.Size = "70, 23"
$Button_Auto.Text = "自動送信"
$Button_Auto.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_Auto)
#ボタンのクリックイベント
$Button_Auto.Add_Click({
    If ($AutoSendId -eq -1) {
        $TextBox_Log.AppendText("自動送信 開始`r`n")
        $Button_Auto.Text = "停止"
        $Script:AutoSendId = 0
        AutoSend  #自動送信
    } Else {
        $TextBox_Log.AppendText("自動送信 停止`r`n")
        $Button_Auto.Text = "自動送信"
        $Script:AutoSendId = -1
    }
})

#---+---+---+---+---+---+---+---+---+---+ BIN保存 ボタン
$Pos_X += 75
#ボタン
$Button_Save = New-Object System.Windows.Forms.Button
$Button_Save.Location = "$Pos_X, $Pos_Y"
$Button_Save.Size = "70, 23"
$Button_Save.Text = "BIN保存"
$Button_Save.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_Save)
#ボタンのクリックイベント
$Button_Save.Add_Click({
    If ($ByteBuf_Len -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("バッファ内にデータがありません。", "警告", "OK", "Warning")
    } Else {
        $DialogFile = New-Object System.Windows.Forms.SaveFileDialog    #ファイル選択ダイアログのオブジェクト取得
        $DialogFile.Filter = 'バイナリファイル(*.bin)|*.bin|全てのファイル(*.*)|*.*'    #フィルタ条件の設定
        $DialogFile.Title = "ファイルの保存先およびファイル名を指定してください"    #タイトルの設定
        $DialogFile.InitialDirectory = $PSScriptRoot    #デフォルト選択ディレクトリの設定
        if ($DialogFile.ShowDialog() -eq "OK") {
            Try {
                [System.IO.File]::WriteAllBytes($DialogFile.FileName, $ByteBuf[0..($ByteBuf_Len - 1)])  #バイト配列のデータをファイルに保存
            } Catch {
                [void][System.Windows.Forms.MessageBox]::Show("ファイルの保存に失敗しました。", "エラー", "OK", "Error")
            }
        }
    }
})

#---+---+---+---+---+---+---+---+---+---+ 設定読込 ボタン
$Pos_X += 75
#ボタン
$Button_IniLoad = New-Object System.Windows.Forms.Button
$Button_IniLoad.Location = "$Pos_X, $Pos_Y"
$Button_IniLoad.Size = "70, 23"
$Button_IniLoad.Text = "設定読込"
$Button_IniLoad.Anchor = (([System.Windows.Forms.AnchorStyles]::Left)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_IniLoad)
#ボタンのクリックイベント
$Button_IniLoad.Add_Click({
    $DialogFile = New-Object System.Windows.Forms.OpenFileDialog    #ファイル選択ダイアログのオブジェクト取得
    $DialogFile.Filter = '設定ファイル(*.ini)|*.ini|全てのファイル(*.*)|*.*'    #フィルタ条件の設定
    $DialogFile.Title = "設定ファイルを選択してください"    #タイトルの設定
    $DialogFile.InitialDirectory = $PSScriptRoot    #デフォルト選択ディレクトリの設定
    if ($DialogFile.ShowDialog() -eq "OK") {
        IniFileLoad $DialogFile.FileName  #設定ファイル読み込み
    }
})

#---+---+---+---+---+---+---+---+---+---+ 最前面切替 ボタン
$Pos_X = $WindowSize_X - 200
#ボタン
$Button_Top = New-Object System.Windows.Forms.Button
$Button_Top.Location = "$Pos_X, $Pos_Y"
$Button_Top.Size = "75, 23"
$Button_Top.Text = '最前面切替'
$Button_Top.Anchor = (([System.Windows.Forms.AnchorStyles]::Right)`
               -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$Form.Controls.Add($Button_Top)
#ボタンのクリックイベント
$Button_Top.Add_Click({
    $Form.TopMost = -not $Form.TopMost  #常に手前に表示 反転
})

#---+---+---+---+---+---+---+---+---+---+ 終了 ボタン
$Pos_X += 80
#ボタン
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = "$Pos_X, $Pos_Y"
$CancelButton.Size = "75, 23"
$CancelButton.Text = '終了'
$CancelButton.Anchor = (([System.Windows.Forms.AnchorStyles]::Right)`
                   -bor ([System.Windows.Forms.AnchorStyles]::Bottom))    #位置固定(画面サイズ変更時など)
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
#$Form.CancelButton = $CancelButton
$Form.Controls.Add($CancelButton)

#---+---+---+---+---+---+---+---+---+---+ サブフォーム（受信待ち）
$cmd = {
    #アセンブリのロード
    Add-Type -AssemblyName System.Windows.Forms
    #サブフォーム
    $FormSub = New-Object System.Windows.Forms.Form
    $FormSub.Size = '200, 100'
    $FormSub.StartPosition = 'manual'
    $FormSub.Location = '200, 0'
    $FormSub.MaximizeBox = $false  #最大化ボタン 非表示
    $FormSub.MinimizeBox = $false  #最小化ボタン 非表示
    $FormSub.text = '受信待ち'
    #サブフォームラベル
    $Label = New-Object System.Windows.Forms.Label
    $Label.location = '45, 20'
    $Label.AutoSize = $True        #文字サイズに合わせ自動調整
    $Label.text = '閉じると受信を終了'
    $FormSub.Controls.Add($Label)
    #サブフォーム表示
    $FormSub.Topmost = $true       #フォームを常に手前に表示
    [void]$FormSub.ShowDialog()
}  

#---+---+---+---+---+---+---+---+---+---+ タイマー（起動オプション テキスト表示）
$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = $RecvChkInterval    #イベントを発生させる間隔(ms)
$Time = {
    If ($OpenFlag) {
        $CheckBox_Dsr.Checked = $ComPortObj.DsrHolding
        $CheckBox_Cts.Checked = $ComPortObj.CtsHolding
        If ($ComPortObj.BytesToRead -gt 0) {    #受信データ有無確認
            #$Timer.Stop()    #タイマー停止
            $RecvStr = RecvSub    #ステータス受信
            #$Timer.Start()   #タイマー起動

            #自動送信
            If ($AutoSendId -gt -1) {
                #停止データチェック
                If ($TextBox_Stop.Text -ne "") {
                    $StopDataAry = $TextBox_Stop.Text.split(",")  #カンマ区切りで分割
                    Foreach ($StopData in $StopDataAry) {
                        #If ($RecvStr -ceq $StopData) {  #完全一致
                        If ($RecvStr -clike $StopData) {  #-clike
                        #If ($RecvStr -cmatch $StopData) {  #-cmatch
                        $TextBox_Log.AppendText("自動送信 停止`r`n")
                            $Button_Auto.Text = "自動送信"
                            $Script:AutoSendId = -1  #自動送信停止
                            Break
                        }
                    }
                }
                If ($AutoSendId -gt -1) {
                    $Script:AutoSendId++                    
                    AutoSend  #次データ自動送信
                }
            }
        }
    }
}
$Timer.Add_Tick($Time)

################################################## メイン処理
IniFileLoad $IniFilePath  #設定ファイル読み込み

$Timer.Start()   #タイマー起動
[void]$Form.ShowDialog()    #フォーム表示
$Timer.Stop()    #タイマー停止

################################################## 終了処理
SerialClose    #シリアルポートクローズ
Exit    #スクリプト終了

################################################## End Of File