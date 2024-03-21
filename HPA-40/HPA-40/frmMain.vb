' HPA-40
' frmMain.vb
' ユーザーインターフェース
'
'CORYRIGHT(C) 2023 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2023/12/18 チュオンスアンハイ

Imports HPA_40.clsAllVariable
Imports HPA_40.clsIniFunc
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Threading.Tasks
Imports System.Threading

Public Class frmMain
#Region "ENUM"
#Region "ENUM-R/Wsign"
    Private Enum RWsign
        SGN_READ = 0      'Read
        SGN_WRITE         'Write
    End Enum
#End Region
#End Region

#Region "構造体"
#Region "構造体_ソケット通信設定値"
    Private Structure sttSocket
        Dim strSktIpAddr As String      'IPアドレス
        Dim intSktPort As Integer       'ポート番号
        Dim intSktIntval As Integer     '送信間隔
        Dim intSktTimeout As Integer    'タイムアウト
    End Structure
#End Region
#Region "構造体_Modbus設定値"
    Private Structure sttModbus
        Dim intModTransId As Integer     '転送ID
        Dim intModProtId As Integer      'プロトコルID
        Dim intModUnitId As Integer      'ユニットID
        Dim intModFunc As Integer        'ファンクションコード
        Dim intModReadAddr As Integer    'Read開始アドレス
        Dim intModReadRegist As Integer  'Readレジスタ数
        Dim intModWriteAddr As Integer   'Write開始アドレス
        Dim intModWriteRegist As Integer 'Writeレジスタ数
    End Structure
#End Region
#Region "構造体_Modbus書込みデータ"
    Private Structure sttWriteData
        Dim intWDataLoRaAddr As Integer 'LoRaアドレス
        Dim intWDataModAddr As Integer  'Modbusアドレス
        Dim intWDataYear As Integer     '年
        Dim intWDataMonth As Integer    '月
        Dim intWDataDay As Integer      '日
        Dim intWDataHour As Integer     '時
        Dim intWDataMin As Integer      '分
        Dim intWDataSec As Integer      '秒
    End Structure
#End Region
#End Region

#Region "ディレクトリ"
    Dim basePath As String = funcGetAppPath() & "\Hakaru\"
    Dim basePath1 As String = funcGetAppPath() & "\Hakaru\ErrLog\"
    Dim xmlFilePath As String = funcGetAppPath() & "\" & UNITSETTING_XML_NAME & ".xml"
    Dim xmlFilePath1 As String = funcGetAppPath() & "\" & FTPSETTING_XML_NAME & ".xml"
#End Region

#Region "定数"
    Private Const APP_NAME_1 As String = "HPA-40"
    Private Const UNITSETTING_XML_NAME As String = "Unitsetting"
    Private Const FTPSETTING_XML_NAME As String = "FTPsetting"
    Private Const APP_NAME_2 As String = "JFEスチール向けエッジサーバ転送"
    Private Const VER_NUM As String = "Ver.1.00"
    Private Const TITLE As String = APP_NAME_2 & " (" & APP_NAME_1 & ")" & " " & VER_NUM
    Private Const MES_RCV_OK_1 As String = "受信OK"
    Private Const MES_RCV_ERR_1 As String = "受信エラー：受信できません"
    Private MES_RCV_ERR_2 As String = ""
    Private Const MES_SND_ERR_1 As String = "送信エラー：TCP接続ができません"
    Private Const MES_SND_ERR_2 As String = "送信エラー：送信できません"
    Private Const MES_CHK_ERR As String = "レジスタ数エラーチェック中"
    Private Const REGI_ONE_MAX As Integer = 125
    Private Const REGI_ALL_MAX As Integer = REGI_ONE_MAX * 2
    Private Const TIME_WRITE_DATA_NUM As Integer = 8   'Modbus書込みデータのデータ数
#End Region

#Region "変数"
    Private gobjClsTcpClient As New clsTcpClient   'clsTcpClientクラス
    Private gsttSocket As sttSocket
    Private gsttModbus As sttModbus
    Private gsttWriteData As sttWriteData
    Private gintCntTmr1Tick As Integer             'tmrTimer1のTickが起こった回数
    Private gintCntTmr2Tick As Integer             'tmrTimer2のTickが起こった回数
    Private gbytSndData(32768) As Byte             '送信データ
    Private gintComMaxCnt As Integer               '全てのデータを得るのに必要な送信回数[（ユーザーの指定レジスタ数 \ REGI_ONE_MAX）+ 1]
    Private gintComNowCnt As Integer = 0           '全てのデータを得る際の現在の送信回数
    Private gintComNowCnt_w As Integer = 0         '全てのデータを得る際の現在の送信回数(Write用)
    Private ginitLoop As Boolean = False           'Loop許可変数 (True→Loop許可, False→Loop許可なし)
    Private gLongReg As Boolean = False            'レジスタ数を分ける許可変数 (False→レジスタ数が多きときに自動で分割して送信する, True→レジスタ数が多きときに自動で分割して送信する)
    Private gTimeWrite As Boolean = True           'Time追加許可変数
    'コントロール管理
    Private gtxtReadData(REGI_ALL_MAX - 1) As TextBox
    Private gtxtWriteData(REGI_ALL_MAX - 1) As TextBox
    Private cancellationTokenSource As CancellationTokenSource
#End Region

#Region "FTP設定"
    Private serverUri As String = GetServerParameter(xmlFilePath1).IPAddress
    Private userName As String = GetServerParameter(xmlFilePath1).UserName
    Private password As String = GetServerParameter(xmlFilePath1).PassWord
    Private remoteDirectory As String = GetServerParameter(xmlFilePath1).RemoteDirectory
#End Region

#Region "イベント_起動"
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'フォーム名設定
        Me.Text = TITLE
        'フォーム初期化
        subFormInit()
        'コントロール初期化
        subControl(True)
    End Sub
#End Region

#Region "イベント_終了"
    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'サーバ側との接続を切断する
        gobjClsTcpClient.mTcpClose()
    End Sub
#End Region

#Region "イベント_ボタン押下"
#Region "イベント_ボタン押下_「開始」ボタン"
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        'コントロール無効化
        subControl(False)
        StartMultipleThreads()
    End Sub
#End Region
#Region "イベント_ボタン押下_「停止」ボタン"
    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        'サーバ側との接続を切断する
        gobjClsTcpClient.mTcpClose()
        'コントロール有効化
        subControl(True)
        StopAllThreads()
    End Sub
#End Region
#End Region

#Region "チャンネル数をチェックし、データラインを作成"
    Private Function CreateDataLine(ByVal channelNumber As String, ByVal dtmNow As DateTime, ByVal dataFilter As String)
        Dim dataLine As String = ""
        Select Case channelNumber
            Case 1  '1 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1
            Case 2  '2 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2
            Case 3  '3 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3
            Case 4  '4 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                Dim revertData4 As String = dataFilter.Substring(108, 16)
                Dim decChannel4 As Single = Convert.ToInt32(revertData4, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3 & ", " & decChannel4
            Case 5  '5 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                Dim revertData4 As String = dataFilter.Substring(108, 16)
                Dim decChannel4 As Single = Convert.ToInt32(revertData4, 16) * 0.0001
                Dim revertData5 As String = dataFilter.Substring(144, 16)
                Dim decChannel5 As Single = Convert.ToInt32(revertData5, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3 & ", " & decChannel4 & ", " & decChannel5
            Case 6  '6 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                Dim revertData4 As String = dataFilter.Substring(108, 16)
                Dim decChannel4 As Single = Convert.ToInt32(revertData4, 16) * 0.0001
                Dim revertData5 As String = dataFilter.Substring(144, 16)
                Dim decChannel5 As Single = Convert.ToInt32(revertData5, 16) * 0.0001
                Dim revertData6 As String = dataFilter.Substring(180, 16)
                Dim decChannel6 As Single = Convert.ToInt32(revertData6, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3 & ", " & decChannel4 & ", " & decChannel5 & ", " & decChannel6
            Case 7  '7 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                Dim revertData4 As String = dataFilter.Substring(108, 16)
                Dim decChannel4 As Single = Convert.ToInt32(revertData4, 16) * 0.0001
                Dim revertData5 As String = dataFilter.Substring(144, 16)
                Dim decChannel5 As Single = Convert.ToInt32(revertData5, 16) * 0.0001
                Dim revertData6 As String = dataFilter.Substring(180, 16)
                Dim decChannel6 As Single = Convert.ToInt32(revertData6, 16) * 0.0001
                Dim revertData7 As String = dataFilter.Substring(216, 16)
                Dim decChannel7 As Single = Convert.ToInt32(revertData7, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3 & ", " & decChannel4 & ", " & decChannel5 & ", " & decChannel6 & ", " & decChannel7
            Case 8  '8 Channel
                Dim revertData1 As String = dataFilter.Substring(0, 16)
                Dim decChannel1 As Single = Convert.ToInt32(revertData1, 16) * 0.0001
                Dim revertData2 As String = dataFilter.Substring(36, 16)
                Dim decChannel2 As Single = Convert.ToInt32(revertData2, 16) * 0.0001
                Dim revertData3 As String = dataFilter.Substring(72, 16)
                Dim decChannel3 As Single = Convert.ToInt32(revertData3, 16) * 0.0001
                Dim revertData4 As String = dataFilter.Substring(108, 16)
                Dim decChannel4 As Single = Convert.ToInt32(revertData4, 16) * 0.0001
                Dim revertData5 As String = dataFilter.Substring(144, 16)
                Dim decChannel5 As Single = Convert.ToInt32(revertData5, 16) * 0.0001
                Dim revertData6 As String = dataFilter.Substring(180, 16)
                Dim decChannel6 As Single = Convert.ToInt32(revertData6, 16) * 0.0001
                Dim revertData7 As String = dataFilter.Substring(216, 16)
                Dim decChannel7 As Single = Convert.ToInt32(revertData7, 16) * 0.0001
                Dim revertData8 As String = dataFilter.Substring(252, 16)
                Dim decChannel8 As Single = Convert.ToInt32(revertData8, 16) * 0.0001
                dataLine = dtmNow.ToString("yyyy/MM/dd") & ", " & dtmNow.ToString("HH:mm:00") & ", " & decChannel1 & ", " & decChannel2 & ", " & decChannel3 & ", " & decChannel4 & ", " & decChannel5 & ", " & decChannel6 & ", " & decChannel7 & ", " & decChannel8
            Case Else
        End Select
        Return dataLine
    End Function
#End Region

#Region "ラストラインの時刻をチェックし、CSVにデータを書き込む"
    Private Sub CheckAndWriteCSV(ByVal csvFilePathTemp As String, ByVal dataLine As String)
        Dim lastLine As String = Nothing
        Dim isDuplicate As Boolean = False
        If File.Exists(csvFilePathTemp) Then
            'ラストラインデータを取る
            For Each line As String In File.ReadLines(csvFilePathTemp)
                If Not String.IsNullOrWhiteSpace(line) Then
                    lastLine = line
                End If
            Next
            '現在の記録時刻と比較する
            If Not lastLine Is Nothing Then
                Dim lastDateTime As String = lastLine.Split(","c)(0) & ", " & lastLine.Split(","c)(1)
                Dim newDateTime As String = dataLine.Split(","c)(0) & ", " & dataLine.Split(","c)(1)
                If lastDateTime = newDateTime Then
                    isDuplicate = True
                End If
            End If
        End If
        '違っている場合、CSVにデータを書き込む
        If Not isDuplicate Then
            Using writer As StreamWriter = New StreamWriter(csvFilePathTemp, True)
                writer.WriteLine(dataLine)
            End Using
            'UploadFileToFtp(serverUri, userName, password, csvFilePathTemp, remoteDirectory)
        End If
    End Sub
#End Region

#Region "関数"
#Region "関数_本アプリの起動ディレクトリパスを取得する"
    '本アプリの起動ディレクトリパスを取得する
    'Parameters:
    '  None
    'Returns:
    '  String：本アプリの起動ディレクトリパス
    Private Function funcGetAppPath() As String
        Dim objFileInfo As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Return objFileInfo.DirectoryName
    End Function
#End Region

#Region "関数_フォーム初期化"
    'フォーム初期化
    'Parameters:
    '  None
    'Returns:
    '  None
    Public Sub subFormInit()
        '受信確認用タイマインターバル
        tmrTimer1.Interval = 50
        '送信間隔管理用タイマインターバル
        tmrTimer2.Interval = 100
        'Loop用のチェックボックス
        ginitLoop = False
        '126バイト以上のデータ通信許可
        gLongReg = False   '許可しない
        'ファンクションコード固定
        gsttModbus.intModFunc = &H17
        '書込みデータ固定
        gTimeWrite = True
        'メッセージ
        frmDebug.lblMessageCommunication.Text = "未通信"
        'テキストボックス_送信データ
        frmDebug.txtTransmit.Text = ""
        'テキストボックス_受信データ
        frmDebug.txtReceive.Text = ""
    End Sub
#End Region

#Region "関数_コントロール処理"
    'ボタンなどのコントロールの有効/無効を切り替える
    'Parameters:
    '  ByVal blnEnable As Boolean：True→有効、False→無効
    'Returns:
    '  None
    Private Sub subControl(ByVal blnEnable As Boolean)
        'テキストボックス
        'ラジオボタン

        '「開始」ボタン
        btnStart.Enabled = blnEnable

        '「停止」ボタン
        If blnEnable = True Then
            btnStop.Enabled = False
        Else
            btnStop.Enabled = True
        End If
    End Sub
#End Region

#Region "関数_入力値チェック"
#Region "関数_入力値チェック_TOP"
    '入力値チェック
    'Parameters:
    '  None
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkUserData(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer) As Boolean
        Dim blnRslt As Boolean = False
        Dim strBuf As String
        'Dim intCnt As Integer
        'Dim intRegNum As Integer

        Dim gwIndex1 As Integer = gwIndex
        Dim unitIndex1 As Integer = unitIndex
        Dim channelIndex1 As Integer = channelIndex
        Console.WriteLine("gwIndex1 la: " & gwIndex1)
        Console.WriteLine("unitIndex1 la: " & unitIndex1)
        Console.WriteLine("channelIndex1 la: " & channelIndex1)
        'IPアドレス
        strBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).IPAddress
        blnRslt = funcChkIpAddr(strBuf)
        If blnRslt = False Then
            MessageBox.Show("IPアドレスの入力値が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        'ポート番号
        strBuf = 502
        blnRslt = funcChkPort(strBuf)
        If blnRslt = False Then
            MessageBox.Show("ポート番号の入力値が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        '送信間隔

        strBuf = 200
        blnRslt = strBuf
        If blnRslt = False Then
            MessageBox.Show("送信間隔の入力値が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        ElseIf CInt(strBuf) < 200 Then
            MessageBox.Show("送信間隔は200[ms]以上で設定してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        ElseIf (CInt(strBuf) Mod 100) <> 0 Then
            MessageBox.Show("送信間隔は100[ms]ごとに設定してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        'タイムアウト
        strBuf = 100
        blnRslt = strBuf
        If blnRslt = False Then
            MessageBox.Show("タイムアウトの入力値が正しくありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        ElseIf CInt(strBuf) < 100 Then
            MessageBox.Show("タイムアウトは100[ms]以上で設定してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        ElseIf (CInt(strBuf) Mod 50) <> 0 Then
            MessageBox.Show("タイムアウトは50[ms]ごとに設定してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        '転送ID
        strBuf = (gsttModbus.intModTransId).ToString("X4")
        blnRslt = funcChkHexWord(strBuf)
        If blnRslt = False Then
            MessageBox.Show("転送IDは16進数4桁（ゼロ埋め）で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        'プロトコルID
        strBuf = (gsttModbus.intModProtId).ToString("X4")
        blnRslt = funcChkHexWord(strBuf)
        If blnRslt = False Then
            MessageBox.Show("プロトコルIDは16進数4桁（ゼロ埋め）で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If
        'ユニットID
        strBuf = (gsttModbus.intModUnitId).ToString("X2")
        blnRslt = funcChkHexByte(strBuf)
        If blnRslt = False Then
            MessageBox.Show("ユニットIDは16進数2桁（ゼロ埋め）で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return blnRslt
            Exit Function
        End If

        'ファンクションコードを仮バッファに入れる
        Dim intFuncBuf As Integer = &H17
        '以下はファンクションコードが03Hか04Hか17Hを選択しているときにチェックする
        If (intFuncBuf = &H17) Then
            'Read開始アドレス
            strBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).StartRegister
            blnRslt = strBuf
            If blnRslt = False Then
                MessageBox.Show("開始アドレスは10進数で0～10000の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            ElseIf CInt(strBuf) >= 10000 Then
                MessageBox.Show("開始アドレスは10進数で0～10000の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            End If
            'Readレジスタ数
            strBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).RegisterNumber
            blnRslt = strBuf
            If blnRslt = False Then
                MessageBox.Show("レジスタ数は10進数で0～" & (REGI_ALL_MAX) & "の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            ElseIf CInt(strBuf) >= (REGI_ALL_MAX + 1) Then
                MessageBox.Show("レジスタ数は10進数で0～" & (REGI_ALL_MAX) & "の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            End If
        End If

        '以下はファンクションコードが06Hか10Hか17Hを選択しているときにチェックする
        If (intFuncBuf = &H17) Then
            'Write開始アドレス
            strBuf = 1000
            blnRslt = strBuf
            If blnRslt = False Then
                MessageBox.Show("開始アドレスは10進数で0～10000の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            ElseIf CInt(strBuf) >= 10000 Then
                MessageBox.Show("開始アドレスは10進数で0～10000の整数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            End If
        End If

        '以下はchkTimeWriteにチェックが入っている時のみ確認する。
        If gTimeWrite = True Then
            'LoRaアドレス
            strBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).LoRaAddress
            blnRslt = strBuf
            If blnRslt = False Then
                MessageBox.Show("LoRaアドレスは10進数の自然数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            End If
            'Modbusアドレス
            strBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).ModbusAddress
            blnRslt = strBuf
            If blnRslt = False Then
                MessageBox.Show("Modbusアドレスは10進数の自然数で入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return blnRslt
                Exit Function
            End If
        End If

        blnRslt = True
        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_「IPアドレス」入力値チェック"
    '「IPアドレス」入力値チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkIpAddr(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False

        If System.Text.RegularExpressions.Regex.IsMatch(strData, "^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])(/\d{1,2})?$") Then
            blnRslt = True
        End If

        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_「ポート番号」入力値チェック"
    '「ポート番号」入力値チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkPort(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False

        If System.Text.RegularExpressions.Regex.IsMatch(strData, "^[0-9]{1,5}$") And (CInt(strData) <= &HFFFF) Then
            blnRslt = True
        End If

        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_符号なし整数チェック"
    '符号なし整数チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkUnsignedIntegers(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False

        blnRslt = System.Text.RegularExpressions.Regex.IsMatch(strData, "^[0-9]+$")

        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_符号あり整数チェック"
    '符号あり整数チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkSignedIntegers(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False

        blnRslt = System.Text.RegularExpressions.Regex.IsMatch(strData, "^[+\-]?[0-9]+$")

        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_数値（1バイト16進数）チェック"
    '数値（1バイト16進数）チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkHexByte(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False
        Dim intCnt As Integer
        Dim chrBuf As Char

        If (strData.Length <> 2) Then
            Return blnRslt
            Exit Function
        End If

        For intCnt = 0 To 1
            chrBuf = strData.Substring(intCnt, 1)
            If ((chrBuf >= "0") And (chrBuf <= "9")) Or ((chrBuf >= "a") And (chrBuf <= "f")) Or ((chrBuf >= "A") And (chrBuf <= "F")) Then

            Else
                Return blnRslt
                Exit Function
            End If
        Next

        blnRslt = True
        Return blnRslt
    End Function
#End Region

#Region "関数_入力値チェック_数値（2バイト16進数）チェック"
    '数値（2バイト16進数）チェック
    'Parameters:
    '  ByVal strData As String：入力値
    'Returns:
    '  Boolean：True→OK、False→NG
    Private Function funcChkHexWord(ByVal strData As String) As Boolean
        Dim blnRslt As Boolean = False
        Dim intCnt As Integer
        Dim chrBuf As Char

        If (strData.Length <> 4) Then
            Return blnRslt
            Exit Function
        End If

        For intCnt = 0 To 3
            chrBuf = strData.Substring(intCnt, 1)
            If ((chrBuf >= "0") And (chrBuf <= "9")) Or ((chrBuf >= "a") And (chrBuf <= "f")) Or ((chrBuf >= "A") And (chrBuf <= "F")) Then

            Else
                Return blnRslt
                Exit Function
            End If
        Next

        blnRslt = True
        Return blnRslt
    End Function
#End Region
#End Region

#Region "関数_入力値をグローバル変数に入れる"
    '「IPアドレス」入力値チェック
    'Parameters:
    '  None
    'Returns:
    '  None
    Private Sub subPutSettingData(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer)
        'テキストボックス_IPアドレス
        Dim gwIndex1 As String = gwIndex
        Dim unitIndex1 As String = unitIndex
        Dim channelIndex1 As String = channelIndex

        gsttSocket.strSktIpAddr = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).IPAddress
        'テキストボックス_ポート番号
        gsttSocket.intSktPort = 502
        'テキストボックス_送信間隔
        gsttSocket.intSktIntval = 200
        'テキストボックス_タイムアウト
        gsttSocket.intSktTimeout = 100
        'テキストボックス_転送ID
        gsttModbus.intModTransId = CInt("&H" & "0000")
        'テキストボックス_プロトコルID
        gsttModbus.intModProtId = CInt("&H" & "0000")
        'テキストボックス_ユニットID
        gsttModbus.intModUnitId = CInt("&H" & "00")
        'テキストボックス_Read開始アドレス
        gsttModbus.intModReadAddr = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).StartRegister
        'テキストボックス_Readレジスタ数
        gsttModbus.intModReadRegist = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).RegisterNumber
        'テキストボックス_Write開始アドレス
        gsttModbus.intModWriteAddr = 1000
        'テキストボックス_Writeレジスタ数
        gsttModbus.intModWriteRegist = 8
        'テキストボックス_LoRaアドレス
        gsttWriteData.intWDataLoRaAddr = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).LoRaAddress
        'テキストボックス_Modbusアドレス
        gsttWriteData.intWDataModAddr = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).ModbusAddress
        'ラジオボタン
        Dim intBuf As Integer = &H17
        gsttModbus.intModFunc = &H17

    End Sub
#End Region

#Region "関数_送信処理"
    Private Function funcSndData1(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer) As Boolean
        Dim blnRslt As Boolean = False
        Dim gwIndex1 As Integer = gwIndex
        Dim unitIndex1 As Integer = unitIndex
        Dim channelIndex1 As Integer = channelIndex
        Console.WriteLine("gwIndex la " & gwIndex1)
        Console.WriteLine("unitIndex la " & unitIndex1)
        Console.WriteLine("channelIndex la " & channelIndex1)
        '全てのデータを得る際の現在の送信回数がmaxを超えていれば初期化する
        If (gintComNowCnt >= gintComMaxCnt) Then
            gintComNowCnt = 0
        End If
        '全てのデータを得る際の現在の送信回数がmaxを超えていれば初期化する
        If (gintComNowCnt_w >= gintComMaxCnt) Then
            gintComNowCnt_w = 0
        End If
        'データ作成
        Dim intSndDataLen As Integer
        If gLongReg = False Then
            'chkLongRegにチェックが入っていなければ、レジスタ数が多きときに自動で分割する
            intSndDataLen = funcGetSndDatDev(gwIndex1, unitIndex1, channelIndex1)
        Else
            'chkLongRegにチェックが入っていれば、レジスタ数が多くてもそのまま1回で送信する
            intSndDataLen = funcGetSndDataKeep(gwIndex1, unitIndex1, channelIndex1)
        End If
        If intSndDataLen = 0 Then
            Return blnRslt
            Exit Function
        End If
        '送信データをテキストボックスに書込む
        Dim strSend As String = ""
        For intCnt = 0 To (intSndDataLen - 1)
            strSend = strSend & "[" & gbytSndData(intCnt).ToString("X2") & "h] "
        Next
        '最後に改行を追加する
        'frmDebug.txtTransmit.AppendText(strSend & vbCrLf)

        'UpdateLabel(strSend)
        'データを送信する
        blnRslt = gobjClsTcpClient.mTcpSend(gbytSndData, intSndDataLen)
        '送信処理が正常に行えなかったときは、終了処理を行う
        If blnRslt = False Then
            Return blnRslt
            Exit Function
        End If
        '全てのデータを得る際の現在の送信回数をカウントアップ
        gintComNowCnt += 1
        '受信確認用タイマスタート
        gintCntTmr1Tick = 0
        'tmrTimer1.Start()
        '受信処理
        blnRslt = gobjClsTcpClient.mTcpRcv
        'Dim blnRslt As Boolean
        Dim intRcvStat As Integer
        'タイマストップ
        'tmrTimer1.Stop()
        'Tickが起こった回数をカウントアップ
        gintCntTmr1Tick += 1

        '受信処理
        blnRslt = gobjClsTcpClient.mTcpRcv
        '受信状態フラグ確認
        intRcvStat = gobjClsTcpClient.pRcvStat
        If intRcvStat = RcvStat.RCV_END Then

            Dim dtmNow As DateTime
            dtmNow = DateTime.Now

            '受信完了
            'Tickが起こった回数を初期化する
            gintCntTmr1Tick = 0

            '受信データを取得する
            Dim bytRcvData() As Byte
            bytRcvData = gobjClsTcpClient.pRcvData
            '---受信データを[RawData]タブのテキストボックスに書込む---
            Dim strRcv As String = ""
            Dim strRcv1 As String = ""
            For intCnt = 0 To (bytRcvData.Length - 1)
                strRcv = strRcv & bytRcvData(intCnt).ToString("X2")
            Next

            For intCnt = 0 To (bytRcvData.Length - 1)
                strRcv1 = strRcv1 & bytRcvData(intCnt).ToString("X2")
            Next

            Dim data As String = strRcv1
            '最初にシステム時間を追加する
            frmDebug.txtReceive.AppendText((dtmNow.ToString("yyyy/MM/dd HH:mm:ss")))
            '最後に改行を追加する
            Dim dataFilter As String = (strRcv).Substring(18)
            'ChannelNumberを取る
            Dim channelNumber As Integer = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).ChannelNumber
            'データラインを作成
            Dim dataLine As String = CreateDataLine(channelNumber, dtmNow, dataFilter)

            '処理されているユニットの情報を取得
            Dim gatewayNo As String = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).GateWayNo
            Dim unitNo As String = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).UnitNo
            Dim unitName As String = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).UnitName

            '現在の時刻情報を取得
            Dim currentDate As DateTime = DateTime.Now
            Dim currentYear As String = currentDate.Year.ToString("0000")
            Dim currentMonth As String = currentDate.Month.ToString("00")
            Dim currentDay As String = currentDate.Day.ToString("00")

            'CSVファイルのディレクトリ構造を作成
            Dim csvfilePath As String = basePath.ToString() & gatewayNo & "\" & unitNo & "\" & currentYear & "\" & currentMonth & "\" & unitName & "_" & currentYear & currentMonth & currentDay & ".csv"
            'CSVファイルのヘッダーを初期
            Dim header As String = "日付,時刻,"
            If Not File.Exists(csvfilePath) Then
                Dim directoryPath As String = Path.GetDirectoryName(csvfilePath)
                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If
                'ヘッダーにチャンネルネームを追加
                While (channelIndex1 <= channelNumber)
                    If (channelIndex1 < channelNumber) Then
                        header = header + GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).ChannelName + ","
                    Else
                        header = header + GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).ChannelName
                    End If
                    channelIndex1 = channelIndex1 + 1
                End While
                If (channelIndex1 = channelNumber + 1) Then
                    channelIndex1 = 1
                End If
                Using sw As New StreamWriter(csvfilePath, False, New UTF8Encoding(True))
                    'CSVにヘッダーを書く
                    sw.WriteLine(header)
                    'UploadFileToFtp(serverUri, userName, password, csvfilePath, remoteDirectory)
                End Using
            Else
                'CSVファイルが存在している場合、そのCSVファイルを引き続き書き込む
                CheckAndWriteCSV(csvfilePath, dataLine)

            End If

            'If Not File.Exists(csvfilePath1) Then
            '    Dim directoryPath As String = Path.GetDirectoryName(csvfilePath1)
            '    If Not Directory.Exists(directoryPath) Then
            '        Directory.CreateDirectory(directoryPath)
            '    End If

            'Else
            '    'CSVファイルが存在している場合、そのCSVファイルを引き続き書き込む
            '    CheckAndWriteCSV(csvfilePath1, dataLine)
            'End If


            'chkLongRegにチェックが入っていない時は受信データの正誤チェックを行う

            If gLongReg = False Then
                '受信データチェック
                blnRslt = funcChkRcvData(bytRcvData)
                If blnRslt = False Then
                    '受信エラー処理
                    SubProcRcvErr()

                    Dim intCnt As String = bytRcvData.Length - 1
                    MES_RCV_ERR_2 = "エラーコード" + bytRcvData(intCnt).ToString("X2") & "H" + " :受信できません"

                    '受信エラー→メッセージを表示する
                    frmDebug.lblMessageCommunication.Text = MES_RCV_ERR_2
                    Exit Function
                End If

                '読み出しコマンドの場合のみ、受信データを[Read]タブに表示する
                If gsttModbus.intModFunc = &H17 Then
                    '---受信データを[Read]タブに書込む---
                    Dim intCnt As Integer
                    Dim intCnt2 As Integer = ((gintComNowCnt - 1) * REGI_ONE_MAX)
                    Dim intBuf As Integer
                    For intCnt = 9 To (bytRcvData.Length - 1) Step 2
                        intBuf = (CInt(bytRcvData(intCnt)) << 8)
                        intBuf = intBuf + bytRcvData(intCnt + 1)
                        intCnt2 += 1
                    Next
                    '---受信データを[Read]タブに書込む---
                End If

                '受信OK→メッセージを表示する

                frmDebug.lblMessageCommunication.Text = MES_RCV_OK_1
            Else
                'chkLongRegにチェックが入っているときはメッセージを表示するする
                frmDebug.lblMessageCommunication.Text = MES_CHK_ERR
            End If

            'データを分けて送受信する場合は、送信間隔に関わらず送信する
            If gintComNowCnt < gintComMaxCnt Then
                'データ送信処理
                blnRslt = funcSndData1(gwIndex1, unitIndex1, channelIndex1)
                '送信処理が正常に行えなかったときは、メッセージを表示する
                If blnRslt = False Then
                    'メッセージを表示する
                    frmDebug.lblMessageCommunication.Text = MES_SND_ERR_1
                End If
                'chkLoopにチェックが入っていれば、送信間隔管理用タイマスタートする

                If ginitLoop = True Then
                    '送信間隔管理用タイマスタート
                    gintCntTmr2Tick = 0
                    'tmrTimer2.Start()
                End If
            Else
                'chkLoopにチェックが入っていなければ、停止処理を行う
                If ginitLoop = False Then
                    'サーバ側との接続を切断する
                    gobjClsTcpClient.mTcpClose()
                End If
            End If

        ElseIf intRcvStat = RcvStat.RCV_MID Then
            '受信中
            If gintCntTmr1Tick >= (gsttSocket.intSktTimeout \ tmrTimer1.Interval) Then
                'タイムアウトした
                'Tickが起こった回数を初期化する
                gintCntTmr1Tick = 0

                '受信エラー処理
                SubProcRcvErr()
                '受信エラー→メッセージを表示する
                frmDebug.lblMessageCommunication.Text = MES_RCV_ERR_1
            Else
                'タイムアウトがまだ→タイマスタート
                'tmrTimer1.Start()
                funcSndData1(gwIndex1, unitIndex1, channelIndex1)
            End If

        ElseIf intRcvStat = RcvStat.RCV_ERR Then
            '受信エラー
            'Tickが起こった回数を初期化する
            gintCntTmr1Tick = 0

            '受信エラー処理
            SubProcRcvErr()
            '受信エラー→メッセージを表示する
            frmDebug.lblMessageCommunication.Text = MES_RCV_ERR_1
            Exit Function
        End If
        blnRslt = True
        Return blnRslt
    End Function
#End Region

#Region "関数_送信データ生成"
#Region "関数_送信データ生成_レジスタ数が126以上で指定されていたら通信を2回に分ける"
    '送信データ生成_レジスタ数が126以上で指定されていたら通信を2回に分ける
    'Parameters:
    '  None
    'Returns:
    '  Integer：送信電文長
    Private Function funcGetSndDatDev(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer) As Integer
        Dim intDataLen As Integer
        Dim intBuf As Integer
        Dim intCnt As Integer
        Dim intRegNum As Integer

        Dim gwIndex1 As Integer = gwIndex
        Dim unitIndex1 As Integer = unitIndex
        Dim channelIndex1 As Integer = channelIndex
        '共通処理_転送ID
        gbytSndData(0) = (gsttModbus.intModTransId And &HFF00) >> 8
        gbytSndData(1) = gsttModbus.intModTransId And &HFF
        '共通処理_プロトコルID
        gbytSndData(2) = (gsttModbus.intModProtId And &HFF00) >> 8
        gbytSndData(3) = gsttModbus.intModProtId And &HFF

        Dim intLenBuf As Integer
        intLenBuf = 0

        '転送バイト数を算出するために「ユニットID」から場合分けを行う
        Select Case (gsttModbus.intModFunc)

            Case &H17   'Write/Read Registers
                '共通処理_ユニットID
                gbytSndData(6 + intLenBuf) = gsttModbus.intModUnitId
                intLenBuf += 1
                'ファンクションコード
                gbytSndData(6 + intLenBuf) = gsttModbus.intModFunc
                intLenBuf += 1
                '読み込み開始アドレス
                intBuf = gsttModbus.intModReadAddr + (gintComNowCnt * REGI_ONE_MAX)
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1
                '読み込みレジスタ数
                If gsttModbus.intModReadRegist > ((gintComNowCnt + 1) * REGI_ONE_MAX) Then
                    intBuf = REGI_ONE_MAX
                Else
                    intBuf = gsttModbus.intModReadRegist - (gintComNowCnt * REGI_ONE_MAX)
                End If
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1

                '書き込む開始アドレス
                intBuf = gsttModbus.intModWriteAddr
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1
                '以下はchkTimeWriteにチェックが入っていない時はWriteタブのデータを書き込む、入っているときはStep2で指定した時刻等を書き込む
                If gTimeWrite = False Then
                    'chkTimeWriteにチェックが入っていない時はWriteタブのデータを書き込む

                    If gsttModbus.intModWriteRegist > 125 Then
                        gintComNowCnt_w += 1
                    End If

                    'レジスタ数
                    If gsttModbus.intModWriteRegist > ((gintComNowCnt_w + 1) * REGI_ONE_MAX) Then
                        intBuf = REGI_ONE_MAX
                    Else
                        intBuf = gsttModbus.intModWriteRegist - (gintComNowCnt_w * REGI_ONE_MAX)
                    End If
                    gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                    intLenBuf += 1
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    'データバイト数
                    intBuf = intBuf * 2
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    '書き込むデータ
                    intRegNum = intBuf \ 2
                    For intCnt = (1 + (gintComNowCnt_w * REGI_ONE_MAX)) To (intRegNum + (gintComNowCnt_w * REGI_ONE_MAX))
                        intBuf = CInt(gtxtWriteData(intCnt - 1).Text)
                        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                        intLenBuf += 1
                        gbytSndData(6 + intLenBuf) = intBuf And &HFF
                        intLenBuf += 1
                    Next
                Else
                    'chkTimeWriteにチェックが入っているときはStep2で指定した時刻等を書き込む
                    'レジスタ数
                    intBuf = TIME_WRITE_DATA_NUM
                    gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                    intLenBuf += 1
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    'データバイト数
                    intBuf = intBuf * 2
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    '書き込むデータ
                    intLenBuf = funcGetSndTimeData(intLenBuf, gwIndex1, unitIndex1, channelIndex1)
                End If

                '転送バイト数
                gbytSndData(4) = (intLenBuf And &HFF00) >> 8
                gbytSndData(5) = intLenBuf And &HFF

                intDataLen = 6 + intLenBuf
        End Select
        Return intDataLen
    End Function
#End Region
#Region "関数_送信データ生成_レジスタ数が126以上で指定されていてもそのままの値で1回のみ送信する"
    '送信データ生成_レジスタ数が126以上で指定されていてもそのままの値で1回のみ送信する
    'Parameters:
    '  None
    'Returns:
    '  Integer：送信電文長
    Private Function funcGetSndDataKeep(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer) As Integer
        Dim intDataLen As Integer
        Dim intBuf As Integer
        Dim intCnt As Integer
        Dim intRegNum As Integer

        Dim gwIndex1 As Integer = gwIndex
        Dim unitIndex1 As Integer = unitIndex
        Dim channelIndex1 As Integer = channelIndex

        '共通処理_転送ID
        gbytSndData(0) = (gsttModbus.intModTransId And &HFF00) >> 8
        gbytSndData(1) = gsttModbus.intModTransId And &HFF
        '共通処理_プロトコルID
        gbytSndData(2) = (gsttModbus.intModProtId And &HFF00) >> 8
        gbytSndData(3) = gsttModbus.intModProtId And &HFF

        Dim intLenBuf As Integer
        intLenBuf = 0

        '転送バイト数を算出するために「ユニットID」から場合分けを行う
        Select Case (gsttModbus.intModFunc)
            Case &H17   'Write/Read Registers
                '共通処理_ユニットID
                gbytSndData(6 + intLenBuf) = gsttModbus.intModUnitId
                intLenBuf += 1
                'ファンクションコード
                gbytSndData(6 + intLenBuf) = gsttModbus.intModFunc
                intLenBuf += 1
                '読み込み開始アドレス
                intBuf = gsttModbus.intModReadAddr
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1
                '読み込みレジスタ数
                intBuf = gsttModbus.intModReadRegist
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1

                '書き込む開始アドレス
                intBuf = gsttModbus.intModWriteAddr
                gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                intLenBuf += 1
                gbytSndData(6 + intLenBuf) = intBuf And &HFF
                intLenBuf += 1
                '以下はchkTimeWriteにチェックが入っていない時はWriteタブのデータを書き込む、入っているときはStep2で指定した時刻等を書き込む
                If gTimeWrite = False Then
                    'chkTimeWriteにチェックが入っていない時はWriteタブのデータを書き込む
                    'レジスタ数
                    intBuf = gsttModbus.intModWriteRegist
                    gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                    intLenBuf += 1
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    'データバイト数
                    intBuf = intBuf * 2
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    '書き込むデータ
                    intRegNum = intBuf \ 2
                    For intCnt = 1 To intRegNum
                        intBuf = CInt(gtxtWriteData(intCnt - 1).Text)
                        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                        intLenBuf += 1
                        gbytSndData(6 + intLenBuf) = intBuf And &HFF
                        intLenBuf += 1
                    Next
                Else
                    'chkTimeWriteにチェックが入っているときはStep2で指定した時刻等を書き込む
                    'レジスタ数
                    intBuf = TIME_WRITE_DATA_NUM
                    gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
                    intLenBuf += 1
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    'データバイト数
                    intBuf = intBuf * 2
                    gbytSndData(6 + intLenBuf) = intBuf And &HFF
                    intLenBuf += 1
                    '書き込むデータ
                    intLenBuf = funcGetSndTimeData(intLenBuf, gwIndex1, unitIndex1, channelIndex1)
                End If

                '転送バイト数
                gbytSndData(4) = (intLenBuf And &HFF00) >> 8
                gbytSndData(5) = intLenBuf And &HFF

                intDataLen = 6 + intLenBuf

            Case Else
                intDataLen = 0

        End Select

        Return intDataLen
    End Function
#End Region
#End Region
#Region "関数_送信データ時間データ生成"
    '送信データ時間データ生成
    'Parameters:
    '  ByVal intLenBuf As Integer：オフセット
    'Returns:
    '  Integer：送信電文長
    Private Function funcGetSndTimeData(ByVal intLenBuf As Integer, ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer) As Integer
        Dim intBuf As Integer
        Dim gwIndex1 As Integer = gwIndex
        Dim unitIndex1 As Integer = unitIndex
        Dim channelIndex1 As Integer = channelIndex
        '書き込むデータ_LoRaアドレス
        intBuf = GetSettingParameter(xmlFilePath, gwIndex1, unitIndex1, channelIndex1).LoRaAddress
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_Modbusアドレス
        intBuf = gsttWriteData.intWDataModAddr
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_年
        intBuf = gsttWriteData.intWDataYear
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_月
        intBuf = gsttWriteData.intWDataMonth
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_日
        intBuf = gsttWriteData.intWDataDay
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_時
        intBuf = gsttWriteData.intWDataHour
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_分
        intBuf = gsttWriteData.intWDataMin
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_秒
        intBuf = gsttWriteData.intWDataSec
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1

        Return intLenBuf
    End Function
#End Region
#Region "関数_受信データチェック"
    '受信データチェック
    'Parameters:
    '  ByVal bytRcvData() As Byte：受信データ
    'Returns:
    '  Boolean：True=OK, False=NG
    Private Function funcChkRcvData(ByVal bytRcvData() As Byte) As Boolean
        Dim blnRslt As Boolean = False
        Dim intBuf As Integer
        Dim intRegNum As Integer
        Dim intAddrNum As Integer

        '転送IDチェック
        intBuf = (CInt(bytRcvData(0)) << 8)
        intBuf = intBuf + bytRcvData(1)
        If intBuf <> gsttModbus.intModTransId Then
            Return blnRslt
            Exit Function
        End If

        'プロトコルIDチェック
        intBuf = (CInt(bytRcvData(2)) << 8)
        intBuf = intBuf + bytRcvData(3)
        If intBuf <> gsttModbus.intModProtId Then
            Return blnRslt
            Exit Function
        End If

        '転送バイト数チェック
        intBuf = (CInt(bytRcvData(4)) << 8)
        intBuf = intBuf + bytRcvData(5)
        If intBuf <> (bytRcvData.Length - 6) Then
            Return blnRslt
            Exit Function
        End If

        'ユニットIDチェック
        intBuf = bytRcvData(6)
        If intBuf <> gsttModbus.intModUnitId Then
            Return blnRslt
            Exit Function
        End If

        'ファンクションコードチェック
        intBuf = bytRcvData(7)
        If intBuf <> gsttModbus.intModFunc Then
            Return blnRslt
            Exit Function
        End If

        '以下はファンクションによって異なる
        If (gsttModbus.intModFunc = &H3) OrElse (gsttModbus.intModFunc = &H4) OrElse (gsttModbus.intModFunc = &H17) Then
            '読み出しコマンドの場合は受信データの「バイト数」の項目が、要求したレジスタ数の2倍であることを確認する
            intBuf = bytRcvData(8)
            If gsttModbus.intModReadRegist > (gintComNowCnt * REGI_ONE_MAX) Then
                intRegNum = REGI_ONE_MAX
            Else
                intRegNum = gsttModbus.intModReadRegist - ((gintComNowCnt - 1) * REGI_ONE_MAX)
            End If

            If intBuf <> (intRegNum * 2) Then
                Return blnRslt
                Exit Function
            End If
        ElseIf (gsttModbus.intModFunc = &H6) Then
            '06H書き込みコマンドの場合は、受信データが送信データと同じであることを確認する
            Dim intCnt As Integer
            For intCnt = 8 To (bytRcvData.Length - 1)
                If bytRcvData(intCnt) <> gbytSndData(intCnt) Then
                    Return blnRslt
                    Exit Function
                End If
            Next
        Else
            '10H書き込みコマンドの場合は、受信データの開始アドレスとレジスタ数をチェックする
            '開始アドレス
            intBuf = (CInt(bytRcvData(8)) << 8)
            intBuf = intBuf + bytRcvData(9)
            intAddrNum = gsttModbus.intModWriteAddr + ((gintComNowCnt - 1) * REGI_ONE_MAX)
            If intBuf <> intAddrNum Then
                Return blnRslt
                Exit Function
            End If
            'レジスタ数
            intBuf = (CInt(bytRcvData(10)) << 8)
            intBuf = intBuf + bytRcvData(11)
            If gsttModbus.intModWriteRegist > (gintComNowCnt * REGI_ONE_MAX) Then
                intRegNum = REGI_ONE_MAX
            Else
                intRegNum = gsttModbus.intModWriteRegist - ((gintComNowCnt - 1) * REGI_ONE_MAX)
            End If
            If intBuf <> intRegNum Then
                Return blnRslt
                Exit Function
            End If
        End If

        blnRslt = True
        Return blnRslt
    End Function
#End Region
#Region "関数_受信エラー処理"
    '受信エラー処理
    'Parameters:
    '  None
    'Returns:
    '  None
    Private Sub SubProcRcvErr()
        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt = 0
        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt_w = 0

        'chkLoopにチェックが入っていなければ、停止処理を行う
        If ginitLoop = False Then
            'サーバ側との接続を切断する
            gobjClsTcpClient.mTcpClose()
        End If
    End Sub
#End Region

#End Region


    Private Sub UploadFileToFtp(ByVal serverUriTemp As String, ByVal userNameTemp As String, ByVal passwordTemp As String, ByVal localFilePath As String, ByVal remoteDirectory As String)
        Dim localFilePathTemp As String = localFilePath
        Dim remoteFileName = Path.GetFileName(localFilePathTemp)
        Dim relativePath As String = localFilePathTemp.Substring(localFilePathTemp.IndexOf("Hakaru"))
        Dim partInPath As String() = relativePath.Split("\"c)
        'partinPath(1) →　GatewayNo
        'partinPath(2) →　UnitNo
        'partinPath(3) →　Year
        'partinPath(4) →　Month
        Dim folders As String() = {"datalog", partInPath(1), partInPath(2), partInPath(3), partInPath(4)}
        Dim remotePath As String = serverUriTemp & "datalog/" & partInPath(1) & "/" & partInPath(2) & "/" & partInPath(3) & "/" & partInPath(4) & "/" & remoteFileName
        Dim currentPath As String = ""

        For Each folder In folders
            currentPath &= "/" & folder
            Dim fullUri As String = serverUriTemp.TrimEnd("/"c) & currentPath
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(fullUri), FtpWebRequest)
            request.Credentials = New NetworkCredential(userName, password)
            request.Method = WebRequestMethods.Ftp.MakeDirectory
            Try
                Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
                    'Console.WriteLine("Created folder: " & fullUri)
                End Using
            Catch ex As WebException
                Dim response As FtpWebResponse = DirectCast(ex.Response, FtpWebResponse)
                If response.StatusCode = FtpStatusCode.ActionNotTakenFileUnavailable Then
                    'Console.WriteLine("Folder already exists: " & fullUri)
                Else
                    'Console.WriteLine("Error creating folder: " & fullUri & " - " & ex.Message)
                End If
            End Try
        Next

        Try
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(remotePath), FtpWebRequest)
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = New NetworkCredential(userName, password)
            Dim fileContents As Byte() = File.ReadAllBytes(localFilePathTemp)
            request.ContentLength = fileContents.Length
            Using requestStream As Stream = request.GetRequestStream()
                requestStream.Write(fileContents, 0, fileContents.Length)
            End Using
            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
                Console.WriteLine("Upload successful")
            End Using

        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub btnDebug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        frmDebug.Show()
    End Sub

    Private Sub StartMultipleThreads()
        Dim GWNumber As Integer = GetSettingParameter(xmlFilePath, 0, 0, 0).GateWayNumber
        Dim threadList As New List(Of Thread)()
        Dim paraList As New List(Of clsThreadPara)()

        For i As Integer = 1 To GWNumber
            Dim tempIndex As Integer = i
            Dim paramsThread As New clsThreadPara(tempIndex, 1, 1)
            Dim thread As New Thread(Sub() TaskProcessing(paramsThread))
            thread.Start()
            'Dim paramsThread As New clsThreadPara(tempIndex, 1, 1)
            'paraList.Add(paramsThread)
            'Dim newThread As New Thread(Sub() TaskProcessing(paraList(tempIndex - 1)))
            'threadList.Add(newThread)
        Next

        'For Each t As Thread In threadList
        '    t.Start()
        'Next

        'Dim paramsThread1 As New clsThreadPara(1, 1, 1)
        'Dim thread1 As New Thread(Sub() TaskProcessing(paramsThread1))
        'thread1.Start()

        'Dim paramsThread2 As New clsThreadPara(2, 1, 1)
        'Dim thread2 As New Thread(Sub() TaskProcessing(paramsThread2))
        'thread2.Start()

        'Dim GWNumber As Integer = GetSettingParameter(xmlFilePath, 0, 0, 0).GateWayNumber

        '' Khởi tạo CancellationTokenSource mới
        'cancellationTokenSource = New CancellationTokenSource()

        '' Đảm bảo danh sách thread được xóa trước khi thêm mới
        ''threads.Clear()

        'For i As Integer = 1 To GWNumber
        '    'Dim paramsThread As New clsThreadPara(i, 1, 1, cancellationTokenSource.Token)
        '    Dim paramsThread As New clsThreadPara(i, 1, 1)
        '    Dim thread As New Thread(Sub() TaskProcessing(paramsThread))
        '    thread.Start()
        'Next

    End Sub

    'Private Sub StopAllThreads()
    '    For Each t As Thread In threads
    '        If t.IsAlive Then
    '            t.Interrupt() ' Hoặc sử dụng một phương pháp khác tùy thuộc vào cách bạn xử lý trong TaskProcessing
    '        End If
    '    Next
    'End Sub

    Private Sub StopAllThreads()
        'If cancellationTokenSource IsNot Nothing Then
        '    cancellationTokenSource.Cancel()
        'End If
    End Sub


    Private Sub TaskProcessing(params As clsThreadPara)
        Dim currentValues As clsThreadPara = params.GetValues()
        Dim GWIndex As Integer = currentValues._gwIndex
        Dim UNITIndex As Integer = currentValues._unitIndex
        Dim CHANNELIndex As Integer = currentValues._channelIndex
        'Dim GWNumber As Integer = GetSettingParameter(xmlFilePath, GWIndex, UNITIndex, CHANNELIndex).GateWayNumber
        Dim UNITNumber As Integer = GetSettingParameter(xmlFilePath, GWIndex, UNITIndex, CHANNELIndex).UnitNumber

        While UNITIndex <= UNITNumber

            Dim blnRslt As Boolean
            '入力値のチェック
            blnRslt = funcChkUserData(GWIndex, UNITIndex, CHANNELIndex)
            If blnRslt = False Then
                Exit Sub
            End If
            '入力値をグローバル変数に書込む
            subPutSettingData(GWIndex, UNITIndex, CHANNELIndex)
            'gLongReg:False →　レジスタ数が多きときに自動で分割して送信する
            If gLongReg = False Then
                gintComMaxCnt = (gsttModbus.intModReadRegist \ (REGI_ONE_MAX + 1)) + 1
            Else
                'gLongReg:True →書込まれたレジスタ数のまま1度に送信する
                gintComMaxCnt = 1
            End If
            '全てのデータを得る際の現在の送信回数を初期化する
            gintComNowCnt = 0
            '全てのデータを得る際の現在の送信回数を初期化する
            gintComNowCnt_w = 0
            'TCP接続を行う
            blnRslt = gobjClsTcpClient.mTcpConnect(gsttSocket.strSktIpAddr, gsttSocket.intSktPort, gsttSocket.intSktTimeout)
            'TCP接続ができなかったときは、終了処理を行う
            If blnRslt = False Then
                'メッセージを表示する
                frmDebug.lblMessageCommunication.Text = MES_SND_ERR_1
                Exit Sub
            End If
            'データ送信処理
            blnRslt = funcSndData1(GWIndex, UNITIndex, CHANNELIndex)
            Task.Delay(1000).Wait()
            '送信処理が正常に行えなかったときは、メッセージを表示する
            If blnRslt = False Then
                'メッセージを表示する
                frmDebug.lblMessageCommunication.Text = MES_SND_ERR_1
                'chkLoopにチェックが入っていなければ、停止処理を行う

                If ginitLoop = False Then
                    'チェックが入っていなければ、停止処理を行う
                    'サーバ側との接続を切断する
                    gobjClsTcpClient.mTcpClose()
                End If
            End If
            'chkLoopにチェックが入っていれば、送信間隔管理用タイマスタートする
            If ginitLoop = True Then
                '送信間隔管理用タイマスタート
                gintCntTmr2Tick = 0
                'データ送信処理
                blnRslt = funcSndData1(GWIndex, UNITIndex, CHANNELIndex)
                '送信処理が正常に行えなかったときは、メッセージを表示する
                If blnRslt = False Then
                    'メッセージを表示する
                    frmDebug.lblMessageCommunication.Text = MES_SND_ERR_1
                End If
            End If

            '    If UNITIndex <= UNITNumber Then
            '        Task.Delay(1000).Wait() ' Tạo độ trễ 1 giây
            '        UNITIndex = UNITIndex + 1
            '        params.Update(GWIndex, UNITIndex, CHANNELIndex)

            '    Else
            '        UNITIndex = 1
            '        Task.Delay(1000).Wait()
            '        params.Update(GWIndex, UNITIndex, CHANNELIndex)
            '    End If

            'Else

            '    UNITIndex = 1

            '    Task.Delay(1000).Wait()

            '    params.Update(GWIndex, UNITIndex, CHANNELIndex)
            If UNITIndex = UNITNumber Then
                UNITIndex = 1
            Else
                UNITIndex = UNITIndex + 1
            End If
            params.Update(GWIndex, UNITIndex, CHANNELIndex)
            Task.Delay(2000).Wait()
        End While
    End Sub
End Class

