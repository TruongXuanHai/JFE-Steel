' HPA-40
' clsTcpClient.vb
' TCPクライアント処理
'
'CORYRIGHT(C) 2023 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2023/12/18 チュオンスアンハイ

Imports HPA_40.clsAllVariable

Public Class clsTcpClient

#Region "変数"
    Private gobjTcp As System.Net.Sockets.TcpClient           'TcpClient
    Private gobjNtwStrm As System.Net.Sockets.NetworkStream   'NetworkStream
    Private gobjMemStrm As New System.IO.MemoryStream()
    Private gintRcvStat As Integer                            '受信状態フラグ
    Private gbytRcvData As Byte()                             '受信データ
#End Region

#Region "プロパティ"
#Region "プロパティ_受信状態フラグ"
    Public ReadOnly Property pRcvStat() As Integer
        Get
            Return gintRcvStat
        End Get
    End Property
#End Region
#Region "プロパティ_受信データ"
    Public ReadOnly Property pRcvData() As Byte()
        Get
            Return gbytRcvData
        End Get
    End Property
#End Region
#End Region

#Region "メソッド"
#Region "メソッド_接続"
    'サーバ側と接続を行う(他クラスとの窓口)
    'Parameters:
    '  ByVal strIpAddr As String：サーバ側のIPアドレス
    '  ByVal intPort As Integer：サーバ側のポート番号
    '  ByVal intTimeout As Integer：送受信タイムアウト時間(ms)

    Public Function mTcpConnect(ByVal strIpAddr As String, ByVal intPort As Integer, ByVal intTimeout As Integer) As Boolean
        Dim blnRslt As Boolean = True
        Try
            Call funcTcpConnect(strIpAddr, intPort, intTimeout)
        Catch ex As Exception
            blnRslt = False
            Return blnRslt
        End Try
        Return blnRslt
    End Function
#End Region
#Region "メソッド_切断"
    'サーバ側との接続を切断する(他クラスとの窓口)
    'Parameters:
    '  None
    'Returns:
    '  Boolean：True→成功、False→失敗
    Public Function mTcpClose() As Boolean
        Dim blnRslt As Boolean = True
        Try
            Call subTcpClose()
        Catch ex As Exception
            blnRslt = False
            Return blnRslt
        End Try
        Return blnRslt
    End Function
#End Region

#Region "メソッド_送信"
    'サーバ側へデータを送信する(他クラスとの窓口)
    'Parameters:
    '  ByVal bytSendData() As Byte：送信電文
    '  ByVal intSendDataLen As Integer：送信電文長
    'Returns:
    '  Boolean：True→成功、False→失敗
    Public Function mTcpSend(ByVal bytSendData() As Byte, ByVal intSendDataLen As Integer)
        Dim blnRslt As Boolean = True
        Try
            Call subTcpSend(bytSendData, intSendDataLen)
        Catch ex As Exception
            blnRslt = False
            Return blnRslt
        End Try
        Return blnRslt
    End Function
#End Region

#Region "メソッド_受信"
    'サーバ側へデータを送信する(他クラスとの窓口)
    'Parameters:
    '  None
    'Returns:
    '  Boolean：True→成功、False→失敗
    Public Function mTcpRcv() As Boolean
        Dim blnRslt As Boolean = True
        Try
            gbytRcvData = funcTcpRcv()
        Catch ex As Exception
            '受信データフラグを変更する
            'gintRcvStat = RcvStat.RCV_ERR
            blnRslt = False
            Return blnRslt
        End Try

        Return blnRslt
    End Function
#End Region
#End Region

#Region "関数"
#Region "関数_接続"
    'サーバ側と接続を行う
    'Parameters:
    '  ByVal strIpAddr As String：サーバ側のIPアドレス
    '  ByVal intPort As Integer：サーバ側のポート番号
    '  ByVal intTimeout As Integer：送受信タイムアウト時間(ms)
    'Returns:
    '  Boolean：True→成功、False→失敗
    Private Function funcTcpConnect(ByVal strIpAddr As String, ByVal intPort As Integer, ByVal intTimeout As Integer)
        Dim blnRslt As Boolean = True
        'TcpClientを作成し、サーバーと接続する
        Try
            gobjTcp = New System.Net.Sockets.TcpClient(strIpAddr, intPort)
            Console.WriteLine("サーバー({0}:{1})と接続しました({2}:{3})。", _
                    DirectCast(gobjTcp.Client.RemoteEndPoint, System.Net.IPEndPoint).Address, _
                    DirectCast(gobjTcp.Client.RemoteEndPoint, System.Net.IPEndPoint).Port, _
                    DirectCast(gobjTcp.Client.LocalEndPoint, System.Net.IPEndPoint).Address, _
                    DirectCast(gobjTcp.Client.LocalEndPoint, System.Net.IPEndPoint).Port)
        Catch ex As Exception
            blnRslt = False
            Return blnRslt
            Exit Function
        End Try
        'NetworkStreamを取得する
        gobjNtwStrm = gobjTcp.GetStream()
        '読み取り、書き込みのタイムアウトを10秒にする
        'デフォルトはInfiniteで、タイムアウトしない
        '(.NET Framework 2.0以上が必要)
        gobjNtwStrm.ReadTimeout = intTimeout
        gobjNtwStrm.WriteTimeout = intTimeout
        Return blnRslt
    End Function
#End Region
#Region "関数_切断"
    'サーバ側との接続を切断する
    'Parameters:
    '  None
    'Returns:
    '  None
    Private Sub subTcpClose()
        '閉じる
        If gobjNtwStrm IsNot Nothing Then
            gobjNtwStrm.Close()
        End If

        If gobjTcp IsNot Nothing Then
            gobjTcp.Close()
            Console.WriteLine("切断しました。")
        End If
    End Sub
#End Region
#Region "関数_送信"
    'サーバ側へデータを送信する
    'Parameters:
    '  ByVal bytSendData() As Byte：送信電文
    '  ByVal intSendDataLen As Integer：送信電文長
    'Returns:
    '  None
    Private Sub subTcpSend(ByVal bytSendData() As Byte, ByVal intSendDataLen As Integer)
        'データを送信する
        gobjNtwStrm.Write(bytSendData, 0, intSendDataLen)

        '受信データフラグを変更する
        gintRcvStat = RcvStat.RCV_MID
        '受信データ変数を初期化
        gbytRcvData = New Byte(255) {}
        gobjMemStrm = New System.IO.MemoryStream()
    End Sub
#End Region

#Region "関数_受信"
    'サーバーから送られたデータを受信する
    'Parameters:
    '  None
    'Returns:
    '  Byte()：受信データ
    Private Function funcTcpRcv() As Byte()
        Dim bytRcvBuf(32768) As Byte      '一時バッファ
        Dim intRcvBufLen As Integer = 0
        Dim blnRslt As Boolean

        'データの一部を受信する
        intRcvBufLen = gobjNtwStrm.Read(bytRcvBuf, 0, bytRcvBuf.Length)
        'Readが0を返した時はサーバーが切断したと判断
        If intRcvBufLen = 0 Then
            Console.WriteLine("サーバーが切断しました。")
            '受信データフラグを変更する
            gintRcvStat = RcvStat.RCV_ERR
            Return gobjMemStrm.ToArray
            Exit Function
        End If

        '受信したデータを蓄積する
        gobjMemStrm.Write(bytRcvBuf, 0, intRcvBufLen)

        '受信完了確認
        blnRslt = funcTcpRcvLenChk(intRcvBufLen)
        If blnRslt = True Then
            '受信データフラグを変更する
            gintRcvStat = RcvStat.RCV_END
        End If

        Return gobjMemStrm.ToArray
    End Function
#End Region
#Region "関数_受信完了確認"
    '全てのデータを受信したかを確認する
    'Parameters:
    '  ByVal intDataLen As Integer：受信データ長
    'Returns:
    '  Boolean：True→完了、False→未受信データあり
    Private Function funcTcpRcvLenChk(ByVal intDataLen As Integer) As Boolean
        Dim blnRslt As Boolean = False
        Dim intTrasLen As Integer      '転送バイト数

        '受信データが6バイト未満であれば、Falseで抜ける
        If intDataLen < 6 Then
            Return blnRslt
            Exit Function
        End If

        '転送バイト数（＝転送バイト数より後に何バイトのデータがあるかを示す）を取得する
        intTrasLen = (CInt(gbytRcvData(4)) << 8)
        intTrasLen = intTrasLen + gbytRcvData(5)
        '期待する受信総バイト数を取得する（6＝転送ID[2バイト]+プロトコルID[2バイト]+転送バイト数[2バイト]）
        intTrasLen = intTrasLen + 6

        '全データを受信済かを確認する
        If intDataLen >= intTrasLen Then
            '全データ受信完了
            blnRslt = True
        End If
        Return blnRslt
    End Function
#End Region
#End Region

End Class
