' HPA-40
' getXmlData.vb
' XMLファイルから設定値を取得
'
'CORYRIGHT(C) 2023 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2023/12/18 チュオンスアンハイ

Imports System.Xml
Imports System.Xml.Linq

Module getXmlData
    ' ユニット設定用
    Public Class UnitXmlData
        Public Property IPAddress As String
        Public Property LoRaAddress As String
        Public Property ModbusAddress As String
        Public Property StartRegister As String
        Public Property RegisterNumber As String

        Public Property GateWayNo As String
        Public Property UnitNo As String
        Public Property UnitName As String
        Public Property ChannelNo As String
        Public Property ChannelName As String

        Public Property GateWayNumber As String
        Public Property UnitNumber As String
        Public Property ChannelNumber As String
        Public Property Cycle As String
    End Class

    ' エッジサーバ設定用
    Public Class ServerXmlData
        Public Property IPAddress As String
        Public Property UserName As String
        Public Property PassWord As String
        Public Property RemoteDirectory As String
    End Class

    Function GetSettingParameter(filePath As String, indexGW As Integer, indexUnit As Integer, indexChannel As Integer) As UnitXmlData
        ' UnitXmlDataクラス変数初期
        Dim data As New UnitXmlData()

        ' XMLドキュメントを作成し、ローディングする
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(filePath)

        ' エレメントを取る
        Dim ipAddressNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/IPAddress")
        Dim loraAddressNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/LoRaAddress")
        Dim modbusAddressNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/RS485Address")
        Dim startRegisterNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/StartRegister")
        Dim registerNumberNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/RegisterNumber")
        Dim gatewayNoNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/GWNo")
        Dim unitNoNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/UnitNo")
        Dim unitNameNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/UnitName")
        Dim channelNoNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/Channel/CH[" & indexChannel & "]/CHNo")
        Dim channelNameNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/Channel/CH[" & indexChannel & "]/CHName")
        Dim gwNumberNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GWNumber")
        Dim unitNumberNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/UnitNumber")
        Dim channelNumberNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Unit[" & indexUnit & "]/ChannelNumber")
        Dim circleNode As XmlNode = xmlDoc.SelectSingleNode("/settings/GW[" & indexGW & "]/Cycle")
        If ipAddressNode IsNot Nothing Then
            data.IPAddress = ipAddressNode.InnerText
        End If
        If loraAddressNode IsNot Nothing Then
            data.LoRaAddress = loraAddressNode.InnerText
        End If
        If modbusAddressNode IsNot Nothing Then
            data.ModbusAddress = modbusAddressNode.InnerText
        End If
        If startRegisterNode IsNot Nothing Then
            data.StartRegister = startRegisterNode.InnerText
        End If
        If registerNumberNode IsNot Nothing Then
            data.RegisterNumber = registerNumberNode.InnerText
        End If
        If gatewayNoNode IsNot Nothing Then
            data.GateWayNo = gatewayNoNode.InnerText
        End If
        If unitNoNode IsNot Nothing Then
            data.UnitNo = unitNoNode.InnerText
        End If
        If unitNameNode IsNot Nothing Then
            data.UnitName = unitNameNode.InnerText
        End If
        If channelNoNode IsNot Nothing Then
            data.ChannelNo = channelNoNode.InnerText
        End If
        If channelNameNode IsNot Nothing Then
            data.ChannelName = channelNameNode.InnerText
        End If
        If gwNumberNode IsNot Nothing Then
            data.GateWayNumber = gwNumberNode.InnerText
        End If
        If unitNumberNode IsNot Nothing Then
            data.UnitNumber = unitNumberNode.InnerText
        End If
        If channelNumberNode IsNot Nothing Then
            data.ChannelNumber = channelNumberNode.InnerText
        End If
        If circleNode IsNot Nothing Then
            data.Cycle = circleNode.InnerText
        End If
        Return data
    End Function

    Function GetServerParameter(filePath As String) As ServerXmlData

        Dim data As New ServerXmlData()

        ' XMLドキュメントを作成し、ローディングする
        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load(filePath)

        ' エレメントを取る
        Dim ipAddressNode As XmlNode = xmlDoc.SelectSingleNode("/settings/IPAddress")
        Dim usernameNode As XmlNode = xmlDoc.SelectSingleNode("/settings/Username")
        Dim passwordNode As XmlNode = xmlDoc.SelectSingleNode("/settings/Password")
        Dim remoteDirectoryNode As XmlNode = xmlDoc.SelectSingleNode("/settings/RemoteDirectory")

        If ipAddressNode IsNot Nothing Then
            data.IPAddress = ipAddressNode.InnerText
        End If
        If usernameNode IsNot Nothing Then
            data.UserName = usernameNode.InnerText
        End If
        If passwordNode IsNot Nothing Then
            data.PassWord = passwordNode.InnerText
        End If
        If remoteDirectoryNode IsNot Nothing Then
            data.RemoteDirectory = remoteDirectoryNode.InnerText
        End If

        Return data
    End Function
End Module
