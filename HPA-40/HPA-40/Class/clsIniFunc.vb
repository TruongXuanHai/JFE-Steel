' HPA-40
' clsIniFunc.vb
' INIファイル制御クラス定義
'
'CORYRIGHT(C) 2023 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2023/12/18 チュオンスアンハイ

Imports System
Imports System.Text
Imports System.IO
Imports System.Runtime.InteropServices


' INIファイル制御クラス
' 相対パスを用いてAPIを呼び出すと、セットしたパスが書き換えられてしまうので、
' パスを一時退避してから呼び出すようにする
Public Class clsIniFunc

    '' 文字列を読み出す
    ' 結果を得る文字列は、StringBuilder型であることに注意
    ' 文字列を受け取るために、StringBuilder型のインスタンスを用意しておくこと
    ' 例： Dim sb As StringBuilder = New StringBuilder(1024)
    '      CIniFunc.GetPrivateProfileString("アプリ１","キー１","default",sb,sb.Capacity,"c:\sample.ini")
    '　　　　　　　　　　
    Public Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpAppName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, _
     ByVal nSize As Integer, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As Integer

    '' 指定セクションのキーの一覧を得る
    Public Declare Function GetPrivateProfileStringByByteArray Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpAppName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As Byte(), _
     ByVal nSize As Integer, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As Integer

    '' 整数値を読み出す 
    Public Declare Function GetPrivateProfileInt Lib "KERNEL32.DLL" Alias "GetPrivateProfileIntA" ( _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpAppName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, _
     ByVal nDefault As Integer, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As Integer

    '' キーと値を書き加える（削除にも利用可）
    ' 第３パラメータにnullまたはNothingを指定すると、指定キーとその値が削除される
    ' 第２パラメータにnullまたはNothingを指定すると、指定セクション内の全てのキーとその値が削除される
    Public Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" ( _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpAppName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpString As String, _
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As Integer

End Class
