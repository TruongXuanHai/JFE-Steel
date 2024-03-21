Imports System.Threading

Public Class clsThreadPara
    Public _gwIndex As Integer
    Public _unitIndex As Integer
    Public _channelIndex As Integer
    Private ReadOnly _cancellationToken As CancellationToken
    Public lockObject As New Object()

    Public ReadOnly Property CancellationToken As CancellationToken
        Get
            Return _cancellationToken
        End Get
    End Property

    Public Sub New(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer, ByVal cancellationToken As CancellationToken)
        _gwIndex = gwIndex
        _unitIndex = unitIndex
        _channelIndex = channelIndex
        _cancellationToken = cancellationToken
    End Sub

    Public Sub Update(ByVal gwIndex As Integer, ByVal unitIndex As Integer, ByVal channelIndex As Integer)
        SyncLock lockObject
            _gwIndex = gwIndex
            _unitIndex = unitIndex
            _channelIndex = channelIndex
        End SyncLock
    End Sub

    Public Function GetValues() As clsThreadPara
        SyncLock lockObject
            Return New clsThreadPara(_gwIndex, _unitIndex, _channelIndex, Me.CancellationToken)
        End SyncLock
    End Function
End Class