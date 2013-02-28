Imports System
Imports System.Runtime.InteropServices

Namespace picotech
    Public Class PinnedArray(Of T)
        Implements IDisposable

        Public _pinnedHandle As GCHandle
        Private _channelCount As Integer

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(array As T())
            MyClass.New()
            _pinnedHandle = GCHandle.Alloc(array, GCHandleType.Pinned)
        End Sub

        Public Sub New(arraySize As Integer)
            MyClass.New(New T(arraySize) {})
        End Sub

        Public ReadOnly Property AddrOfPinnedObject() As IntPtr
            Get
                Return _pinnedHandle.AddrOfPinnedObject()
            End Get
        End Property

        Public ReadOnly Property Target As T()
            Get
                Return CType(_pinnedHandle.Target, T())
            End Get
        End Property

        Default Public Property Item(ByVal index As Integer) As T
            Get

            End Get
            Set(value As T)

            End Set
        End Property

        Public Shared Widening Operator CType(ByVal a As PinnedArray(Of T)) As T()
            If IsNothing(a) Then
                Return Nothing
            End If

            Return TryCast(a._pinnedHandle.Target, T())
        End Operator

#Region "IDisposable Support"
        Private disposedValue As Boolean ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Me.disposedValue Then Return
            Me.disposedValue = True

            If disposing Then
                ' TODO: マネージ状態を破棄します (マネージ オブジェクト)。
            End If

            ' TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
            ' TODO: 大きなフィールドを null に設定します。
            _pinnedHandle.Free()
        End Sub

        ' TODO: 上の Dispose(ByVal disposing As Boolean) にアンマネージ リソースを解放するコードがある場合にのみ、Finalize() をオーバーライドします。
        'Protected Overrides Sub Finalize()
        '    ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Namespace
