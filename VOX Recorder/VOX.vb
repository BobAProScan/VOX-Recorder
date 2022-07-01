Public Class VOX
    Implements IDisposable

    Private _IsLatched As Boolean
    Private _MaxValue As Integer = 100
    Private _ThreasholdLevel As Integer
    Private _ThreasholdTrigger As Boolean

    Public Event UpdateGUI(Code As Integer, Data As Object)

    Public Sub InputAudio(Value As Integer)

        If ProgramFinishedLoading = False Then Return

        ThreasholdProcess(Value)
        VOXProcess()

    End Sub

    Private Sub ThreasholdProcess(value As Integer)
        'controls green light

        Static Temp As Boolean

        If value >= _ThreasholdLevel Then
            _ThreasholdTrigger = True
        Else
            _ThreasholdTrigger = False
        End If

        If Temp <> _ThreasholdTrigger Then
            RaiseEvent UpdateGUI(2, Nothing)
        End If
        Temp = _ThreasholdTrigger

    End Sub

    Private Sub VOXProcess()
        'controls red light

        Static starttime As Date

        If _ThreasholdTrigger = True Then
            starttime = Now
            If _IsLatched = False Then
                _IsLatched = True
                RaiseEvent UpdateGUI(3, Nothing)
            End If
        End If

        If _IsLatched = True Then
            Dim diff As TimeSpan = Now.Subtract(starttime)
            If diff.TotalSeconds >= VOXDELAY Then
                _IsLatched = False
                RaiseEvent UpdateGUI(3, Nothing)
            End If
        End If

    End Sub

    Public Property ThreasholdLevel() As Integer

        Get
            Return _ThreasholdLevel
        End Get
        Set(Value As Integer)
            If Value > _MaxValue Then Value = _MaxValue
            If Value < 0 Then Value = 0
            If Value = ThreasholdLevel Then Exit Property
            _ThreasholdLevel = Value
        End Set

    End Property

    Public ReadOnly Property IsThreasholdTriggered() As Boolean

        Get
            Return _ThreasholdTrigger
        End Get

    End Property

    Public ReadOnly Property IsLatched As Boolean

        Get
            Return _IsLatched
        End Get

    End Property

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                'free managed resources when explicitly called
            End If

            'free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
