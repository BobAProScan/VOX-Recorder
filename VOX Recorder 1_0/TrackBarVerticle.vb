Public Class TrackBarVerticle
    Inherits UserControl

    Private _MaxValue As Integer = 100
    Private _Value As Integer
    Private _SmallChangePercent As Integer = 5
    Private _DragModeEnabled As Boolean
    Private _StartDragPoint As Point
    Private _UnLocked As Boolean = True

    Public Event ValueChanged As EventHandler

    Private Sub UserControl1_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize

        Button1.Left = CInt((Me.Width / 2) - ((Button1.Width) - (Button1.Width / 2)))
        Invalidate()

    End Sub

    Private Sub MyTrackBarAudio_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Dim pen1 As New Pen(SystemColors.ControlDark)

        e.Graphics.DrawLine(pen1, CInt(Me.Width / 2) - 2, 1, CInt(Me.Width / 2) - 2, Me.Height - 2)
        pen1.Color = SystemColors.ControlLightLight
        pen1.Width = 2
        e.Graphics.DrawLine(pen1, CInt(Me.Width / 2), 1, CInt(Me.Width / 2), Me.Height - 2)

        pen1.Dispose()

    End Sub

    Private Sub MoveSlider(delta As Integer)

        If _UnLocked = False Then Return

        If delta < 0 AndAlso (Button1.Top + delta) <= 0 Then
            Button1.Top = 0
        Else
            If delta > 0 AndAlso (Button1.Top + Button1.Height + delta) >= Me.Height Then
                Button1.Top = Me.Height - Button1.Height
            Else
                Button1.Top += delta
            End If
        End If

        Dim distance As Integer = Me.Height - Button1.Height
        Dim percent As Single = CType(Button1.Top, Single) / CType(distance, Single)
        Dim temp As Integer
        If _MaxValue < 0 Then temp = _MaxValue * -1 Else temp = _MaxValue
        Dim movement As Integer = Convert.ToInt32(percent * CType((temp), Single))

        _Value = CInt(IIf((_MaxValue >= 0), _MaxValue - movement, _MaxValue + movement))

        RaiseEvent ValueChanged(Me, New EventArgs)
        Invalidate()

    End Sub

    Private Sub MoveSlider()

        If _UnLocked = False Then Return

        Dim distance As Integer
        If _MaxValue < 0 Then distance = _MaxValue * -1 Else distance = _MaxValue
        Dim percent As Single = CType(_Value, Single) / CType(distance, Single)
        If Single.IsNaN(percent) Then Return

        Button1.Top = Me.Height - Button1.Height - Convert.ToInt32(percent * CType((Me.Height - Button1.Height), Single))
        Invalidate()

    End Sub

    Private Sub Button1_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseDown

        If _UnLocked = False Then Return

        _DragModeEnabled = True

        _StartDragPoint = New Point(e.X, e.Y)

    End Sub

    Private Sub Button1_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseUp

        If _UnLocked = False Then Return

        _DragModeEnabled = False

    End Sub

    Private Sub Button1_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseMove

        If _UnLocked = False Then Return

        If _DragModeEnabled = False Then Return

        Dim delta As Integer = e.Y - _StartDragPoint.Y

        If delta = 0 Then Return

        MoveSlider(delta)

    End Sub

    Private Sub MyTrackBar_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

        If _UnLocked = False Then Return

        Dim delta As Integer = e.Y - Button1.Top - Button1.Height

        If delta = 0 Then Return

        MoveSlider(delta)

    End Sub

    Private Sub MyTrackBar_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp

        If _UnLocked = False Then Return

    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, keyData As System.Windows.Forms.Keys) As Boolean

        If _UnLocked = False Then Return False

        Select Case keyData
            Case Keys.Down, Keys.Left
                If _SmallChangePercent = 1 Then
                    _Value = _Value - 1
                Else
                    _Value = _Value - CInt((_SmallChangePercent / _MaxValue) * (_MaxValue))
                End If
                If _Value < 0 Then _Value = 0
                MoveSlider()
                RaiseEvent ValueChanged(Me, New EventArgs)
            Case Keys.Up, Keys.Right
                If _SmallChangePercent = 1 Then
                    _Value = _Value + 1
                Else
                    _Value = _Value + CInt((_SmallChangePercent / _MaxValue) * (_MaxValue))
                End If
                If _Value > _MaxValue Then _Value = _MaxValue
                MoveSlider()
                RaiseEvent ValueChanged(Me, New EventArgs)
            Case Keys.Home
                _Value = 0
                MoveSlider()
                RaiseEvent ValueChanged(Me, New EventArgs)
            Case Keys.End
                _Value = _MaxValue
                MoveSlider()
                RaiseEvent ValueChanged(Me, New EventArgs)
        End Select

        Return False

    End Function

    Public Property Value() As Integer

        Get
            Return _Value
        End Get
        Set(Value As Integer)
            If Value > _MaxValue Then Value = _MaxValue
            If Value < 0 Then Value = 0
            _Value = Value
            MoveSlider()
        End Set

    End Property

    Public Property Maximum() As Integer

        Get
            Return _MaxValue
        End Get
        Set(Value As Integer)
            _MaxValue = Value
            If Value < 0 Then Value = 0
            MoveSlider()
        End Set

    End Property

    Public Property SmallChangePercent() As Integer

        Get
            Return _SmallChangePercent
        End Get
        Set(Value As Integer)
            If Value > _MaxValue Then Value = _MaxValue
            If Value < 0 Then Value = 0
            _SmallChangePercent = Value
        End Set

    End Property

    Public Property UnLocked() As Boolean

        Get
            Return _UnLocked
        End Get
        Set(Value As Boolean)
            _UnLocked = Value
        End Set

    End Property

    Public Property ButtonBorderRadius As Integer

        Get
            Return _Button1.BorderRadius
        End Get
        Set(Value As Integer)
            _Button1.BorderRadius = Value
            Invalidate()
        End Set

    End Property

    Public Property ButtonColor1 As Color

        Get
            Return _Button1.Color1
        End Get
        Set(Value As Color)
            _Button1.Color1 = Value
            Invalidate()
        End Set

    End Property

    Public Property ButtonColor2 As Color

        Get
            Return _Button1.Color2
        End Get
        Set(Value As Color)
            _Button1.Color2 = Value
            Invalidate()
        End Set

    End Property

    Public Property ButtonColorBorder As Color

        Get
            Return _Button1.ColorBorder
        End Get
        Set(Value As Color)
            _Button1.ColorBorder = Value
            Invalidate()
        End Set

    End Property

    Public Property ButtonColorIcon As Color

        Get
            Return _Button1.ColorIcon
        End Get
        Set(Value As Color)
            _Button1.ColorIcon = Value
            Invalidate()
        End Set

    End Property

End Class
