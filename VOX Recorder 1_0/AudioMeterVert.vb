Public Class AudioMeterVert
    Inherits UserControl

    Private _MaxValue As Integer = 90
    Private _Value As Integer
    Private _BarHeight As Integer
    Private _MeterBrush As LinearGradientBrush
    Private _ColorBlend As New ColorBlend(3)

    Private Sub AudioMeterVert_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        _ColorBlend.Colors = New Color() {Color.Lime, Color.Yellow, Color.Red}
        _ColorBlend.Positions = New Single() {0.0F, 0.8F, 1.0F}
        _MeterBrush = New LinearGradientBrush(New Point(0, Me.Height), New Point(0, 0), Color.Lime, Color.Red)
        _MeterBrush.InterpolationColors = _ColorBlend

    End Sub

    Private Sub AudioMeterVert_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        '_Value = 1 '100 'FOR TESTING

        _BarHeight = CInt((Me.Height - 5) * CSng(_Value / _MaxValue))
        e.Graphics.FillRectangle(_MeterBrush, 2, Me.Height - _BarHeight - 4, Me.Width - 9, _BarHeight - 1)

        ''draws divider lines
        'For i As Integer = Me.Height - 8 To 0 Step -4
        '    e.Graphics.DrawLine(Pens.Red, 2, i, Me.Width - 2, i)
        'Next

    End Sub

    Protected Overrides Sub OnResize(e As System.EventArgs)

        Invalidate()

    End Sub

    Public Property Value() As Integer

        Get
            Value = _Value
        End Get
        Set(Value As Integer)
            If Value > _MaxValue Then Value = _MaxValue
            If Value < 0 Then Value = 0
            If Value = _Value Then Exit Property
            _Value = Value

            Invalidate()
        End Set

    End Property

End Class
