<System.ComponentModel.DefaultEvent("Click")>
Public Class Buttons

    Private _IsMouseDown As Boolean = False
    Private _IsMouseOver As Boolean = False
    Private _IsFocused As Boolean = False
    Private _FocusedEnabled As Boolean = False
    Private _BorderRadius As Integer = 10
    Private _Color1 As Color = Color.SkyBlue
    Private _Color2 As Color = Color.RoyalBlue
    Private _ColorBorder As Color = Color.DimGray
    Private _ColorIcon As Color = Color.WhiteSmoke

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        SetStyle(ControlStyles.SupportsTransparentBackColor Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw, True)
        Me.UpdateStyles()

        MyBase.Size = New Size(50, 25) 'sets default size when control placed on form

    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        Dim gp As New GraphicsPath

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        If _BorderRadius < 1 Then _BorderRadius = 1

        gp.AddArc(New Rectangle(0, 0, _BorderRadius, _BorderRadius), 180, 90)
        gp.AddArc(New Rectangle(Width - _BorderRadius - 1, 0, _BorderRadius, _BorderRadius), -90, 90)
        gp.AddArc(New Rectangle(Width - _BorderRadius - 1, Height - _BorderRadius - 1, _BorderRadius, _BorderRadius), 0, 90)
        gp.AddArc(New Rectangle(0, Height - _BorderRadius - 1, _BorderRadius, _BorderRadius), 90, 90)
        gp.CloseFigure()

        e.Graphics.FillPath(LGB1, gp)
        e.Graphics.DrawPath(New Pen(_ColorBorder, 1), gp)

        If _FocusedEnabled AndAlso _IsFocused Then
            ControlPaint.DrawFocusRectangle(e.Graphics, New Rectangle(ClientRectangle.X + 4, ClientRectangle.Y + 4, ClientRectangle.Width - 8, ClientRectangle.Height - 8))
        End If

        MyBase.OnPaint(e)

    End Sub

    Private Function LGB1() As LinearGradientBrush

        Dim lgb As LinearGradientBrush = Nothing

        Dim color_1 As Color = _Color1
        Dim color_2 As Color = _Color2
        If _IsMouseOver = True Then
            color_1 = ControlPaint.Light(color_1)
        End If
        lgb = New LinearGradientBrush(Point.Empty, New Point(0, Me.Height), color_2, color_1)
        Dim cb As New ColorBlend
        cb.Colors = New Color() {color_2, color_1, color_1, color_2}
        If _IsMouseDown = False Then
            cb.Positions = New Single() {0.0F, 0.2F, 0.3F, 1.0F}
        Else
            cb.Positions = New Single() {0.0F, 0.7F, 0.8F, 1.0F}
        End If
        lgb.InterpolationColors = cb

        Return lgb

    End Function

    Protected Overrides Sub OnMouseMove(e As System.Windows.Forms.MouseEventArgs)

        MyBase.OnMouseMove(e)

        If e.Button = Windows.Forms.MouseButtons.None Or e.Button = Windows.Forms.MouseButtons.Left Then
            If e.Button = Windows.Forms.MouseButtons.None Then
                _IsMouseDown = False
            End If
            _IsMouseOver = True
        Else
            If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                _IsMouseOver = False
            Else
                _IsMouseOver = True
            End If
        End If

        If e.Button = Windows.Forms.MouseButtons.Left Then
            If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                _IsMouseDown = False
            Else
                _IsMouseDown = True
            End If
        End If
        Invalidate()

    End Sub

    Protected Overrides Sub OnMouseLeave(e As System.EventArgs)

        MyBase.OnMouseLeave(e)

        _IsMouseOver = False
        _IsMouseDown = False

        Invalidate()

    End Sub

    Protected Overrides Sub OnMouseDown(e As System.Windows.Forms.MouseEventArgs)

        MyBase.OnMouseDown(e)

        If e.Button = Windows.Forms.MouseButtons.Left Then
            If New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                _IsMouseDown = True
            Else
                _IsMouseDown = False
            End If

            Focus()
        End If

        Invalidate()

    End Sub

    Protected Overrides Sub OnMouseUp(e As System.Windows.Forms.MouseEventArgs)

        MyBase.OnMouseUp(e)

        If e.Button = Windows.Forms.MouseButtons.Left Then
            _IsMouseDown = False
        End If

        Invalidate()

    End Sub

    Protected Overrides Sub OnEnter(e As System.EventArgs)

        MyBase.OnEnter(e)

        _IsFocused = True

        Invalidate()

    End Sub

    Protected Overrides Sub OnLeave(e As System.EventArgs)

        MyBase.OnLeave(e)

        _IsFocused = False

        Invalidate()

    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)

        MyBase.OnKeyDown(e)

        If e.KeyCode = Keys.Space Then
            _IsMouseDown = True
        End If

        Invalidate()

    End Sub

    Protected Overrides Sub OnKeyUp(e As System.Windows.Forms.KeyEventArgs)

        MyBase.OnKeyUp(e)

        If e.KeyCode = 32 Then
            _IsMouseDown = False
            MyBase.OnClick(e)
        End If

        Invalidate()

    End Sub

    Protected Overrides Sub OnTextChanged(e As System.EventArgs)

        MyBase.OnTextChanged(e)
        Invalidate()

    End Sub

    Protected Overrides Sub OnEnabledChanged(e As System.EventArgs)

        MyBase.OnEnabledChanged(e)
        Invalidate()

    End Sub

    Public Property BorderRadius As Integer

        Get
            Return _BorderRadius
        End Get
        Set(Value As Integer)
            _BorderRadius = Value
            Invalidate()
        End Set

    End Property

    Public Property Color1 As Color

        Get
            Return _Color1
        End Get
        Set(Value As Color)
            _Color1 = Value
            Invalidate()
        End Set

    End Property

    Public Property Color2 As Color

        Get
            Return _Color2
        End Get
        Set(Value As Color)
            _Color2 = Value
            Invalidate()
        End Set

    End Property

    Public Property ColorBorder As Color

        Get
            Return _ColorBorder
        End Get
        Set(Value As Color)
            _ColorBorder = Value
            Invalidate()
        End Set

    End Property

    Public Property ColorIcon As Color

        Get
            Return _ColorIcon
        End Get
        Set(Value As Color)
            _ColorIcon = Value
            Invalidate()
        End Set

    End Property

    Public Property FocusedEnabled As Boolean

        Get
            Return _FocusedEnabled
        End Get
        Set(Value As Boolean)
            _FocusedEnabled = Value
        End Set

    End Property

End Class
