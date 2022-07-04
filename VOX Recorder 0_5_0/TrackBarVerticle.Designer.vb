<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TrackBarVerticle
    Inherits System.Windows.Forms.UserControl

    Public Sub New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw, True)
        Me.UpdateStyles()

    End Sub

    Protected Overloads Overrides Sub Dispose(disposing As Boolean)
        On Error Resume Next
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New Buttons()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BorderRadius = 3
        Me.Button1.Color1 = System.Drawing.Color.LightBlue
        Me.Button1.Color2 = System.Drawing.Color.CornflowerBlue
        Me.Button1.ColorBorder = System.Drawing.Color.CornflowerBlue
        Me.Button1.ColorIcon = System.Drawing.Color.WhiteSmoke
        Me.Button1.FocusedEnabled = True
        Me.Button1.Location = New System.Drawing.Point(0, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 17)
        Me.Button1.TabIndex = 2
        '
        'TrackBarVerticle
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.Button1)
        Me.Name = "TrackBarVerticle"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Buttons

End Class
