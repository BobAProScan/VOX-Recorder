<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCheckNewestVersion
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(96, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Version 1.0 Now Available"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button3
        '
        Me.Button3.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button3.Location = New System.Drawing.Point(333, 167)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(80, 30)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.AllowDrop = True
        Me.Button1.AutoSize = True
        Me.Button1.Location = New System.Drawing.Point(109, 58)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(215, 30)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Download And Update Now"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.AllowDrop = True
        Me.Button2.AutoSize = True
        Me.Button2.Location = New System.Drawing.Point(109, 105)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(215, 30)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "View Change Log"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'FrmCheckNewestVersion
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.CancelButton = Me.Button3
        Me.ClientSize = New System.Drawing.Size(432, 216)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "FrmCheckNewestVersion"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Check For Newest Version"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents Button3 As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Button1 As Button
    Public WithEvents Button2 As Button
End Class
