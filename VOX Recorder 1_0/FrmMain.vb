
Public Class FrmMain

    Public WithEvents BCCalls1 As BCCalls
    Public WithEvents VOX1 As VOX
    Public SaveConfigTimer As System.Timers.Timer
    Private WithEvents WaveIn1 As WaveIn

    Private Delegate Sub UpdateUIDelegate(Code As Integer, Data As Object)

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        d("Program Started - Time: " & Now.ToShortTimeString)

        fMain = Me

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Catch ex As Exception
        End Try
        CheckForIllegalCrossThreadCalls = True
        Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Me.MinimumSize = New Size(911, 578)
        FormatStatusLogGrid()
        SetupAbout()
        VOX1 = New VOX
        ReadConfig()
        StartWaveIn()
        SaveConfigTimer = New System.Timers.Timer(5000)
        SaveConfigTimer.SynchronizingObject = Me 'required because invokerequired comes back true
        SaveConfigTimer.AutoReset = False
        AddHandler SaveConfigTimer.Elapsed, AddressOf SaveConfigTimerEvent
        Cursor = Cursors.Arrow
        Application.DoEvents()

        ProgramFinishedLoading = True


#If Not DEBUG Then
        Tone_GroupBox1.Visible = False
#End If

    End Sub

    Private Sub SaveConfigTimerEvent(source As Object, e As ElapsedEventArgs)

        WriteConfig()

    End Sub

    Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize

        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub FrmMain_Move(sender As Object, e As EventArgs) Handles Me.Move

        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub FrmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If TP1_Button3.Text = "Stop Recorder" Then
            Using New CenteredMessageBox(Me)
                If MessageBox.Show(Me, "Stop Recorder" & vbCr & vbCr & "Are You Sure?", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Cancel Then
                    e.Cancel = True
                    Return
                End If
            End Using
        End If

        ProgramClosing = True
        WriteConfig()
        If WaveIn1 IsNot Nothing Then
            WaveIn1.Dispose()
            WaveIn1 = Nothing
        End If
        If BCCalls1 IsNot Nothing Then
            BCCalls1.Dispose()
            BCCalls1 = Nothing
        End If
        If VOX1 IsNot Nothing Then
            VOX1.Dispose()
            VOX1 = Nothing
        End If
        If PlayMP3_NAudioStreaming1 IsNot Nothing Then
            PlayMP3_NAudioStreaming1.Dispose()
            PlayMP3_NAudioStreaming1 = Nothing
        End If

        d("Program Closed - Time: " & Now.ToShortTimeString)

    End Sub

    Public Sub TP1_Button3_Click(sender As Object, e As EventArgs) Handles TP1_Button3.Click

        If TP1_Button3.Text = "Start Recorder" Then
            If CheckFields() = False Then Return
            Dim errorStr As String = HTTP.Send(TP1_TextBox1.Text, TP1_TextBox2.Text, String.Empty, String.Empty, String.Empty, String.Empty, True, Nothing)
            errorStr = errorStr.Replace("OK", String.Empty)
            If errorStr <> String.Empty Then
                WriteLogs("Connection failed - Error: " & errorStr)
                Using New CenteredMessageBox(Me)
                    MessageBox.Show(Me, "Connection failed - Error: " & errorStr, ProgramName)
                End Using
                Return
            End If
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()
            Dim client As New WebClient
            Dim downloadString As String = client.DownloadString("https://www.broadcastify.com/calls/node/" & TP1_TextBox2.Text)
            Dim int1 As Integer = downloadString.IndexOf("<title>Node " & TP1_TextBox2.Text)
            If int1 > 0 Then
                Dim int2 As Integer = downloadString.IndexOf(":", int1)
                If int2 > 0 Then
                    Dim int3 As Integer = downloadString.IndexOf(" - Live</title>")
                    TP1_Label4.Text = downloadString.Substring(int2 + 2, int3 - int2 - 2)
                End If
            End If
            TP1_Button3.Text = "Stop Recorder"
            BCCalls1 = New BCCalls
            BCCalls1.Start()
            EnableControls(False)
            Me.Cursor = Cursors.Arrow
        ElseIf TP1_Button3.Text = "Stop Recorder" Then
            Using New CenteredMessageBox(Me)
                If MessageBox.Show(Me, "Stop Recorder" & vbCr & vbCr & "Are You Sure?", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Cancel Then
                    Return
                End If
            End Using
            Me.Cursor = Cursors.WaitCursor
            TP1_Button3.Text = "Start Recorder"
            If BCCalls1 IsNot Nothing Then
                BCCalls1.Dispose()
                BCCalls1 = Nothing
            End If
            EnableControls(True)
            Me.Cursor = Cursors.Arrow
        End If

    End Sub
    Private Sub EnableControls(Enable As Boolean)

        TP1_TextBox1.Enabled = Enable
        TP1_TextBox2.Enabled = Enable
        TP1_TextBox4.Enabled = Enable
        TP1_TextBox5.Enabled = Enable

    End Sub

    Private Sub TP1_Button4_Click(sender As Object, e As EventArgs) Handles TP1_Button4.Click

        If Directory.Exists(ProgramPath & "Recordings") = False Then
            Directory.CreateDirectory(ProgramPath & "Recordings")
        End If

        Process.Start("explorer.exe", ProgramPath & "Recordings")

    End Sub

    Private Sub TP1_TextBox1_2_4_5_TextChanged(sender As Object, e As EventArgs) Handles TP1_TextBox1.TextChanged, TP1_TextBox2.TextChanged, TP1_TextBox4.TextChanged, TP1_TextBox5.TextChanged

        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub TP1_CheckBox1_2_3_CheckedChanged(sender As Object, e As EventArgs) Handles TP1_CheckBox1.CheckedChanged, TP1_CheckBox2.CheckedChanged, TP1_CheckBox3.CheckedChanged

        Me.TopMost = TP1_CheckBox2.Checked
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub Audio_RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles Audio_RadioButton1.CheckedChanged

        If WaveIn1 IsNot Nothing Then
            If Audio_RadioButton1.Checked Then
                WaveIn1.Channel = 1
            ElseIf Audio_RadioButton2.Checked Then
                WaveIn1.Channel = 2
            End If
            ClearAudioControls()
            If ProgramFinishedLoading = True Then SaveConfigTimer.Start()
        End If

    End Sub

    Private Sub Audio_Button1_Click(sender As Object, e As EventArgs) Handles Audio_Button1.Click

        Try
            If Environment.OSVersion.Version.Major < 6 Then
                Process.Start("sndvol32")
            Else 'vista or above
                Process.Start("control.exe", "mmsys.cpl ,1") 'recording tab
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Audio_VOXTrackBar1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles Audio_VOXTrackBar1.ValueChanged

        If VOX1 IsNot Nothing Then
            VOX1.ThreasholdLevel = Audio_VOXTrackBar1.Value
        End If
        Audio_Label9.Text = Audio_VOXTrackBar1.Value.ToString
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub Audio_VOXPictureBox1_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Audio_VOXPictureBox1.Paint

        Dim width As Integer = Audio_VOXPictureBox1.Width - 1
        Dim height As Integer = Audio_VOXPictureBox1.Height - 1

        If VOX1 IsNot Nothing AndAlso VOX1.IsThreasholdTriggered = True AndAlso Audio_VOXPictureBox1.Enabled = True Then
            e.Graphics.FillRectangle(New SolidBrush(Color.Lime), 0, 0, width, height)
        Else
            e.Graphics.FillRectangle(New SolidBrush(Color.DarkGreen), 0, 0, width, height)
        End If

        e.Graphics.DrawRectangle(New Pen(Color.Black), 0, 0, width, height)

    End Sub

    Private Sub Audio_VOXPictureBox2_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Audio_VOXPictureBox2.Paint

        Dim width As Integer = Audio_VOXPictureBox2.Width - 1
        Dim height As Integer = Audio_VOXPictureBox2.Height - 1

        If VOX1 IsNot Nothing AndAlso VOX1.IsLatched = True AndAlso Audio_VOXPictureBox2.Enabled = True Then
            e.Graphics.FillRectangle(New SolidBrush(Color.Red), 0, 0, width, height)
        Else
            e.Graphics.FillRectangle(New SolidBrush(Color.Maroon), 0, 0, width, height)
        End If

        e.Graphics.DrawRectangle(New Pen(Color.Black), 0, 0, width, height)

    End Sub

    Private Sub Audio_ComboBox1_DropDownClosed(sender As Object, e As EventArgs) Handles Audio_ComboBox1.DropDownClosed

        WaveInID = Audio_ComboBox1.Text
        StartWaveIn()
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub Tone_ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Tone_ComboBox1.SelectedIndexChanged

        Dim cb As ComboBox = CType(sender, ComboBox)

        TestToneType = cb.Text
        Tone_ComboBox1.Text = TestToneType

    End Sub

    Private Sub Tone_RadioButton1_2_3_4__CheckedChanged(sender As Object, e As EventArgs) Handles Tone_RadioButton1.CheckedChanged, Tone_RadioButton2.CheckedChanged, Tone_RadioButton3.CheckedChanged, Tone_RadioButton4.CheckedChanged

        If Tone_RadioButton1.Checked = True Then
            CaptureTestTone = 0
        ElseIf Tone_RadioButton2.Checked = True Then
            CaptureTestTone = 1
        ElseIf Tone_RadioButton3.Checked = True Then
            CaptureTestTone = 2
        ElseIf Tone_RadioButton4.Checked = True Then
            CaptureTestTone = 3
        End If
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub Tone_TrackBarVerticle1_ValueChanged(sender As Object, e As System.EventArgs) Handles Tone_TrackBarVerticle1.ValueChanged

        TestToneLevel = Tone_TrackBarVerticle1.Value - 90
        Tone_TrackBarVerticle1.Value = TestToneLevel + 90
        Tone_Label3.Text = CStr(TestToneLevel) & " dB"
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub Tone_TrackBarVerticle2_ValueChanged(sender As Object, e As System.EventArgs) Handles Tone_TrackBarVerticle2.ValueChanged

        TestToneFrequency = (Tone_TrackBarVerticle2.Value + 1) * 50
        Tone_TrackBarVerticle2.Value = CInt(TestToneFrequency / 50) - 1
        Tone_Label5.Text = CStr(TestToneFrequency)
        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub TP2_Button1_Click(sender As Object, e As EventArgs) Handles TP2_Button1.Click

        Dim p As New Process()
        p.StartInfo.UseShellExecute = True
        p.StartInfo.FileName = ProgramPath & ProgramName & " Log.txt"
        p.Start()

    End Sub

    Private Sub TP3_Button2_Click(sender As Object, e As EventArgs) Handles TP2_Button2.Click

        If TP2_DataGridView1.Rows.Count = 0 Then Return

        Using New CenteredMessageBox(Me)
            If MessageBox.Show(Me, TP2_Button2.Text & vbCr & vbCr & "Are You Sure?", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then Return
        End Using

        TP2_DataGridView1.Rows.Clear()

    End Sub

    Public Function UpdateGUI(Code As Integer, Data As Object) As Object Handles WaveIn1.UpdateGUI, BCCalls1.UpdateGUI, VOX1.UpdateGUI

        If ProgramClosing = True Then Return Nothing

        Try
            If InvokeRequired = False Then
                Return UpdateGUI1(Code, Data)
            ElseIf Code < 100 Then
                Return Invoke(New UpdateUIDelegate(AddressOf UpdateGUI1), New Object() {Code, Data})
            Else '100 and above uses async invoke
                Return BeginInvoke(New UpdateUIDelegate(AddressOf UpdateGUI1), New Object() {Code, Data})
            End If
        Catch ex As Exception

        End Try

        Return Nothing

    End Function

    Private Function UpdateGUI1(Code As Integer, Obj As Object) As Object

        If ProgramClosing = True Then Return Nothing

        Select Case Code
            Case 1 'Audio meter
                Dim int1() As Integer = CType(Obj, Integer())
                If TabControl1.SelectedTab Is TabPage1 Then
                    Audio_MeterVert1.Value = int1(0)
                    Audio_Label10.Text = CStr(int1(1)) & " dB"
                End If
            Case 2
                Audio_VOXPictureBox1.Invalidate()
            Case 3
                Audio_VOXPictureBox2.Invalidate()
                If BCCalls1 IsNot Nothing Then
                    If VOX1.IsLatched = True Then
                        BCCalls1.StartAudio()
                    Else
                        BCCalls1.StopAudio()
                    End If
                End If
            Case 4
                Dim str1() As String = TryCast(Obj, String())
                If str1 IsNot Nothing Then WriteLogs(str1(1))
        End Select

        Return Nothing

    End Function

    Private Sub FormatStatusLogGrid()

        Dim Column1 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn
        Dim Column2 As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn

        Column1.Name = "Time"
        Column1.HeaderText = "Time"
        Column1.Width = 170

        Column2.Name = "Messages"
        Column2.HeaderText = "Messages"
        Column2.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        TP2_DataGridView1.Columns.Add(Column1)
        TP2_DataGridView1.Columns.Add(Column2)

    End Sub

    Private Sub SetupAbout()

        Dim str1 As String

        TP4_Label2.Text = "An application specifically designed for the Broadcastify Calls platform. Connect a receiver audio output to the computer audio input. The receiver must stay tuned to a specific conventional frequency And Not scanning."
        TP4_Label3.Text = "Version: " & Version

        str1 = "https://github.com/BobAProScan/VOX-Recorder"
        TP4_LinkLabel1.Links.Add(0, TP4_LinkLabel1.Text.Length + 1, str1)

        TP4_LinkLabel2.Links.Add(0, TP4_LinkLabel2.Text.Length, TP4_LinkLabel2.Text)

        str1 = "https://www.proscan.org"
        TP4_LinkLabel3.Links.Add(0, TP4_LinkLabel3.Text.Length + 1, str1)

        str1 = "https://wiki.radioreference.com/index.php/Broadcastify-Calls"
        TP4_LinkLabel4.Links.Add(0, TP4_LinkLabel4.Text.Length + 1, str1)

    End Sub

    Private Sub TP4_LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles TP4_LinkLabel2.LinkClicked

        Try
            Dim p As New Process()
            p.StartInfo.UseShellExecute = True
            p.StartInfo.FileName = "mailto:" & e.Link.LinkData.ToString() & "?subject=" & ProgramName
            p.Start()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub tp4_LinkLabel1_3_4_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles TP4_LinkLabel1.LinkClicked, TP4_LinkLabel3.LinkClicked, TP4_LinkLabel4.LinkClicked

        Try
            Dim p As New Process()
            p.StartInfo.UseShellExecute = True
            p.StartInfo.FileName = e.Link.LinkData.ToString()
            p.Start()
        Catch ex As Exception
        End Try

    End Sub
    Private Function CheckFields() As Boolean

        TP1_TextBox1.Text = TP1_TextBox1.Text.Trim
        TP1_TextBox2.Text = TP1_TextBox2.Text.Trim
        TP1_TextBox4.Text = TP1_TextBox4.Text.Trim
        TP1_TextBox5.Text = TP1_TextBox5.Text.Trim

        If TP1_TextBox1.Text = String.Empty Then
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, TP1_Label1.Text & " must Not be blank", ProgramName)
            End Using
            TP1_TextBox1.Focus()
            Return False
        End If
        If TP1_TextBox2.Text = String.Empty Then
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, TP1_Label2.Text & " must Not be blank", ProgramName)
            End Using
            TP1_TextBox2.Focus()
            Return False
        End If
        If TP1_TextBox4.Text = String.Empty Then
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, TP1_Label6.Text & " must Not be blank", ProgramName)
            End Using
            TP1_TextBox4.Focus()
            Return False
        End If
        If TP1_TextBox5.Text = String.Empty Then
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, TP1_Label7.Text & " must Not be blank", ProgramName)
            End Using
            TP1_TextBox5.Focus()
            Return False
        End If

        Return True

    End Function

    Public Sub SetupWaveInComboIn()

        Dim WaveInDevCaps() As String = GetWaveInDevCaps()

        Audio_ComboBox1.Items.Clear()
        Audio_ComboBox1.Items.AddRange(WaveInDevCaps)
        Audio_ComboBox1.Text = WaveInID
        If Audio_ComboBox1.SelectedIndex = -1 Then Audio_ComboBox1.SelectedIndex = 0


    End Sub

    Public Sub StartWaveIn()

        Static tempWaveInID As String
        If tempWaveInID = WaveInID Then
            Return
        End If
        tempWaveInID = WaveInID

        If WaveIn1 IsNot Nothing Then
            WaveIn1.Dispose()
            WaveIn1 = Nothing
        End If
        ClearAudioControls()

        If WaveInID <> "None" Then
            If WaveInID = String.Empty Then WaveInID = "Primary Sound Capture Device"
            Dim DeviceID As Integer = GetWaveInDeviceID(WaveInID)
            If DeviceID = 9999 Then
                WaveInID = "Primary Sound Capture Device"
                DeviceID = -1
            End If
            Dim WaveErrorDesc As String = String.Empty
            WaveIn1 = New WaveIn(DeviceID, 22050, 50, CInt(22050 * 0.2), WaveErrorDesc)
            Audio_RadioButton1_CheckedChanged(Nothing, Nothing)
            If WaveErrorDesc <> String.Empty Then
                WaveIn1.Dispose()
                WaveIn1 = Nothing
                ClearAudioControls()
                If ProgramFinishedLoading = True Then
                    Using New CenteredMessageBox(Me)
                        MessageBox.Show(Me, "Unable To open the Input Sound Device" & vbCr & vbCr &
                                    "The input Is Not enabled Or Not plugged in" & vbCr & vbCr &
                                    "Verify the input Is enabled And plugged in, in the Windows Mixer" & vbCr & vbCr &
                                    "Windows Message: " & WaveErrorDesc, ProgramName)
                    End Using
                End If
            End If
        End If

    End Sub

    Public Sub ClearAudioControls()

        Audio_MeterVert1.Value = -90
        Audio_Label10.Text = "-90 dB"

    End Sub

    Private Sub WriteLogs(Data As String)

        Dim time As String = Now.ToString

        WriteStatusGrid(time, Data, TP2_DataGridView1)
        WriteFileLog(time, Data, ProgramPath & ProgramName & " Log.txt", ProgramName & " Log - Last 1000 Entries")
        If Data.ToLower.Contains("error") Then TP1_Label5.Text = "Last Error: " & time & " - " & Data

    End Sub

    Private Sub WriteStatusGrid(Time As String, Message As String, DGV As DataGridView)

        Try
            Dim data() As String = Message.Split(CChar(vbCr))

            DGV.Rows.Add(New String() {Time, data(0)})
            If data.Length = 2 Then
                DGV.Rows.Add(New String() {String.Empty, data(1)})
            End If
            If DGV.Height > 0 Then
                If DGV.Rows.Count > 0 Then DGV.FirstDisplayedScrollingRowIndex = DGV.Rows.Count - 1
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub WriteFileLog(Time As String, Message As String, PathFile As String, Header As String)

        Try
            Dim lines() As String = Nothing
            Dim list As List(Of String)

            If File.Exists(PathFile) Then
                lines = File.ReadAllLines(PathFile)
            End If
            If lines Is Nothing Then
                list = New List(Of String)
            Else
                list = New List(Of String)(lines)
                list.Remove(Header)
                list.Remove(Header.Replace(" - Last 1000 Entries", String.Empty))
            End If
            Dim count As Integer
            For i As Integer = list.Count - 1 To 0 Step -1
                If count > 1000 - 2 Then '1 extra for the new row
                    list.RemoveAt(i)
                End If
                count += 1
            Next
            list.Insert(0, Header)
            list.Add(Time & " - " & Message.Replace(ProgramPath, String.Empty))
            File.WriteAllLines(PathFile, list.ToArray)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub TP1_RadioButton1_2_3_CheckedChanged(sender As Object, e As EventArgs) Handles TP1_RadioButton1.CheckedChanged, TP1_RadioButton2.CheckedChanged, TP1_RadioButton3.CheckedChanged

        If TP1_RadioButton1.Checked = True Then
            Loopback = 1
            If PlayMP3_NAudioStreaming1 IsNot Nothing Then
                PlayMP3_NAudioStreaming1.Dispose()
                PlayMP3_NAudioStreaming1 = Nothing
            End If
        ElseIf TP1_RadioButton2.Checked = True Then
            Loopback = 2
            PlayMP3_NAudioStreaming1 = New PlayMP3_NAudioStreaming
        ElseIf TP1_RadioButton3.Checked = True Then
            Loopback = 3
            PlayMP3_NAudioStreaming1 = New PlayMP3_NAudioStreaming
        End If
        TP1_CheckBox6.Enabled = Not TP1_RadioButton1.Checked

        If ProgramFinishedLoading = True Then SaveConfigTimer.Start()

    End Sub

    Private Sub TP1_CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles TP1_CheckBox6.CheckedChanged

        LoopbackAboveThreshold = TP1_CheckBox6.Checked

    End Sub

    Private Sub TP1_Button1_Click(sender As Object, e As EventArgs) Handles TP1_Button1.Click

        Try
            If IsNumeric(TP1_TextBox2.Text) Then
                Dim p As New Process()
                p.StartInfo.UseShellExecute = True
                p.StartInfo.FileName = "https://www.broadcastify.com/calls/manage/" & TP1_TextBox2.Text
                p.Start()
            Else
                Using New CenteredMessageBox(Me)
                    MessageBox.Show(Me, "System ID Not set" & vbCr & vbCr & "Set the System ID And try again", ProgramName)
                End Using
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub TP1_TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TP1_TextBox2.KeyPress

        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub

    Private Sub TP1_TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TP1_TextBox2.TextChanged

        Dim digitsOnly As New Regex("[^\d]")
        TP1_TextBox2.Text = digitsOnly.Replace(TP1_TextBox2.Text, String.Empty)

    End Sub

    Private Sub TP1_TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TP1_TextBox4.KeyPress

        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub

    Private Sub TP1_TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TP1_TextBox4.TextChanged

        Dim digitsOnly As New Regex("[^\d]")
        TP1_TextBox4.Text = digitsOnly.Replace(TP1_TextBox4.Text, String.Empty)

    End Sub

    Private Sub TP1_TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TP1_TextBox5.KeyPress

        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.Handled = True
        End If

    End Sub

    Private Sub TP1_TextBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TP1_TextBox5.TextChanged

        Dim digitsOnly As New Regex("[^\d.]")
        TP1_TextBox5.Text = digitsOnly.Replace(TP1_TextBox5.Text, String.Empty)

    End Sub

    Private Sub TP1_TextBox2_LostFocusd(ByVal sender As Object, ByVal e As System.EventArgs) Handles TP1_TextBox2.LostFocus

        TP1_Button1.Text = "Manage Calls System: Node " & TP1_TextBox2.Text & " Web page"

    End Sub

    Private Sub TP1_TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TP1_TextBox5.LostFocus

        Try
            If IsNumeric(TP1_TextBox5.Text) Then
                Dim dbl As Double = CDbl(TP1_TextBox5.Text)
                TP1_TextBox5.Text = dbl.ToString("0.00000")
            Else
                TP1_TextBox5.Text = String.Empty
            End If
        Catch ex As Exception
            TP1_TextBox5.Text = String.Empty
        End Try

    End Sub

    Private Sub TP4_Button1_Click(sender As Object, e As EventArgs) Handles TP4_Button1.Click

        Dim tempTopMost As Boolean = Me.TopMost
        Me.TopMost = False

        Dim FrmCheckNewestVersion1 As New FrmCheckNewestVersion
        FrmCheckNewestVersion1.ShowDialog(Me)
        FrmCheckNewestVersion1.Dispose()
        FrmCheckNewestVersion1 = Nothing

        Me.TopMost = tempTopMost

    End Sub

    Private Sub TP2_DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles TP2_DataGridView1.SelectionChanged
        'prevents 1st cell highlighting when program loads

        TP2_DataGridView1.ClearSelection()

    End Sub
End Class
