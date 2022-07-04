Module [Module]

    Public Sub d(obj1 As Object)
        'Makes it easier to type d then Debug.WriteLine

#If DEBUG Then
        Debug.WriteLine(CStr(obj1).Replace(vbTab, ","))
#End If

    End Sub

    Public Sub ReadConfig()

        Try
            With fMain
                'defaults
                .TimeSync1.ComboBox1.Text = "time.nist.gov"
                .TP1_TextBox1.Text = String.Empty
                .TP1_TextBox2.Text = String.Empty
                .TP1_Label4.Text = String.Empty
                .TP1_TextBox4.Text = "1"
                .TP1_TextBox5.Text = String.Empty
                .TP1_CheckBox1.Checked = False
                .TP1_CheckBox2.Checked = False
                .TP1_CheckBox3.Checked = True
                .TP1_CheckBox4.Checked = False
                .TP1_CheckBox5.Checked = True
                .TP1_RadioButton1.Checked = True
                .TP1_CheckBox6.Checked = True

                WaveInID = "Primary Sound Capture Device"
                .Audio_VOXTrackBar1.Value = 30
                .Audio_RadioButton1.Checked = True
                TestToneType = "Sine"
                TestToneLevel = -30
                TestToneFrequency = 1000

                If File.Exists(Application.ExecutablePath.Replace("exe", "cfg")) = True Then
                    Dim lines() As String = File.ReadAllLines(Application.ExecutablePath.Replace("exe", "cfg"))
                    Dim str1 As String
                    Dim int1 As Integer
                    Dim temp() As String
                    For Each line As String In lines
                        int1 = InStr(line, "]=")
                        If int1 > 1 Then
                            str1 = Mid(line, int1 + 2)
                            Select Case Mid(line, 2, int1 - 2).Trim
                                Case "PROGRAM STATE"
                                    Select Case str1
                                        Case "Normal"
                                            .WindowState = FormWindowState.Normal
                                        Case "Minimized"
                                            .WindowState = FormWindowState.Minimized
                                        Case "Maximized"
                                            .WindowState = FormWindowState.Maximized
                                    End Select
                                Case "PROGRAM SIZE LOCATION"
                                    If str1 <> String.Empty AndAlso InStr(str1, ",") > 0 Then
                                        temp = str1.Split(CChar(","))
                                        If temp.Length = 4 AndAlso IsNumeric(temp(0)) AndAlso IsNumeric(temp(1)) AndAlso IsNumeric(temp(2)) AndAlso IsNumeric(temp(3)) Then
                                            .Bounds = New Rectangle(CInt(temp(0)), CInt(temp(1)), CInt(temp(2)), CInt(temp(3)))
                                        End If
                                    End If
                                Case "TAB INDEX"
                                    For Each tp As TabPage In .TabControl1.TabPages
                                        If tp.Name = str1 Then
                                            .TabControl1.SelectedTab = tp
                                        End If
                                    Next
                                Case "TIME SYNC SERVER"
                                    If str1 <> String.Empty Then .TimeSync1.ComboBox1.Text = str1
                                Case "API KEY"
                                    If str1 <> String.Empty Then .TP1_TextBox1.Text = str1
                                Case "SYSTEM ID"
                                    If str1 <> String.Empty Then .TP1_TextBox2.Text = str1
                                Case "SYSTEM NAME"
                                    If str1 <> String.Empty Then .TP1_Label4.Text = str1
                                Case "SLOT ID"
                                    If str1 <> String.Empty Then .TP1_TextBox4.Text = str1
                                Case "CONV FREQUENCY"
                                    If str1 <> String.Empty Then .TP1_TextBox5.Text = str1
                                Case "AUTO START"
                                    If str1 = "True" Then .TP1_CheckBox1.Checked = True Else .TP1_CheckBox1.Checked = False
                                Case "TOP MOST"
                                    If str1 = "True" Then .TP1_CheckBox2.Checked = True Else .TP1_CheckBox2.Checked = False
                                Case "LESS 2 SECONDS"
                                    If str1 = "True" Then .TP1_CheckBox3.Checked = True Else .TP1_CheckBox3.Checked = False
                                Case "PLAYBACK"
                                    If str1 = "True" Then .TP1_CheckBox4.Checked = True Else .TP1_CheckBox4.Checked = False
                                Case "SAVE RECORDINGS"
                                    If str1 = "True" Then .TP1_CheckBox5.Checked = True Else .TP1_CheckBox5.Checked = False
                                Case "LOOPBACK"
                                    Select Case str1
                                        Case "1"
                                            .TP1_RadioButton1.Checked = True
                                        Case "2"
                                            .TP1_RadioButton2.Checked = True
                                        Case "3"
                                            .TP1_RadioButton3.Checked = True
                                    End Select
                                Case "LOOPBACK ABOVE THRESHOLD"
                                    If str1 = "True" Then .TP1_CheckBox6.Checked = True Else .TP1_CheckBox6.Checked = False
                                Case "WAVEIN ID"
                                    If str1 <> String.Empty Then WaveInID = str1
                                Case "VOX THRESHOLD"
                                    If IsNumeric(str1) = True Then .Audio_VOXTrackBar1.Value = CInt(str1)
                                Case "CHANNEL SELECTED"
                                    Select Case str1
                                        Case "1"
                                            .Audio_RadioButton1.Checked = True
                                        Case "2"
                                            .Audio_RadioButton2.Checked = True
                                    End Select
                                Case "TEST TONE TYPE"
                                    If str1 <> String.Empty Then TestToneType = str1
                                Case "TEST TONE MODE"
                                    Select Case str1
                                        Case "1"
                                            .Tone_RadioButton1.Checked = True
                                        Case "2"
                                            .Tone_RadioButton2.Checked = True
                                        Case "3"
                                            .Tone_RadioButton3.Checked = True
                                        Case "4"
                                            .Tone_RadioButton4.Checked = True
                                    End Select
                                Case "TEST TONE LEVEL"
                                    If IsNumeric(str1) = True Then TestToneLevel = CInt(str1)
                                Case "TEST TONE FREQUENCY"
                                    If IsNumeric(str1) = True Then TestToneFrequency = CInt(str1)
                            End Select
                        End If
                    Next
                End If
                Application.DoEvents()
                If .TP1_CheckBox1.Checked = True Then .TP1_Button3_Click(Nothing, Nothing)
                .TP1_Button1.Text = "Manage Calls System: Node " & .TP1_TextBox2.Text & " Web page"
                .VOX1.ThreasholdLevel = .Audio_VOXTrackBar1.Value
                .Audio_Label9.Text = .Audio_VOXTrackBar1.Value.ToString
                .SetupWaveInComboIn()
                .Tone_ComboBox1.Text = TestToneType
                .Tone_TrackBarVerticle1.Value = TestToneLevel + 90
                .Tone_Label3.Text = CStr(TestToneLevel) & " dB"
                .Tone_TrackBarVerticle2.Value = CInt(TestToneFrequency / 50) - 1
                .Tone_Label5.Text = CStr(TestToneFrequency)
            End With
        Catch ex As Exception
            Using New CenteredMessageBox(fMain)
                MessageBox.Show(fMain, "Read Config Error: " & ex.Message, ProgramName)
            End Using
        End Try

    End Sub

    Public Sub WriteConfig()

        Static Running As Boolean = False
        If Running = True Then Return
        Running = True

        Dim sb As New StringBuilder

        Try
            With fMain
                sb.AppendLine(ProgramName & " " & "CONFIGURATION FILE")
                sb.AppendLine(String.Empty)
                sb.AppendLine("[PROGRAM STATE]=" & .WindowState.ToString)
                sb.AppendLine("[PROGRAM SIZE LOCATION]=" & .Bounds.X & "," & .Bounds.Y & "," & .Bounds.Width & "," & .Bounds.Height)
                sb.AppendLine("[TAB INDEX]=" & .TabControl1.SelectedTab.Name)
                sb.AppendLine("[TIME SYNC SERVER]=" & .TimeSync1.ComboBox1.Text)
                sb.AppendLine("[API KEY]=" & .TP1_TextBox1.Text)
                sb.AppendLine("[SYSTEM ID]=" & .TP1_TextBox2.Text)
                sb.AppendLine("[SYSTEM NAME]=" & .TP1_Label4.Text)
                sb.AppendLine("[SLOT ID]=" & .TP1_TextBox4.Text)
                sb.AppendLine("[CONV FREQUENCY]=" & .TP1_TextBox5.Text)
                sb.AppendLine("[AUTO START]=" & .TP1_CheckBox1.Checked)
                sb.AppendLine("[TOP MOST]=" & .TP1_CheckBox2.Checked)
                sb.AppendLine("[LESS 2 SECONDS]=" & .TP1_CheckBox3.Checked)
                sb.AppendLine("[PLAYBACK]=" & .TP1_CheckBox4.Checked)
                sb.AppendLine("[SAVE RECORDINGS]=" & .TP1_CheckBox5.Checked)
                If .TP1_RadioButton1.Checked = True Then
                    sb.AppendLine("[LOOPBACK]=" & 1)
                ElseIf .TP1_RadioButton2.Checked = True Then
                    sb.AppendLine("[LOOPBACK]=" & 2)
                ElseIf .TP1_RadioButton3.Checked = True Then
                    sb.AppendLine("[LOOPBACK]=" & 3)
                End If
                sb.AppendLine("[LOOPBACK ABOVE THRESHOLD]=" & .TP1_CheckBox6.Checked)
                sb.AppendLine("[WAVEIN ID]=" & WaveInID)
                sb.AppendLine("[VOX THRESHOLD]=" & .Audio_VOXTrackBar1.Value)
                If .Audio_RadioButton1.Checked = True Then
                    sb.AppendLine("[CHANNEL SELECTED]=" & 1)
                ElseIf .Audio_RadioButton2.Checked = True Then
                    sb.AppendLine("[CHANNEL SELECTED]=" & 2)
                End If
                sb.AppendLine("[TEST TONE TYPE]=" & TestToneType)
                If .Tone_RadioButton1.Checked = True Then
                    sb.AppendLine("[TEST TONE MODE]=" & 1)
                ElseIf .Tone_RadioButton2.Checked = True Then
                    sb.AppendLine("[TEST TONE MODE]=" & 2)
                ElseIf .Tone_RadioButton3.Checked = True Then
                    sb.AppendLine("[TEST TONE MODE]=" & 3)
                ElseIf .Tone_RadioButton4.Checked = True Then
                    sb.AppendLine("[TEST TONE MODE]=" & 4)
                End If
                sb.AppendLine("[TEST TONE LEVEL]=" & TestToneLevel)
                sb.AppendLine("[TEST TONE FREQUENCY]=" & TestToneFrequency)
            End With

            File.WriteAllText(Application.ExecutablePath.Replace("exe", "cfg"), sb.ToString)
        Catch ex As Exception
            Using New CenteredMessageBox(fMain)
                MessageBox.Show(fMain, "Write Config Error: " & ex.Message, ProgramName)
            End Using
        End Try

        Running = False

    End Sub

    Public Function ConvertBytesToShorts(Bytes() As Byte) As Short()

        Dim shorts(CInt(Math.Ceiling(Bytes.Length / 2)) - 1) As Short
        Buffer.BlockCopy(Bytes, 0, shorts, 0, Bytes.Length)

        Return shorts

    End Function

    Public Function ConvertShortsToBytes(Shorts() As Short) As Byte()

        Dim bytes((Shorts.Length * 2) - 1) As Byte
        Buffer.BlockCopy(Shorts, 0, bytes, 0, bytes.Length)

        Return bytes

    End Function

    Public Function ToUnixTime([Date] As DateTime) As Long

        Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

        Return Convert.ToInt64(([Date] - epoch).TotalSeconds)

    End Function

    Public Function FromUnixTime(unixTime As Long) As DateTime

        Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

        Return epoch.AddSeconds(unixTime)

    End Function

    Public Function CombineBytes(a() As Byte, b() As Byte) As Byte()

        Dim c(a.Length + b.Length - 1) As Byte

        System.Buffer.BlockCopy(a, 0, c, 0, a.Length)
        System.Buffer.BlockCopy(b, 0, c, a.Length, b.Length)

        Return c

    End Function

    Public Function LimitShort(Input As Double, Level As Double) As Short

        Dim temp As Double = Input * Level

        If temp < -32768 Then temp = -32768
        If temp > 32767 Then temp = 32767

        Try
            Return Convert.ToInt16(temp)
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Public Function LimitShort(Input As Double) As Short

        If Input < -32768 Then Input = -32768
        If Input > 32767 Then Input = 32767

        Try
            Return Convert.ToInt16(Input)
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Public Function GetWaveInDevCaps() As String()

        Dim i As Integer
        Dim rc As Integer
        Dim str1(-1) As String
        Dim Caps As WinAPI.WAVEINCAPS = Nothing
        Dim sz As Integer = Marshal.SizeOf(Caps)
        Dim format As Integer


        ReDim str1(WinAPI.waveInGetNumDevs + 1)
        str1(0) = "Primary Sound Capture Device"
        For i = 0 To WinAPI.waveInGetNumDevs - 1
            rc = WinAPI.waveInGetDevCaps(i, Caps, sz)
            str1(i + 1) = Caps.szPname.Trim
            format = Caps.dwFormats
        Next i
        str1(str1.Length - 1) = "None"

        Return str1

    End Function

    Public Function GetWaveInDeviceID(StringName As String) As Integer

        Dim i As Integer
        Dim rc As Integer
        Dim Caps As WinAPI.WAVEINCAPS = Nothing
        Dim sz As Integer = Marshal.SizeOf(Caps)

        If StringName = "Primary Sound Capture Device" Then Return -1
        For i = 0 To WinAPI.waveInGetNumDevs - 1
            rc = WinAPI.waveInGetDevCaps(i, Caps, sz)
            If StringName = Caps.szPname.Trim Then Return i
        Next i

        Return 9999

    End Function


End Module
