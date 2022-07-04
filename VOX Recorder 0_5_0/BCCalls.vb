Public Class BCCalls
    Implements IDisposable

    Private MemoryStream As MemoryStream
    Private Collecting As Boolean
    Private WorkerId As Integer = 1
    Private StartTimeTicks As String = String.Empty
    Private TimeOutStopWatch As Stopwatch
    Private WaitAudioRestore As Boolean
    Private ConvertPCMToMP31 As ConvertPCMToMP3
    Private KeepAlivePollTimer As System.Timers.Timer
    Private ApiKey As String
    Private SystemID As String
    Private SystemName As String
    Private SlotID As String
    Private ConvFreq As String

    Public Event UpdateGUI(Code As Integer, Data As Object)

    Public Sub Start()

        ApiKey = fMain.TP1_TextBox1.Text
        SystemID = fMain.TP1_TextBox2.Text
        SystemName = fMain.TP1_Label4.Text
        SlotID = fMain.TP1_TextBox4.Text
        ConvFreq = fMain.TP1_TextBox5.Text
        StartMP3Encoder()
        KeepAlivePollTimer = New System.Timers.Timer(1000 * 60 * 10) 'Poll interval 10 minutes
        AddHandler KeepAlivePollTimer.Elapsed, AddressOf KeepAlivePollTimerEvent
        KeepAlivePollTimer.Start()
        CallsRunning = True
        RaiseEvent UpdateGUI(4, New String() {"BC", "Feed started"}) 'Logging

    End Sub

    Private Sub KeepAlivePollTimerEvent(source As Object, e As ElapsedEventArgs)

        Dim errorStr As String = HTTP.Send(ApiKey, SystemID, String.Empty, String.Empty, String.Empty, String.Empty, True, Nothing)
        errorStr = errorStr.Replace("OK", String.Empty)
        If errorStr <> String.Empty Then
            RaiseEvent UpdateGUI(4, New String() {"BC", "Keep Alive poll failed - Error: " & errorStr}) 'Logging
        End If

    End Sub

    Public Sub StartAudio()
        d("START - " & Now.ToString)

        StartTimeTicks = Now.Ticks.ToString
        Collecting = False
        MemoryStream = New MemoryStream
        Collecting = True
        WaitAudioRestore = False
        TimeOutStopWatch = Stopwatch.StartNew

    End Sub

    Public Sub AudioPCM(PCMShorts() As Short, SampleRate As Integer)

        If Collecting = True Then
            Dim b() As Byte = ConvertShortsToBytes(PCMShorts)
            MemoryStream.Write(b, 0, b.Length)
            If TimeOutStopWatch.Elapsed.TotalSeconds > 120 Then ' 2 minutes
                Collecting = False
                MemoryStream.Dispose()
                RaiseEvent UpdateGUI(4, New String() {"BC", "Recording void because of continuous audio for 2 minutes"}) 'Logging
                WaitAudioRestore = True
            End If
        End If
    End Sub

    Public Sub StopAudio()

        Collecting = False
        If WaitAudioRestore = True Then Return
        If MemoryStream Is Nothing Then Return

        MemoryStream.Flush()
        MemoryStream.Close()
        Dim audio() As Byte = MemoryStream.ToArray
        Dim startTime As New DateTime(CLng(StartTimeTicks))
        Dim duration As Double = Math.Round(Now.Subtract(startTime).TotalSeconds, 1)
        Dim actualLocalTimeDT As DateTime = startTime.Add(ComputerTimeOffset)
        Dim actualUDTTimeDT As DateTime = actualLocalTimeDT.ToUniversalTime
        Dim epoch As String = ToUnixTime(actualUDTTimeDT).ToString

        ''FOR TESTING
        'd("STOP STOP STOP STOP STOP STOP STOP STOP STOP STOP STOP STOP STOP - " & Now.ToString)
        'd("Duration Before=" & duration.ToString)
        'd("Length Before=" & audio.Length)
        'd("SIZE PER SECOND=" & Math.Round(audio.Length / Now.Subtract(startTime).TotalSeconds, 0))
        'd("Local Time=" & startTime.ToString)
        'd("Actual Local Time1=" & actualLocalTimeDT.ToString)
        'd("Actual Local Time2=" & FromUnixTime(ToUnixTime(actualLocalTimeDT)).ToString)
        'd("Actual UTC Time=" & actualUDTTimeDT.ToString)
        'd("Epoch=" & epoch)
        ''END TESTING

        If audio.Length - (((22050 * 2) - 500) * CInt(VOXDELAY)) > 0 Then 'trims trailing X seconds which is always dead air. X = = VOX delay
            ReDim Preserve audio(audio.Length - (((22050 * 2) - 500) * CInt(VOXDELAY)))
            duration = Math.Round(duration - VOXDELAY, 1)
        End If

        ''FOR TESTING
        'd("Duration After=" & duration.ToString)
        'd("Length After=" & audio.Length)
        ''END TESTING

        If fMain.TP1_CheckBox3.Checked AndAlso duration <= 2 Then Return

        Try
            If ConvertPCMToMP31 IsNot Nothing Then
                ConvertPCMToMP31.mp3Encode(ConvertBytesToShorts(audio), New String() {ApiKey, SystemID, SlotID, ConvFreq, duration.ToString, epoch})
            End If
        Catch ex As Exception
            RaiseEvent UpdateGUI(4, New String() {"BC", "Error: " & ex.Message}) 'Logging
            StopMP3Encoder()
            StartMP3Encoder()
        End Try

    End Sub

    Public Sub AudioMP3(audio() As Byte, Args() As String)

        audio = InsertID3v2Tags(audio)

        Dim ApiKey As String = Args(0)
        Dim SystemID As String = Args(1)
        Dim SlotID As String = Args(2)
        Dim Freq As String = Args(3)
        Dim duration As String = Args(4)
        Dim epoch As String = Args(5)
        Dim path As String = ProgramPath & "Recordings\"

        BC_CallsUpload(ApiKey, SystemID, SlotID, Freq, duration, epoch, audio)
        Try
            If fMain.TP1_CheckBox5.Checked Then
                If Directory.Exists(path) = False Then Directory.CreateDirectory(path)
                File.WriteAllBytes(path & epoch & ".mp3", audio)
            End If
        Catch ex As Exception
            RaiseEvent UpdateGUI(4, New String() {"BC", "Recording fie write - Error: " & ex.Message}) 'Logging
        End Try

    End Sub

    Private Function InsertID3v2Tags(audio() As Byte) As Byte()

        Dim ms As New MemoryStream

        MS_Write(ms, Encoding.ASCII.GetBytes("ID3"))
        MS_Write(ms, {&H3, &H0}) 'two version bytes - ID3v2.3.0 
        MS_Write(ms, {&H0}) 'flags
        MS_Write(ms, {&H0, &H0, &H0, &H0}) 'place holder for size
        WriteMP3TextFrame(ms, "TIT2", SystemName) 'Title
        WriteMP3TextFrame(ms, "TPE1", "Conv. Frequency:" & ConvFreq) 'Artist
        WriteMP3TextFrame(ms, "TCON", "Receiver Audio") 'Genre
        WriteMP3TextFrame(ms, "TPUB", "ProScan") 'Publisher
        WriteMP3TextFrame(ms, "TCOM", ProgramName & " " & Version) 'Composer
        WriteMP3TextFrame(ms, "TYER", Now.Add(ComputerTimeOffset).Year.ToString) 'Year
        WriteMP3TextFrame(ms, "COMM", "Date:" & Now.Add(ComputerTimeOffset).ToString("yyyy-MM-dd HH:mm:ss.fff") & ";System:" & SystemName & ";Frequency:" & ConvFreq & ";") 'Comment
        Dim bytes(3) As Byte
        Dim int1 As Integer = CInt(ms.Length - 10) 'frame size excluding frame header (frame size - 10)
        For i As Integer = 3 To 0 Step -1
            bytes(i) = CByte(int1 Mod 128)
            int1 \= 128
        Next
        ms.Seek(6, SeekOrigin.Begin)
        MS_Write(ms, bytes) 'Updates size
        ms.Seek(0, SeekOrigin.End)
        MS_Write(ms, audio)

        Return ms.ToArray

    End Function

    Private Sub WriteMP3TextFrame(ms As MemoryStream, TextFrameType As String, Text As String)

        Text = CleanASCII_1(Text.Trim)
        Text = Mid(Text, 1, 253)

        'frame type
        MS_Write(ms, Encoding.ASCII.GetBytes(TextFrameType))

        'size
        Dim b() As Byte
        If TextFrameType = "COMM" Then
            b = BitConverter.GetBytes(Text.Length + 6) 'same as 4 + 4 + 1 + 2 + 4 + Text.Length + 1 - 10 -- The size is calculated as frame size excluding frame header (frame size - 10).
        Else
            b = BitConverter.GetBytes(Text.Length + 2) 'same as 4 + 4 + 1 + 2 + Text.Length + 1 - 10 -- The size is calculated as frame size excluding frame header (frame size - 10).
        End If
        Array.Reverse(b)
        MS_Write(ms, b)

        'text encoding
        MS_Write(ms, {0})

        'flags
        MS_Write(ms, {&H0, &H0})

        If TextFrameType = "COMM" Then
            'language - eng, null terminator for short description
            MS_Write(ms, {&H65, &H6E, &H67, &H0})
        End If

        'text
        MS_Write(ms, Encoding.ASCII.GetBytes(Text))

        'terminate char
        MS_Write(ms, {0})

    End Sub

    Private Sub MS_Write(MS As MemoryStream, Bytes As Byte())

        MS.Write(Bytes, 0, Bytes.Length)

    End Sub
    Private Function CleanASCII_1(Str As String) As String

        Dim sb As New StringBuilder(Str.Length)

        For Each c As Char In Str
            If Asc(c) > 31 AndAlso Asc(c) < 127 Then
                sb.Append(c)
            End If
        Next

        Return sb.ToString

    End Function

    Private Sub BC_CallsUpload(ApiKey As String, SystemId As String, SlotID As String, Freq As String, Duration As String, Epoch As String, mp3Audio() As Byte)

        QueueAdd(ApiKey, SystemId, SlotID, Freq, Duration, Epoch, mp3Audio)

    End Sub

    Private Sub QueueAdd(ApiKey As String, SystemId As String, SlotID As String, Freq As String, Duration As String, Epoch As String, mp3Audio As Byte())

        Static WorkerQueue As Queue(Of QueueItem(Of Integer)) = New Queue(Of QueueItem(Of Integer))()
        MyQueuedBackgroundWorker.QueueWorkItem(WorkerQueue, Math.Min(System.Threading.Interlocked.Increment(WorkerId), WorkerId - 1),
        Function(args)
            Dim threadMessage As String = String.Format("Thread started at '{0}', Task Number={1}", DateTime.Now.ToString("HH:mm:ss.fff"), args.Argument)
            Dim errorStr As String = HTTP.Send(ApiKey, SystemId, SlotID, Freq, Duration, Epoch, False, mp3Audio)
            Return New With {Key .WorkerId = args.Argument, Key .Message = threadMessage, Key .ErrorStr = errorStr, Key .Audio = mp3Audio}
        End Function, Sub(args)
                          Dim completeMessage As String = String.Format("COMPLETED at '{0}' for Task Number={1}, Message={2}, " & vbCrLf & "ERROR={3}", DateTime.Now.ToString("HH:mm:ss.fff"), args.Result.WorkerId, args.Result.Message, args.Result.ErrorStr)
                          d(completeMessage)
                          If String.IsNullOrEmpty(args.Result.ErrorStr) = False Then
                              RaiseEvent UpdateGUI(4, New String() {"BC", "Failed upload - Error: " & args.Result.ErrorStr}) 'Logging
                          End If
                          If fMain.TP1_CheckBox4.Checked Then
                              PlayMP3_NAudio.Play(mp3Audio)
                          End If
                      End Sub)

    End Sub

    Private Sub StartMP3Encoder()

        If ProgramClosing = True Then Return

        Try
            If ConvertPCMToMP31 Is Nothing Then ConvertPCMToMP31 = New ConvertPCMToMP3("Mono", 22050, 32, 22050)
        Catch ex As Exception
            StopMP3Encoder()
            RaiseEvent UpdateGUI(4, New String() {"BC", "MP3 encoder failed to initialize - Error: " & ex.Message}) 'Logging
            Using New CenteredMessageBox(fMain)
                MessageBox.Show(fMain, "MP3 encoder failed to initialize - Error: " & ex.Message, ProgramName)
            End Using
        End Try

    End Sub

    Private Sub StopMP3Encoder()

        If ConvertPCMToMP31 IsNot Nothing Then
            ConvertPCMToMP31.Dispose()
            ConvertPCMToMP31 = Nothing
        End If

    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                'free managed resources when explicitly called
                CallsRunning = False
                StopMP3Encoder()
                If MemoryStream IsNot Nothing Then
                    MemoryStream.Dispose()
                End If
                KeepAlivePollTimer.Close()
                RemoveHandler KeepAlivePollTimer.Elapsed, AddressOf KeepAlivePollTimerEvent
                KeepAlivePollTimer.Dispose()
                KeepAlivePollTimer = Nothing
                RaiseEvent UpdateGUI(4, New String() {"BC", "Feed stopped"}) 'Logging
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