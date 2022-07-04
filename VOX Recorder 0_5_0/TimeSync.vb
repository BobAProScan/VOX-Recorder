Public Class TimeSync
    'https://www.epochconverter.com/

    Private Const TIMER_INTERVAL As Integer = 1000 * 60 * 10 '10 minutes

    Public ComputerTimeOffset As TimeSpan
    Private Timer As System.Timers.Timer
    Private WithEvents GetNTPBackGroundWorker As BackgroundWorker

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ComboBox1.Items.AddRange(New Object() {"time.nist.gov", "time.windows.com", "time.google.com", "ntp1.inrim.it", "ntp2.inrim.it", "oceania.pool.ntp.org", "pool.ntp.org", "north-america.pool.ntp.org", "europe.pool.ntp.org", "asia.pool.ntp.org"})

    End Sub

    Private Sub TimeSync_Load(sender As Object, e As EventArgs) Handles Me.Load

        Timer = New System.Timers.Timer(TIMER_INTERVAL)
        Timer.SynchronizingObject = Me 'required because invokerequired comes back true
        AddHandler Timer.Elapsed, AddressOf TimerEvent

        GetClock()

    End Sub

    Private Sub TimerEvent(source As Object, e As ElapsedEventArgs)

        GetClock()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        GetClock()

    End Sub

    Private Sub ComboBox1_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox1.DropDownClosed

        If ProgramFinishedLoading = True Then fMain.SaveConfigTimer.Start()

    End Sub

    Private Sub GetClock()

        If Timer Is Nothing Then Return 'this is just to make sure ComboBox1.SelectedIndexChanged doesn't execute before Me.Load executes
        If String.IsNullOrEmpty(ComboBox1.Text) = True Then Return

        GetNTPBackGroundWorker = New BackgroundWorker
        GetNTPBackGroundWorker.RunWorkerAsync(ComboBox1.Text)

    End Sub

    Private Sub GetNTPBackGroundWorker_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles GetNTPBackGroundWorker.DoWork

        Dim timeNTPserver As DateTime
        Dim servername As String = CStr(e.Argument)
        Dim stratum As String = String.Empty
        Dim roundtrip As Double

        timeNTPserver = GetNetworkTime(servername, stratum, roundtrip) 'in UTC format
        For I As Integer = 0 To 9
            If timeNTPserver > DateTime.MinValue Then Exit For
            Thread.Sleep(100)
            timeNTPserver = GetNetworkTime(servername, stratum, roundtrip) 'in UTC format
        Next

        e.Result = New Object() {timeNTPserver, stratum, roundtrip, Now}

    End Sub

    Private Sub GetNTPBackGroundWorker_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles GetNTPBackGroundWorker.RunWorkerCompleted

        Dim obj() As Object = CType(e.Result, Object())
        Dim timeNTPServer As DateTime = CDate(obj(0))
        If timeNTPServer > DateTime.MinValue Then
            ComputerTimeOffset = New TimeSpan(timeNTPServer.Ticks - CDate(obj(3)).ToUniversalTime.Ticks)
            If ComputerTimeOffset <> Nothing Then
                Label4.Text = CStr(obj(1)) 'stratum
                Label6.Text = String.Format("{0:N2} ms", CDbl(obj(2))) 'roundtrip
                Label8.Text = String.Format("{0:N3}", ComputerTimeOffset.TotalSeconds) & " Seconds" 'time offset
                Label10.Text = Now.Add(ComputerTimeOffset).ToString 'last sync
                Label12.Text = Now.Add(ComputerTimeOffset).AddMilliseconds(TIMER_INTERVAL).ToString 'next sync
                Dim actualUDTTimeDT As DateTime = Now.Add(ComputerTimeOffset).ToUniversalTime
                Label14.Text = ToUnixTime(actualUDTTimeDT).ToString
            End If
            ''FOR TESTING
            ''perfect
            'Dim ntpServertime As String = timeNTPServer.ToString("MM/dd/yy HH:mm:ss.fff")
            'Dim comptime As String = Now.ToString("MM/dd/yy HH:mm:ss.fff")
            'Dim ActualLocalTime As String = Now.Add(ComputerTimeOffset).ToString("MM/dd/yy HH:mm:ss.fff")
            'Dim UTCActualtime As String = Now.Add(ComputerTimeOffset).ToUniversalTime.ToString("MM/dd/yy HH:mm:ss.fff")
            ''https://stackoverflow.com/questions/2883576/how-do-you-convert-epoch-time-in-c
            ''veryfies Epoch - https://www.epochconverter.com/
            'Dim epochtime1 As Long = ToUnixTime(Now.Add(ComputerTimeOffset).ToUniversalTime)
            'Dim fromepochtime1 As String = FromUnixTime(epochtime1).ToString
            'Dim epochtime2 As Long = ToUnixTime(timeNTPServer)
            'Dim fromepochtime2 As String = FromUnixTime(epochtime2).ToString
            'd("           ")
            'd("NTP Server Time=      " & ntpServertime) 'UTC
            'd("Computer Time=        " & comptime)
            'd("Actual Time Local=    " & ActualLocalTime)
            'd("UTC Actual Time=      " & UTCActualtime)
            'd("Epoch Time1=          " & epochtime1)
            'd("From Epoch Time1=     " & fromepochtime2)
            'd("Epoch Time2=          " & epochtime2)
            'd("From Epoch Time2=     " & fromepochtime2)
            ''END TESTING
        Else
            d("Server unreachable: " & ComboBox1.Text & vbCrLf & vbCrLf & "Wait for next update or select another server")
            'Message(Me, "Server unreachable: " & ComboBox1.Text & vbCrLf & vbCrLf & "Try selecting another server")
        End If
        Timer.Stop()
        Timer.Start()

        GetNTPBackGroundWorker = Nothing

    End Sub

    Private Function GetNetworkTime(NtpServer As String, ByRef Statum As String, ByRef RoundTrip As Double) As DateTime

        Try
            Const DaysTo1900 As Integer = 1900 * 365 + 95
            Const TicksPerSecond As Long = 10000000L
            Const TicksPerDay As Long = 24 * 60 * 60 * TicksPerSecond
            Const TicksTo1900 As Long = DaysTo1900 * TicksPerDay
            Dim ntpData(47) As Byte
            ntpData(0) = &H1B
            Dim addresses As IPAddress() = Dns.GetHostEntry(NtpServer).AddressList
            Dim ipEndPoint As New IPEndPoint(addresses(0), 123)
            Dim pingDuration As Long = Stopwatch.GetTimestamp
            Using socket As New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
                socket.Connect(ipEndPoint)
                socket.ReceiveTimeout = 5000
                socket.Send(ntpData)
                pingDuration = Stopwatch.GetTimestamp
                socket.Receive(ntpData)
                pingDuration = Stopwatch.GetTimestamp - pingDuration
            End Using
            Dim pingTicks As Long = CLng(pingDuration * TicksPerSecond / Stopwatch.Frequency)
            RoundTrip = New TimeSpan(pingTicks).TotalMilliseconds
            Dim intPart As Long = CLng(ntpData(40)) << 24 Or CLng(ntpData(41)) << 16 Or CLng(ntpData(42)) << 8 Or CLng(ntpData(43))
            Dim fractPart As Long = CLng(ntpData(44)) << 24 Or CLng(ntpData(45)) << 16 Or CLng(ntpData(46)) << 8 Or CLng(ntpData(47))
            Dim netTicks As Long = intPart * TicksPerSecond + (fractPart * TicksPerSecond >> 32)
            Dim networkDateTime As New DateTime(CLng(TicksTo1900 + netTicks + pingTicks / 2))
            Statum = ntpData(1).ToString
            Return networkDateTime 'this is in UTC format already so don't convert
        Catch
        End Try

        Return DateTime.MinValue

    End Function

    Private Sub TimeSync_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        If Timer IsNot Nothing Then
            RemoveHandler Timer.Elapsed, AddressOf TimerEvent
            Timer.Dispose()
        End If

    End Sub

End Class
