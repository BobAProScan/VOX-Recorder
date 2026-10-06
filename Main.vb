
Module Main
    'https://iconsplace.com/blue-icons/radio-tower-icon-2/

    Public Const VOXDELAY As Double = 2.0

    Public fMain As FrmMain
    Public ProgramName As String = My.Application.Info.ProductName
    Public Version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor
    Public ProgramPath As String = Application.StartupPath & "\"
    Public ProgramFinishedLoading As Boolean
    Public ProgramClosing As Boolean
    Public ComputerTimeOffset As TimeSpan
    Public CallsRunning As Boolean
    Public WaveInID As String
    Public TestToneType As String
    Public CaptureTestTone As Integer
    Public TestToneLevel As Integer
    Public TestToneFrequency As Integer
    Public PlayMP3_NAudioStreaming1 As PlayMP3_NAudioStreaming
    Public Loopback As Integer
    Public LoopbackAboveThreshold As Boolean

    Public Sub Main()

        If Environment.OSVersion.Version.Major >= 6 Then
            WinAPI.SetProcessDPIAware
        End If
        Application.EnableVisualStyles()
        Application.CurrentCulture = CultureInfo.InvariantCulture
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture
        Application.Run(New FrmMain)

        fMain.Dispose()
        fMain = Nothing

    End Sub

End Module

