Imports NAudio

Public Class PlayMP3_NAudio

    Public Shared Sub Play(MP3Bytes As Byte())

        Dim waveOut As New Wave.WaveOut()
        waveOut.Init(New Wave.Mp3FileReader(New MemoryStream(MP3Bytes)))
        'd(waveOut.DeviceNumber)
        waveOut.Play()
        While True
            If waveOut.PlaybackState <> 1 Then Exit While
            Application.DoEvents()
        End While
        waveOut.Stop()
        waveOut.Dispose()

    End Sub

End Class

Public Class PlayMP3_NAudioStreaming
    Implements IDisposable

    Private bufferedWaveProvider As Wave.BufferedWaveProvider
    Private waveOut As Wave.WaveOut
    Private disposedValue As Boolean

    Public Sub New()

        bufferedWaveProvider = New Wave.BufferedWaveProvider(New Wave.WaveFormat(22050, 16, 1))
        bufferedWaveProvider.DiscardOnBufferOverflow = True
        bufferedWaveProvider.BufferLength = 22050 '= 1/2 second - 4410 bytes / sec
        waveOut = New Wave.WaveOut
        'd(waveOut.DeviceNumber) '0
        waveOut.Init(bufferedWaveProvider)

    End Sub

    Public Sub Play(PCMShorts As Short())

        Dim buffer() As Byte = ConvertShortsToBytes(PCMShorts)
        Play(buffer)

    End Sub
    Public Sub Play(PCMBytes As Byte())

        If bufferedWaveProvider Is Nothing Then Return
        If waveOut Is Nothing Then Return

        bufferedWaveProvider.AddSamples(PCMBytes, 0, PCMBytes.Length)

        Static count As Integer
        If count > 1000 Then count = 3
        If count > 2 Then
            waveOut.Play()
        End If
        count += 1

    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If bufferedWaveProvider IsNot Nothing Then
                    bufferedWaveProvider.ClearBuffer()
                    bufferedWaveProvider = Nothing
                End If
                If waveOut IsNot Nothing Then
                    waveOut.Stop()
                    waveOut.Dispose()
                End If
            End If
            disposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Public Class PlayMP3_MCI_A

    Public Shared Sub Play(File As String)

        Dim _RetVal As Integer
        WinAPI.mciSendString("close temp_alias", vbNullString, 0, IntPtr.Zero)
        _RetVal = WinAPI.mciSendString("open " & Chr(34) & File & Chr(34) & " Type mpegvideo alias temp_alias", vbNullString, 0, IntPtr.Zero)
        If _RetVal = 0 Then
            WinAPI.mciSendString("play temp_alias", vbNullString, 0, IntPtr.Zero)
        End If

    End Sub

End Class

Public Class PlayMP3_MCI_B

    Public Shared Sub Play(MP3Bytes As Byte())

        Dim path As String = "temp.mp3"
        Dim _RetVal As Integer
        File.WriteAllBytes(path, MP3Bytes)
        WinAPI.mciSendString("close temp_alias", vbNullString, 0, IntPtr.Zero)
        _RetVal = WinAPI.mciSendString("open " & Chr(34) & path & Chr(34) & " Type mpegvideo alias temp_alias", vbNullString, 0, IntPtr.Zero)
        If _RetVal = 0 Then
            WinAPI.mciSendString("play temp_alias wait", vbNullString, 0, IntPtr.Zero)
        End If
        File.Delete(path)

    End Sub

End Class

