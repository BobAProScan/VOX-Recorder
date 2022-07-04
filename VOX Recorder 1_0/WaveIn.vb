Public Class WaveIn
    Implements IDisposable

    Private Const BUFFERSECONDSSIZE As Integer = 1200 '20 seconds

    Private _DeviceID As Integer
    Private _WaveInRunning As Boolean
    Private _WaveFormat As WinAPI.WAVEFORMATEX
    Private _WaveHeader(0) As WinAPI.WAVEHDR
    Private _HeaderHandle As GCHandle
    Private _WavePtr As IntPtr
    Private _Result As Integer
    Private _CountBuffer As Integer
    Private _lEventHandle As IntPtr
    Private _SamplesPerSec As Integer
    Private _Buffers As Integer
    Private _BufferSize As Integer
    Private _BufferStereo(-1) As Short
    Private _BufferMono(-1) As Short
    Private _Channel As Integer '1 = Left, 2 = Right

    Public Event UpdateGUI(Code As Integer, Data As Object)

    Public Sub New(DeviceID As Integer, SamplesPerSec As Integer, Buffers As Integer, BufferSize As Integer, ByRef WaveErrorDesc As String)

        _DeviceID = DeviceID
        _SamplesPerSec = SamplesPerSec
        _Buffers = Buffers
        _BufferSize = BufferSize
        If _BufferSize Mod 2 = 1 Then _BufferSize -= 1 'if odd then make even
        ReDim _BufferStereo(_BufferSize - 1)
        ReDim _BufferMono(CInt(_BufferSize / 2) - 1)
        With _WaveFormat
            .wFormatTag = WinAPI.WAVE_FORMAT_PCM
            .nChannels = 2
            .nSamplesPerSec = _SamplesPerSec
            .wBitsPerSample = 16
            .nBlockAlign = CShort(.nChannels * (.wBitsPerSample / 8))
            .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
        End With

        Try
            _lEventHandle = WinAPI.CreateEvent(IntPtr.Zero, False, False, Nothing)
        Catch ex As Exception
        End Try
        WinAPI.ResetEvent(_lEventHandle)

        For i As Integer = 1 To 5 'works well - if an error open then try up to 5 times
            _Result = WinAPI.waveInOpen(_WavePtr, _DeviceID, _WaveFormat, _lEventHandle, 0, WinAPI.CALLBACK_EVENT)
            If _Result = WinAPI.MMRESULT.MMSYSERR_NOERROR Then Exit For
            If i > 4 Then
                WaveErrorDesc = WinAPI.GetErrorDesc(_Result)
                Return
            End If
        Next

        ReDim _WaveHeader(_Buffers - 1)
        _HeaderHandle = GCHandle.Alloc(_WaveHeader, GCHandleType.Pinned)
        For _CountBuffer = 0 To _Buffers - 1
            With _WaveHeader(_CountBuffer)
                .lpData = Marshal.AllocHGlobal(CInt(_WaveFormat.nAvgBytesPerSec * 2))
                .dwUser = New IntPtr(_CountBuffer)
                .dwBufferLength = _BufferSize * 2
                .dwBytesRecorded = 0
                .dwFlags = 0
                .dwLoops = 0
                .lpNext = IntPtr.Zero
                .reserved = 0
            End With
            WinAPI.waveInPrepareHeader(_WavePtr, _WaveHeader(_CountBuffer), Marshal.SizeOf(_WaveHeader(_CountBuffer)))
            WinAPI.waveInAddBuffer(_WavePtr, _WaveHeader(_CountBuffer), Marshal.SizeOf(_WaveHeader(_CountBuffer)))
        Next

        _Result = WinAPI.waveInStart(_WavePtr)
        If _Result <> WinAPI.MMRESULT.MMSYSERR_NOERROR Then
            WaveErrorDesc = WinAPI.GetErrorDesc(_Result)
            Return
        End If
        _WaveInRunning = True

        Dim t As New Thread(AddressOf WaveIn)
        t.Name = "Audio"
        t.Start()

    End Sub

    Private Sub WaveIn()

        While True
            Try
                If _WaveInRunning = False Then Return
                WinAPI.WaitForSingleObject(_lEventHandle, 1000)
                WinAPI.ResetEvent(_lEventHandle)
                For _CountBuffer = 0 To _Buffers - 1
                    If _WaveInRunning = False Then Return
                    While (_WaveHeader(_CountBuffer).dwFlags And WinAPI.WHDR_DONE) <> WinAPI.WHDR_DONE
                        Thread.Sleep(1)
                        If _WaveInRunning = False Then Return
                    End While
                    If _WaveInRunning = False Then Return
                    If ProgramClosing = True Then Return
                    WaveInPointer(_WaveHeader(_CountBuffer).lpData)
                    If _WaveInRunning = False Then Return
                    _Result = WinAPI.waveInAddBuffer(_WavePtr, _WaveHeader(_CountBuffer), Marshal.SizeOf(_WaveHeader(_CountBuffer)))
                    If _Result <> WinAPI.MMRESULT.MMSYSERR_NOERROR Then
                        'd("Error: Capture - " & _Result)
                    End If
                Next
            Catch ex As Exception
                'dont do anything here, normal if exiting
            End Try
        End While

    End Sub

    Public Sub WaveInPointer(Ptr As IntPtr)

        If ProgramFinishedLoading = False Then Return
        If _WaveInRunning = False Then Return

        Marshal.Copy(Ptr, _BufferStereo, 0, _BufferStereo.Length)
        ProcessAudio1()

    End Sub

    Private Sub ProcessAudio1()

        If CaptureTestTone > 0 Then
            DSP.ToneGenerator(TestToneType, _BufferStereo, TestToneFrequency, TestToneLevel, _SamplesPerSec, False, CaptureTestTone)
        End If

        ProcessAudio2(_BufferStereo, _BufferMono)

        If CallsRunning = True AndAlso fMain.BCCalls1 IsNot Nothing Then
            fMain.BCCalls1.AudioPCM(_BufferMono, _SamplesPerSec)
            If Loopback = 3 Then
                If (LoopbackAboveThreshold = True AndAlso fMain.VOX1.IsLatched) OrElse (LoopbackAboveThreshold = False) Then
                    If PlayMP3_NAudioStreaming1 IsNot Nothing Then PlayMP3_NAudioStreaming1.Play(_BufferMono)
                End If
            End If
        End If
        If Loopback = 2 Then
            If (LoopbackAboveThreshold = True AndAlso fMain.VOX1.IsLatched) OrElse (LoopbackAboveThreshold = False) Then
                If PlayMP3_NAudioStreaming1 IsNot Nothing Then PlayMP3_NAudioStreaming1.Play(_BufferMono)
            End If
        End If

    End Sub

    Private Sub ProcessAudio2(BufferStereo() As Short, BufferMono() As Short)

        Static intervalSw As New Stopwatch
        Static intervalCount As Long
        Static intervalTotal As Double
        Static sw As New Stopwatch
        Static leftList As New List(Of Integer)
        Static rightList As New List(Of Integer)

        Dim timesPerSecond As Integer
        Dim i As Integer
        Dim int1 As Integer
        Dim int2 As Integer
        Dim short1 As Short
        Dim total1 As Double
        Dim count1 As Integer
        Dim total2 As Double
        Dim count2 As Integer
        Dim temp As Double
        Dim amplitude As Double = 10 ^ (6 / 20)
        Dim leftlevel As Double
        Dim leftValue As Integer
        Dim leftMeterValue As Integer
        Dim leftDBvalue As Integer = -90
        Dim leftTotal As Double
        Dim leftCount As Integer
        Dim rightIndex As Integer
        Dim rightlevel As Double
        Dim rightValue As Integer
        Dim rightMeterValue As Integer
        Dim rightDBvalue As Integer = -90
        Dim rightTotal As Double
        Dim rightCount As Integer

        intervalTotal += intervalSw.ElapsedMilliseconds
        intervalSw.Reset()
        intervalSw.Start()
        If intervalTotal < 1 Then Return
        intervalCount += 1
        timesPerSecond = CInt(1000 / (intervalTotal / intervalCount))

        If _Channel = 1 Then
            leftlevel = 1
        ElseIf _Channel = 2 Then
            rightlevel = 1
        End If

        For i = 0 To BufferStereo.Length - 1
            If i Mod 2 = 0 Then 'even (left)
                short1 = CShort(BufferStereo(i) * leftlevel)
                BufferStereo(i) = short1
                leftTotal += Math.Pow(short1, 2)
                leftCount += 1
            Else 'odd (right)
                short1 = CShort(BufferStereo(i) * rightlevel)
                BufferStereo(i) = short1
                rightTotal += Math.Pow(short1, 2)
                rightCount += 1
                temp = LimitShort(((BufferStereo(i) / 2) + (BufferStereo(i - 1) / 2)), amplitude)
                BufferMono(rightIndex) = CShort(temp)
                rightIndex += 1
            End If
        Next

        'Left
        total1 = 0
        count1 = 0
        total2 = 0
        count2 = 0
        If leftCount > 0 Then leftValue = CInt(Math.Sqrt(leftTotal / leftCount * 2))
        leftList.Add(leftValue)
        If leftList.Count > ((timesPerSecond * BUFFERSECONDSSIZE)) Then
            leftList.RemoveAt(0)
        End If
        int1 = CInt(leftList.Count - (timesPerSecond * 0.15)) - 2 '.15 second
        int2 = CInt(leftList.Count - (timesPerSecond * 0.15)) - 2 '.15 second
        For i = 0 To leftList.Count - 1
            'meter
            If i > int1 Then
                total1 += leftList(i)
                count1 += 1
            End If
            'indicator
            If i > int2 Then
                total2 += leftList(i)
                count2 += 1
            End If
        Next
        'meter
        If count1 > 0 Then
            leftMeterValue = CInt(total1 / count1)
            If leftMeterValue > 0 Then leftMeterValue = CInt(20 * Math.Log10(32767 / leftMeterValue) * -1) Else leftMeterValue = -90
            If leftMeterValue > 0 Then leftMeterValue = 0
            leftMeterValue += 90 'scales -90 - 0 to 0 - 90
        End If
        'indicator
        If count2 > 0 Then
            leftDBvalue = CInt(total2 / count2)
            If leftDBvalue > 0 Then leftDBvalue = CInt(20 * Math.Log10(32767 / leftDBvalue) * -1) Else leftDBvalue = -90
            If leftDBvalue < -90 Then leftDBvalue = -90
            If leftDBvalue > 0 Then leftDBvalue = 0
        End If

        'Right
        total1 = 0
        count1 = 0
        total2 = 0
        count2 = 0
        If rightCount > 0 Then rightValue = CInt(Math.Sqrt(rightTotal / rightCount * 2))
        rightList.Add(rightValue)
        If rightList.Count > ((timesPerSecond * BUFFERSECONDSSIZE)) Then
            rightList.RemoveAt(0)
        End If
        int1 = CInt(rightList.Count - (timesPerSecond * 0.15)) - 2 '.15 second 
        int2 = CInt(rightList.Count - (timesPerSecond * 0.15)) - 2 '.15 second 
        For i = 0 To rightList.Count - 1
            'meter
            If i > int1 Then
                total1 += rightList(i)
                count1 += 1
            End If
            'indicator
            If i > int2 Then
                total2 += rightList(i)
                count2 += 1
            End If
        Next
        'meter
        If count1 > 0 Then
            rightMeterValue = CInt(total1 / count1)
            If rightMeterValue > 0 Then rightMeterValue = CInt(20 * Math.Log10(32767 / rightMeterValue) * -1) Else rightMeterValue = -90
            If rightMeterValue > 0 Then rightMeterValue = 0
            rightMeterValue += 90 'scales -90 - 0 to 0 - 90
        End If
        'indicator
        If count2 > 0 Then
            rightDBvalue = CInt(total2 / count2)
            If rightDBvalue > 0 Then rightDBvalue = CInt(20 * Math.Log10(32767 / rightDBvalue) * -1) Else rightDBvalue = -90
            If rightDBvalue < -90 Then rightDBvalue = -90
            If rightDBvalue > 0 Then rightDBvalue = 0
        End If

        If sw.IsRunning = False Then
            sw.Start()
        End If
        If sw.ElapsedMilliseconds > 50 Then
            If _Channel = 1 Then
                Static templeftmetervalue As Integer
                Static templeftDBvalue1 As Integer
                If templeftmetervalue <> leftMeterValue OrElse templeftDBvalue1 <> leftDBvalue Then
                    RaiseEvent UpdateGUI(1, {leftMeterValue, leftDBvalue}) 'Audio meter
                End If
                templeftmetervalue = leftMeterValue
                templeftDBvalue1 = leftDBvalue
            ElseIf _Channel = 2 Then
                Static temprightmetervalue As Integer
                Static temprightDBvalue1 As Integer
                If temprightmetervalue <> rightMeterValue OrElse temprightDBvalue1 <> rightDBvalue Then
                    RaiseEvent UpdateGUI(1, {rightMeterValue, rightDBvalue}) 'Audio meter
                End If
                temprightmetervalue = rightMeterValue
                temprightDBvalue1 = rightDBvalue
            End If
            sw.Reset()
            sw.Start()
        End If
        If fMain.VOX1 IsNot Nothing Then
            fMain.VOX1.InputAudio(Math.Max(leftMeterValue, rightMeterValue))
        End If

    End Sub

    Public WriteOnly Property Channel() As Integer

        Set(Value As Integer)
            _Channel = Value
        End Set

    End Property

    Private disposedValue As Boolean = False 'detects redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)

        If Not Me.disposedValue Then
            On Error Resume Next
            _WaveInRunning = False
            Application.DoEvents() 'flushes the level meters before zeroing
            If _WavePtr <> IntPtr.Zero Then
                WinAPI.waveInReset(_WavePtr)
                If _HeaderHandle.IsAllocated Then _HeaderHandle.Free()
                For _CountBuffer = 0 To _Buffers - 1
                    WinAPI.waveInUnprepareHeader(_WavePtr, _WaveHeader(_CountBuffer), Marshal.SizeOf(_WaveHeader(_CountBuffer)))
                    Marshal.FreeHGlobal(_WaveHeader(_CountBuffer).lpData)
                Next
                WinAPI.waveInClose(_WavePtr)
                _WavePtr = IntPtr.Zero
                _WavePtr = Nothing
                _WaveHeader = Nothing
                _lEventHandle = Nothing
            End If
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
