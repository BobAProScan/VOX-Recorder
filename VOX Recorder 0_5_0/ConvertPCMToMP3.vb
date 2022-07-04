Public Class ConvertPCMToMP3
    Implements IDisposable

    Private Const BE_MP3_MODE_STEREO As Integer = 0
    Private Const BE_MP3_MODE_MONO As Integer = 3
    Private Const LQP_NOPRESET As Integer = -1 'quality 3
    Private Const LQP_NORMAL_QUALITY As Integer = 0 'quality 3
    Private Const LQP_LOW_QUALITY As Integer = 1 'quality 9
    Private Const LQP_HIGH_QUALITY As Integer = 2 'quality 2
    Private Const LQP_VERYHIGH_QUALITY As Integer = 5 'quality 0
    Private Const VBR_METHOD_NONE As Integer = -1
    Private Const VBR_METHOD_DEFAULT As Integer = 0
    Private Const MPEG2 As Integer = 0
    Private Const MPEG1 As Integer = 1

    Private _Be_Config_Format As WinAPI.BE_CONFIG_FORMAT_LHV1
    Private _HandleInput As GCHandle
    Private _PtrInput As IntPtr
    Private _OutputBuffer() As Byte
    Private _HandleStream As Integer
    Private _dwSamples As Integer
    Private _dwBufferSize As Integer
    Private _BytesOutput As Integer
    Private _Result As Integer

    Public Sub New(Mode As String, SampleRate As Integer, mp3BitRate As Integer, ResampleRate As Integer)

        If SampleRate = 0 Then
            Throw New Exception("SampleRate = 0 not supported")
            Return
        ElseIf mp3BitRate < 8 Then
            Throw New Exception("mp3BitRate < 8 not supported")
            Return
        End If

        With _Be_Config_Format
            .dwConfig = 256
            .dwStructVersion = 1
            .dwStructSize = Marshal.SizeOf(_Be_Config_Format)
            Select Case Mode
                Case "Stereo"
                    .nMode = BE_MP3_MODE_STEREO
                Case Else
                    .nMode = BE_MP3_MODE_MONO
            End Select
            .dwSampleRate = SampleRate
            .nPreset = LQP_NORMAL_QUALITY
            .dwBitrate = mp3BitRate
            .dwMaxBitrate = 0
            .dwPsyModel = 0
            .dwEmphasis = 0
            .bPrivate = 0
            .bCRC = 0
            .bCopyright = 0
            .bOriginal = 1
            .bWriteVBRHeader = 0
            .bEnableVBR = 0
            .nVBRQuality = 0
            .dwVbrAbr_bps = 0
            .nVbrMethod = VBR_METHOD_NONE
            .bNoRes = 0
            .bStrictIso = 0
            .nQuality = 0
            If SampleRate < 32000 Then
                .dwMpegVersion = MPEG2
            Else
                .dwMpegVersion = MPEG1
            End If
            If ResampleRate = 0 Then
                .dwReSampleRate = SampleRate 'prevents resample
            Else
                .dwReSampleRate = ResampleRate
            End If
        End With

        _Result = -1
        _Result = WinAPI.beInitStream(_Be_Config_Format, _dwSamples, _dwBufferSize, _HandleStream)
        If _Result <> 0 Then
            Throw New Exception("beInitStream failed to initialize")
            Return
        End If

        _HandleInput = GCHandle.Alloc(New Byte((_dwBufferSize) - 1) {}, GCHandleType.Pinned)
        _PtrInput = _HandleInput.AddrOfPinnedObject
        ReDim _OutputBuffer(_dwBufferSize - 1)

    End Sub

    Public Sub mp3Encode(BufferShort() As Short, Optional Str() As String = Nothing)

        Try
            Dim mp3bytes(-1) As Byte
            Dim count As Integer
            While count + _dwSamples < BufferShort.Length
                Marshal.Copy(BufferShort, count, _PtrInput, _dwSamples)
                If WinAPI.beEncodeChunk(_HandleStream, _dwSamples, _PtrInput, _OutputBuffer, _BytesOutput) = 0 Then
                    If _BytesOutput > 0 Then
                        Dim bytes(_BytesOutput - 1) As Byte
                        Buffer.BlockCopy(_OutputBuffer, 0, bytes, 0, _BytesOutput)
                        mp3bytes = CombineBytes(mp3bytes, bytes)
                    End If
                End If
                count += _dwSamples
            End While

            'encodes remaining
            If BufferShort.Length - count > -1 Then
                Marshal.Copy(BufferShort, count, _PtrInput, BufferShort.Length - count)
                If WinAPI.beEncodeChunk(_HandleStream, CShort(BufferShort.Length - count), _PtrInput, _OutputBuffer, _BytesOutput) = 0 Then
                    If _BytesOutput > 0 Then
                        Dim bytes(_BytesOutput - 1) As Byte
                        Buffer.BlockCopy(_OutputBuffer, 0, bytes, 0, _BytesOutput)
                        mp3bytes = CombineBytes(mp3bytes, bytes)
                    End If
                End If
            End If
            SendOut(mp3bytes, Str)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub SendOut(MP3Bytes() As Byte, Str() As String)

        Try
            If fMain.BCCalls1 IsNot Nothing Then fMain.BCCalls1.AudioMP3(MP3Bytes, Str)
        Catch ex As Exception
        End Try

    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            On Error Resume Next
            WinAPI.beDeinitStream(_HandleStream, _OutputBuffer, _BytesOutput)
            WinAPI.beCloseStream(_HandleStream)
            If _HandleInput.IsAllocated Then _HandleInput.Free()
            _PtrInput = IntPtr.Zero
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