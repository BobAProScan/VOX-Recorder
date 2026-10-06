Option Strict Off

Public Class Mp3FrameDetector
    ' Common MP3 bitrates index matrix (Kbps) - simplified for Layer III, MPEG 1 & 2
    Private Shared ReadOnly BitratesMpeg1L3 As Integer() = {0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, -1}
    Private Shared ReadOnly BitratesMpeg2L3 As Integer() = {0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, -1}

    ' Common MP3 sample rates matrix (Hz)
    Private Shared ReadOnly SampleRatesMpeg1 As Integer() = {44100, 48000, 32000, -1}
    Private Shared ReadOnly SampleRatesMpeg2 As Integer() = {22050, 24000, 16000, -1}

    Public Shared Function TryParseHeader(header As Byte(), ByRef frameLength As Integer) As Boolean

        frameLength = 0
        If header Is Nothing OrElse header.Length < 4 Then Return False

        ' 1. Check sync word (First byte 0xFF, first 3 bits of second byte are 1s)
        If header(0) <> &HFF OrElse (header(1) And &HE0) <> &HE0 Then Return False

        ' 2. Decode versions and layers
        Dim mpegVersionBits As Integer = (header(1) And &H18) >> 3 ' 3 = MPEG 1, 2 = MPEG 2
        Dim layerBits = (header(1) And &H6) >> 1       ' 1 = Layer III (MP3)

        If layerBits <> 1 Then Return False ' Not Layer III

        ' 3. Extract Bitrate and Sample Rate indexes
        Dim bitrateIndex = (header(2) And &HF0) >> 4
        Dim sampleRateIndex = (header(2) And &HC) >> 2
        Dim paddingBit = (header(2) And &H2) >> 1

        If bitrateIndex = &HF OrElse bitrateIndex = &H0 Then Return False ' Invalid/Free bitrates
        If sampleRateIndex = &H3 Then Return False                     ' Invalid sample rate

        ' Determine actual values based on MPEG version
        Dim bitrate = If(mpegVersionBits = 3, BitratesMpeg1L3(bitrateIndex), BitratesMpeg2L3(bitrateIndex))
        Dim sampleRate = If(mpegVersionBits = 3, SampleRatesMpeg1(sampleRateIndex), SampleRatesMpeg2(sampleRateIndex))

        If bitrate <= 0 OrElse sampleRate <= 0 Then Return False

        ' 4. Calculate frame size (Layer III formula)
        ' Coefficients: 144 for MPEG 1 Layer III, 72 for MPEG 2 Layer III
        Dim coefficient = If(mpegVersionBits = 3, 144, 72)
        frameLength = coefficient * (bitrate * 1000) / sampleRate + paddingBit

        Return True

    End Function

    Public Shared Sub FindMp3Frames(Audio() As Byte, ByRef IndexFirstFrame As Integer, ByRef IndexLastFrame As Integer)

        Dim frameLength As Integer = 0

        ' Loop through the data leaving room for a 4-byte header
        For i As Integer = 0 To Audio.Length - 4 - 1
            ' Look for potential sync word
            If Audio(i) = &HFF AndAlso (Audio(i + 1) And &HE0) = &HE0 Then
                Dim potentialHeader = New Byte(3) {}
                Array.Copy(Audio, i, potentialHeader, 0, 4)
                If TryParseHeader(potentialHeader, frameLength) Then
                    If IndexFirstFrame = -1 Then IndexFirstFrame = i
                    IndexLastFrame = i
                    i += frameLength - 1
                End If
            End If
        Next
    End Sub

End Class
