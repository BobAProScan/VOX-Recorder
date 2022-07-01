Public Class DSP

    Public Shared Sub ToneGenerator(Type As String, Buffer() As Short, Frequency As Integer, DBLevel As Integer, SamplesPerSec As Integer, Summed As Boolean, Channel As Integer)

        Dim i As Integer
        Dim dbl1 As Double
        Dim angle As Double
        Dim amplitude As Double = 10 ^ (DBLevel / 20) * 32767
        Dim count As Long
        Static count1 As Long

        If Channel = 1 Then 'stereo
            SamplesPerSec = CInt(SamplesPerSec * 2)
        End If

        count = count1

        Select Case Type
            Case "Sine"
                angle = (2 * Math.PI * Frequency) / (SamplesPerSec)
                For i = 0 To Buffer.Length - 1
                    If (Channel = 1) OrElse (Channel = 2 AndAlso i Mod 2 = 0) OrElse (Channel = 3 AndAlso i Mod 2 = 1) Then
                        If Summed = False Then
                            Buffer(i) = LimitShort(Math.Sin(count * angle), amplitude)
                        Else
                            Buffer(i) = LimitShort(Buffer(i) + Math.Sin(count * angle), amplitude)
                        End If
                        count += 1
                    Else
                        Buffer(i) = 0
                    End If
                Next
            Case "Square"
                angle = (2 * Frequency) / (SamplesPerSec)
                For i = 0 To Buffer.Length - 1
                    If (Channel = 1) OrElse (Channel = 2 AndAlso i Mod 2 = 0) OrElse (Channel = 3 AndAlso i Mod 2 = 1) Then
                        If count Mod 2 = 0 Then dbl1 = If(Math.Round(count * angle) Mod 2 = 0, amplitude, -amplitude)
                        If Summed = False Then
                            Buffer(i) = LimitShort(dbl1)
                        Else
                            Buffer(i) = LimitShort(Buffer(i) + dbl1)
                        End If
                        count += 1
                    Else
                        Buffer(i) = 0
                    End If
                Next
            Case "Triangular"
                angle = (2 * Frequency) / (SamplesPerSec)
                For i = 0 To Buffer.Length - 1
                    If (Channel = 1) OrElse (Channel = 2 AndAlso i Mod 2 = 0) OrElse (Channel = 3 AndAlso i Mod 2 = 1) Then
                        Dim ip As Integer = CInt(Math.Round(count * angle))
                        If count Mod 2 = 0 Then dbl1 = 2 * amplitude * (1 - 2 * (ip Mod 2)) * (count * angle - ip)
                        If Summed = False Then
                            Buffer(i) = LimitShort(dbl1)
                        Else
                            Buffer(i) = LimitShort(Buffer(i) + dbl1)
                        End If
                        count += 1
                    Else
                        Buffer(i) = 0
                    End If
                Next
            Case "Sawtooth"
                For i = 0 To Buffer.Length - 1
                    If (Channel = 1) OrElse (Channel = 2 AndAlso i Mod 2 = 0) OrElse (Channel = 3 AndAlso i Mod 2 = 1) Then
                        angle = (count * Frequency) / (SamplesPerSec)
                        If count Mod 2 = 0 Then dbl1 = 2 * amplitude * (angle - Math.Round(angle))
                        If Summed = False Then
                            Buffer(i) = LimitShort(dbl1)
                        Else
                            Buffer(i) = LimitShort(Buffer(i) + dbl1)
                        End If
                        count += 1
                    Else
                        Buffer(i) = 0
                    End If
                Next
            Case "Noise"
                Dim rnd As New Random()
                For i = 0 To Buffer.Length - 1
                    If (Channel = 1) OrElse (Channel = 2 AndAlso i Mod 2 = 0) OrElse (Channel = 3 AndAlso i Mod 2 = 1) Then
                        dbl1 = rnd.Next(-CInt(amplitude), CInt(amplitude))
                        If Summed = False Then
                            Buffer(i) = LimitShort(dbl1)
                        Else
                            Buffer(i) = LimitShort(Buffer(i) + dbl1)
                        End If
                    Else
                        Buffer(i) = 0
                    End If
                Next
        End Select

        count1 = count

    End Sub

End Class