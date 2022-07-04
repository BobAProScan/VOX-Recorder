Public Class HTTP

    Private Shared Obj As New Object
    Public Shared Function Send(ApiKey As String, SystemId As String, SlotID As String, Freq As String, Duration As String, Epoch As String, TestConnection As Boolean, Audio() As Byte) As String
        'https://briangrinstead.com/blog/multipart-form-post-in-c/

        SyncLock Obj
            Dim postParameters As New Dictionary(Of String, Object)()
            Dim postURL As String = "https://api.broadcastify.com/call-upload"
            If TestConnection = True Then
                postParameters.Add("test", "1")
                postParameters.Add("apiKey", ApiKey)
                postParameters.Add("systemId", SystemId)
            Else
                postParameters.Add("apiKey", ApiKey)
                postParameters.Add("systemId", SystemId)
                postParameters.Add("callDuration", Duration)
                postParameters.Add("ts", Epoch)
                postParameters.Add("tg", SlotID)
                postParameters.Add("src", "0")
                postParameters.Add("freq", Freq)
                postParameters.Add("enc", "mp3")
            End If

            Dim responseString As String = SendMultipartFormData(postURL, postParameters)
            If responseString.Contains("Error2") Then 'send again - 1 time if -> Error2: No such host is known. (api.broadcastify.com:443)
                Thread.Sleep(100)
                responseString = SendMultipartFormData(postURL, postParameters)
            End If
            If responseString.StartsWith("0 http") Then
                responseString = SendAudio(responseString.Substring(2), Audio)
                If responseString.Contains("Error3") Then
                    Thread.Sleep(100)
                    responseString = SendAudio(responseString.Substring(2), Audio) 'send again - 1 time if -> Error3: No such host is known. (s3.amazonaws.com:443)
                End If
            End If

            Return responseString

        End SyncLock

    End Function

    Private Shared Function SendMultipartFormData(postUrl As String, postParameters As Dictionary(Of String, Object)) As String

        Dim responseString As String = String.Empty
        Try
            Dim formDataBoundary As String = String.Format("----------{0:N}", Guid.NewGuid())
            Dim contentType As String = "multipart/form-data; boundary=" & formDataBoundary
            Dim formData As Byte() = GetMultipartFormData(postParameters, formDataBoundary)
            Dim request As HttpWebRequest = TryCast(WebRequest.Create(postUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = contentType
            request.UserAgent = ProgramName & " Version " & Version
            request.ContentLength = formData.Length
            request.Timeout = 8000
            Dim response As HttpWebResponse = Nothing
            Using requestStream As Stream = request.GetRequestStream()
                requestStream.Write(formData, 0, formData.Length)
            End Using
            response = TryCast(request.GetResponse(), HttpWebResponse)
            Using SR As New StreamReader(response.GetResponseStream())
                responseString = SR.ReadToEnd()
            End Using
            If response.StatusCode <> HttpStatusCode.OK Then
                responseString = "HTTP Error1: " & response.StatusDescription
            End If
        Catch ex As Exception
            responseString = "HTTP Error2: " & ex.Message
        End Try

        Return responseString

    End Function

    Private Shared Function GetMultipartFormData(postParameters As Dictionary(Of String, Object), boundary As String) As Byte()

        Dim formDataStream As Stream = New System.IO.MemoryStream()
        Dim needsCLRF As Boolean = False
        Dim encoding As Encoding = Encoding.UTF8

        For Each param As KeyValuePair(Of String, Object) In postParameters
            If needsCLRF Then formDataStream.Write(encoding.GetBytes(vbCrLf), 0, encoding.GetByteCount(vbCrLf))
            needsCLRF = True
            Dim postData As String = String.Format("--{0}" & vbCrLf & "Content-Disposition: form-data; name=""{1}""" & vbCrLf & vbCrLf & "{2}", boundary, param.Key, param.Value)
            formDataStream.Write(encoding.GetBytes(postData), 0, encoding.GetByteCount(postData))
        Next
        Dim footer As String = vbCrLf & "--" & boundary & "--" & vbCrLf
        formDataStream.Write(encoding.GetBytes(footer), 0, encoding.GetByteCount(footer))
        formDataStream.Position = 0
        Dim formData(CInt(formDataStream.Length - 1)) As Byte
        formDataStream.Read(formData, 0, formData.Length)
        formDataStream.Close()

        Return formData

    End Function

    Private Shared Function SendAudio(URL As String, Audio() As Byte) As String

        If Audio.Length = 0 Then Return "No audio to send error"

        Dim responseString As String = String.Empty
        Try
            Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)
            request.Method = "PUT"
            request.ContentType = "audio/mpeg"
            request.ContentLength = Audio.Length
            request.UserAgent = ProgramName & " Version " & Version
            request.Timeout = 8000
            Dim response As HttpWebResponse = Nothing
            Using stream As Stream = request.GetRequestStream()
                stream.Write(Audio, 0, Audio.Length)
                response = CType(request.GetResponse(), HttpWebResponse)
            End Using
            Using SR As New StreamReader(response.GetResponseStream())
                responseString = SR.ReadToEnd()
            End Using
        Catch ex As Exception
            responseString = "HTTP Error3: " & ex.Message
        End Try

        Return responseString

    End Function

End Class

