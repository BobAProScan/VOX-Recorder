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

    Private Shared Function FromUnixTime(epoch As Long) As DateTime
        Dim origin As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Return origin.AddSeconds(epoch)
    End Function


    Public Shared Function SendToRdio(Freq As String, Duration As String, Epoch As String, Audio() As Byte) As String
        If Not fMain.RdioMode.Checked Then Return "Rdio upload skipped: Mode unchecked."
        If String.IsNullOrEmpty(fMain.RDIO_ApiKey.Text) Then Return "Rdio upload ERROR: API key is empty."
        If String.IsNullOrEmpty(fMain.RDIO_SystemID.Text) Then Return "Rdio upload ERROR: System ID is empty."
        If String.IsNullOrEmpty(fMain.RDIO_TalkgroupID.Text) Then Return "Rdio upload ERROR: Talkgroup ID is empty."
        If String.IsNullOrEmpty(fMain.RDIO_Url.Text) Then Return "Rdio upload ERROR: URL is empty."




        SyncLock Obj
            Dim postParameters As New Dictionary(Of String, Object)()
            Dim postURL As String = fMain.RDIO_Url.Text.TrimEnd("/"c) & "/api/call-upload"
            Dim callTime As DateTime = FromUnixTime(CLng(Epoch))
            Dim startEpoch As Long = CLng(Epoch)
            Dim endEpoch As Long = startEpoch + CLng(Math.Round(Double.Parse(Duration)))
            Dim callNumber As String = (DateTime.Now.Ticks Mod 100000).ToString("D5") ' crude unique-ish ID
            Dim fileName As String = String.Format("{0}_{1}_{2}_{3}.mp3", fMain.RDIO_TalkgroupID.Text, startEpoch, endEpoch, callNumber)
            postParameters.Add("dateTime", callTime.ToString("o"))
            postParameters.Add("audioName", fileName)
            postParameters.Add("audioType", "audio/mpeg")
            postParameters.Add("frequency", Freq)
            postParameters.Add("key", fMain.RDIO_ApiKey.Text)
            postParameters.Add("source", "0")
            postParameters.Add("system", fMain.RDIO_SystemID.Text)
            postParameters.Add("talkgroup", fMain.RDIO_TalkgroupID.Text)
            postParameters.Add("audio", Audio)

            Dim responseString As String = Rdio_SendMultipartFormData(postURL, postParameters)
            Return responseString
        End SyncLock
    End Function

    Public Shared Function Rdio_SendMultipartFormData(url As String, formFields As Dictionary(Of String, Object)) As String
        Dim boundary As String = "------------------------" & DateTime.Now.Ticks.ToString("x")
        Dim newLine As String = vbCrLf
        Dim encoding As Text.Encoding = Text.Encoding.UTF8

        Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "multipart/form-data; boundary=" & boundary
        request.KeepAlive = True

        Using requestStream As Stream = request.GetRequestStream()
            For Each field As KeyValuePair(Of String, Object) In formFields
                Dim isFile As Boolean = TypeOf field.Value Is Byte() AndAlso field.Key.ToLower() = "audio"

                If isFile Then
                    ' Use audioName and audioType for headers
                    Dim fileName As String = If(formFields.ContainsKey("audioName"), CStr(formFields("audioName")), "file.mp3")
                    Dim mimeType As String = If(formFields.ContainsKey("audioType"), CStr(formFields("audioType")), "application/octet-stream")

                    Dim fileHeader As String = $"--{boundary}{newLine}" &
                                           $"Content-Disposition: form-data; name=""{field.Key}""; filename=""{fileName}""{newLine}" &
                                           $"Content-Type: {mimeType}{newLine}{newLine}"
                    requestStream.Write(encoding.GetBytes(fileHeader), 0, encoding.GetByteCount(fileHeader))
                    requestStream.Write(CType(field.Value, Byte()), 0, CType(field.Value, Byte()).Length)
                    requestStream.Write(encoding.GetBytes(newLine), 0, encoding.GetByteCount(newLine))
                Else
                    Dim fieldData As String = $"--{boundary}{newLine}" &
                                          $"Content-Disposition: form-data; name=""{field.Key}""{newLine}{newLine}" &
                                          CStr(field.Value) & newLine
                    requestStream.Write(encoding.GetBytes(fieldData), 0, encoding.GetByteCount(fieldData))
                End If
            Next

            ' End boundary
            Dim endBoundary As String = $"--{boundary}--{newLine}"
            requestStream.Write(encoding.GetBytes(endBoundary), 0, encoding.GetByteCount(endBoundary))
        End Using

        ' Read the response
        Try
            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream())
                    Return reader.ReadToEnd()
                End Using
            End Using
        Catch ex As WebException
            Using reader As New StreamReader(ex.Response.GetResponseStream())
                Return "Upload failed: " & reader.ReadToEnd()
            End Using
        End Try
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

