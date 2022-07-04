Imports Zip

Public Class FrmCheckNewestVersion

    Private NewestPathFilename As String = String.Empty
    Private LatestVersion As String = String.Empty

    Private Sub FrmCheckNewestVersion_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Try
            Cursor = Cursors.WaitCursor
            Application.DoEvents()

            Using wc As New WebClient()
                Dim str1 As String = wc.DownloadString("https://www.proscan.org/currentversion.php")
                Dim data() As String = str1.Split(CChar(","))
                Select Case ProgramName
                    Case "ProScan"
                        If data.Length > 0 Then LatestVersion = data(0)
                    Case "ProScan Client"
                        If data.Length > 1 Then LatestVersion = data(1)
                    Case "RadioFeed"
                        If data.Length > 2 Then LatestVersion = data(2)
                    Case "RadioFeed - Broadcastify Edition"
                        If data.Length > 3 Then LatestVersion = data(3)
                    Case "VOX Recorder"
                        If data.Length > 4 Then LatestVersion = data(4)
                End Select
            End Using
            If String.IsNullOrEmpty(LatestVersion) = False AndAlso LatestVersion.Contains(".") Then
                If GetVersionValue(LatestVersion) > GetVersionValue(Version) Then
                    Label1.Text = "Version " & LatestVersion & " Now Available"
                    Button1.Enabled = True
                Else
                    Label1.Text = "Version " & Version & " Up To Date"
                    Button1.Enabled = False
                End If
            Else
                Using New CenteredMessageBox(Me)
                    MessageBox.Show(Me, "Unable To Retrive Latest Version" & vbCrLf & vbCrLf & "Try Again")
                End Using
            End If

            Cursor = Cursors.Arrow
        Catch ex As Exception
            Cursor = Cursors.Arrow
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, "Unable To Retrive Latest Version" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Try Again")
            End Using
        End Try

    End Sub

    Private Function GetVersionValue(Version As String) As Integer

        Dim data() As String

        data = Version.Split(CChar("."))
        If data.Length > 1 Then
            Return (CInt(data(0)) * 100) + (CInt(data(1)) * 10)
        Else
            Return 0
        End If

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim DownloadURL As String = "https://www.proscan.org/" & ProgramName.Replace(" ", "_") & "_" & LatestVersion.Replace(".", "_") & ".zip"
        NewestPathFilename = GetDownloadsPath() & "\" & Path.GetFileName(DownloadURL)
        Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Try
            Using wc As New WebClient()
                wc.DownloadFile(DownloadURL, NewestPathFilename)
            End Using
            Dim fi As FileInfo = New FileInfo(NewestPathFilename)
            Try
                Using zip As ZipFile = ZipFile.Read(fi.FullName)
                    For Each entry As ZipEntry In zip
                        entry.Extract(fi.DirectoryName, ExtractExistingFileAction.OverwriteSilently)
                    Next
                End Using
            Catch ex1 As Exception
            End Try
            Process.Start("explorer.exe", "/select," & NewestPathFilename)
            fMain.Close()
        Catch ex As Exception
            Cursor = Cursors.Arrow
            Using New CenteredMessageBox(Me)
                MessageBox.Show(Me, "Unable To Download Or Save Latest Version" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & "Try Again")
            End Using
        End Try
        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim p As New Process()
        p.StartInfo.UseShellExecute = True
        Select Case ProgramName
            Case "ProScan"
                p.StartInfo.FileName = "https://www.proscan.org/proscan_whats_new.html"
            Case "ProScan Client"
                p.StartInfo.FileName = "https://www.proscan.org/proscan_free_client_whats_new.html"
            Case "RadioFeed"
                p.StartInfo.FileName = "https://www.proscan.org/radiofeed_whats_new.html"
            Case "RadioFeed - Broadcastify Edition"
            Case "VOX Recorder"
                p.StartInfo.FileName = "https://www.proscan.org/vox_recorder_whats_new.html"
        End Select
        p.Start()

    End Sub

    Private Function GetDownloadsPath() As String
        If Environment.OSVersion.Version.Major >= 6 Then
            Dim pathPtr As IntPtr
            If WinAPI.SHGetKnownFolderPath(New Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, IntPtr.Zero, pathPtr) = 0 Then
                Dim path1 As String = Marshal.PtrToStringUni(pathPtr)
                Marshal.FreeCoTaskMem(pathPtr)
                Return path1
            End If
        End If
        Return Path.Combine(Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.Personal)), "Downloads")

    End Function

End Class
