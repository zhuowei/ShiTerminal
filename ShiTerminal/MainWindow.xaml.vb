Imports System.Diagnostics
Imports System.Threading

Class MainWindow
    Public mainProcess As Process
    Private Delegate Sub AppendDelegate(textData As String)
    Private myLock As New Object()
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim startInfo As New ProcessStartInfo("cmd.exe")
        startInfo.RedirectStandardError = True
        startInfo.RedirectStandardInput = True
        startInfo.RedirectStandardOutput = True
        startInfo.UseShellExecute = False
        startInfo.CreateNoWindow = True
        mainProcess = Process.Start(startInfo)
        startGobbling()
        Me.inputText.Focus()
    End Sub
    Private Sub startGobbling()
        Dim inThread As New Thread(AddressOf gobbleStdIn)
        inThread.Start()
        Dim errThread As New Thread(AddressOf gobbleStdErr)
        errThread.Start()
        Dim inThreadPost As New Thread(AddressOf postStdIn)
        inThreadPost.Start()
    End Sub
    Private stdoutBuf As String = ""
    Private stderrBuf As String = ""
    Private stdinKick As Long = 0
    Private Sub gobbleStdIn()
        Dim mystr As String
        While Not mainProcess.StandardOutput.EndOfStream
            mystr = Char.ConvertFromUtf32(mainProcess.StandardOutput.Read())
            SyncLock myLock
                stdoutBuf += mystr
            End SyncLock
        End While
    End Sub
    Private Sub postStdIn()
        Dim myAppend As New AppendDelegate(AddressOf scrollTheTextBlock)
        While True
            If stdoutBuf.Length > 0 Then
                SyncLock myLock
                    mainText.Dispatcher.Invoke(myAppend, stdoutBuf)
                    stdoutBuf = ""
                End SyncLock
            End If
            If stderrBuf.Length > 0 Then
                SyncLock myLock
                    mainText.Dispatcher.Invoke(myAppend, stderrBuf)
                    stderrBuf = ""
                End SyncLock
            End If
            If mainProcess.HasExited Then
                End
            End If
            Thread.Sleep(1)
        End While
    End Sub
    Private Sub scrollTheTextBlock(mystring As String)
        Dim mystr = mainText.Text + mystring
        Dim maxlen = (80 * 50)
        If mystr.Length < maxlen Then
            mainText.Text = mystr
        Else
            mainText.Text = mystr.Substring(mystr.Length - 1 - maxlen)
        End If
        mainScroll.ScrollToBottom()
    End Sub
    Private Sub gobbleStdErr()
        Dim mystr As String
        While Not mainProcess.StandardError.EndOfStream
            mystr = Char.ConvertFromUtf32(mainProcess.StandardError.Read())
            SyncLock myLock
                stderrBuf += mystr
            End SyncLock
        End While
    End Sub

    Private Sub inputText_KeyDown(sender As Object, e As KeyEventArgs) Handles inputText.KeyDown
        If e.Key = Key.Enter Then
            mainProcess.StandardInput.WriteLine(inputText.Text)
            inputText.Text = ""
        End If
    End Sub
End Class
