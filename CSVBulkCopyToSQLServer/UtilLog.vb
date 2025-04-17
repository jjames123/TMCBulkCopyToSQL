

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Imports System.IO
Imports System.Configuration
Public Class UtilLog
    Public Shared Sub LogInfo(Info As String)
        Dim LogPath As String = SQLAction.GetLogPath()
        Using w As StreamWriter = File.AppendText(LogPath & Convert.ToString("\log.txt"))

            Log(Info, w)
        End Using

    End Sub

    Public Shared Sub Log(logMessage As String, w As TextWriter)
        Try
            w.Write(vbCr & vbLf & "Log Entry : ")
            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            w.WriteLine("  :")
            w.WriteLine("  :{0}", logMessage)
            w.WriteLine("-------------------------------")

        Catch ex As Exception
        End Try
    End Sub

End Class