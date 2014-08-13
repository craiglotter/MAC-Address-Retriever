Imports System
Imports System.Management
Imports System.IO
Imports System.Reflection

Module Application_Code

    Sub Main()
        Try

            Dim mc As System.Management.ManagementClass
            Dim mo As ManagementObject
            mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            Dim found As Integer = 0
            For Each mo In moc
                If mo.Item("IPEnabled") = True Then
                    Console.WriteLine(mo.Item("MacAddress").ToString())
                    found = found + 1
                End If
            Next
            If found = 0 Then
                Console.WriteLine("Fail. No IPEnabled MAC addresses located.")
            End If
        Catch ex As Exception
            Console.WriteLine("Fail. Check Error Log for more details.")
            Error_Handler(ex, "Main Code")
        End Try
    End Sub

    Private Function ApplicationPath() As String
        Return _
        Path.GetDirectoryName([Assembly].GetEntryAssembly().Location)
    End Function

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_TFSR_Error_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Console.WriteLine("An error occurred in MAC Address Retriever's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

End Module
