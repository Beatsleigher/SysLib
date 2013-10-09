Imports System.Threading

Public Class Form1

    Private Sub BUtton1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each pr As Process In Process.GetProcessesByName("explorer.exe")
            pr.CloseMainWindow()
            Thread.Sleep(10000)
            If Not pr.HasExited Then
                pr.Kill()
                Thread.Sleep(3000)
            End If
        Next
    End Sub
End Class
