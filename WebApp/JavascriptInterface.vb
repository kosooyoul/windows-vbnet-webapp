Imports System.Security.Permissions

Public Module JavascriptInterface
    <PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
    <System.Runtime.InteropServices.ComVisibleAttribute(True)>
    Public Class External

        Public MyWebAppForm As WebAppForm

        Public Sub New(form As WebAppForm)
            Me.MyWebAppForm = form
        End Sub

        Public Sub Exec(command As String, Optional show As Boolean = True, Optional wait As Boolean = False)
            If show Then
                Shell(command, AppWinStyle.NormalFocus, wait, -1)
            Else
                Shell(command, AppWinStyle.Hide, wait, -1)
            End If
        End Sub

        Public Sub OpenFile(callbackStr As String)
            Me.MyWebAppForm.OpenFile(callbackStr)
        End Sub

        Public Sub Command(command As String, callbackStr As String)
            Dim p As Process = New Process
            Dim output As String

            With p
                .StartInfo.CreateNoWindow = True
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.UseShellExecute = False
                .StartInfo.FileName = "cmd"
                .StartInfo.Arguments = "/c " & command
                .Start()
                output = .StandardOutput.ReadToEnd
                .WaitForExit()
            End With

            Me.MyWebAppForm.WebCallback(output, callbackStr)
        End Sub

        Public Sub MoveCursor(x As Integer, y As Integer)
            Me.MyWebAppForm.MoveCursor(x, y)
        End Sub

    End Class

End Module
