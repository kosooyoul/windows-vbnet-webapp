Imports System.Security.Permissions

Public Module JavascriptInterface
    <PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
    <System.Runtime.InteropServices.ComVisibleAttribute(True)>
    Public Class External

        Public form As WebAppForm

        Public Sub New(form As WebAppForm)
            Me.form = form
        End Sub

        Public Sub exec(command As String, Optional show As Boolean = True, Optional wait As Boolean = False)
            If show Then
                Shell(command, AppWinStyle.NormalFocus, wait, -1)
            Else
                Shell(command, AppWinStyle.Hide, wait, -1)
            End If
        End Sub

        Public Sub openFile(callbackStr As String)
            form.OpenFile(callbackStr)
        End Sub


        Public Sub command(command As String, callbackStr As String)
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

            'MsgBox(output)
            'output = Trim(Mid(output, Len("shhell")))
            'output = Mid(output, 1, Math.Max(InStr(output, " ") - 1, 0))

            form.WebCallback(output, callbackStr)
            ' MessageBox.Show(Len(output) & " : " & output)

            ' Shell("adb shell kill -9 " + output, AppWinStyle.Hide, False, -1)
        End Sub

    End Class
End Module
