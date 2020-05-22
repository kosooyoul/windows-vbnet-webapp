Imports System.Runtime.InteropServices

Module DetectKey
    Public STATE_UP = "up"
    Public STATE_KEEP_UP = "keep_up"
    Public STATE_DOWN = "down"
    Public STATE_KEEP_DOWN = "keep_down"

    Private Declare Ansi Function GetAsyncKeyState Lib "user32" (ByVal vKey As Int32) As Int32

    Private States As New Dictionary(Of String, Boolean)()

    Public ReadOnly Property ControlKey() As Boolean
        Get
            Return ((GetAsyncKeyState(CInt(Keys.ControlKey)) And &H8000) <> 0)
        End Get
    End Property

    Public ReadOnly Property AltKey() As Boolean
        Get
            Return ((GetAsyncKeyState(CInt(Keys.Menu)) And &H8000) <> 0)
        End Get
    End Property

    Public ReadOnly Property ShiftKey() As Boolean
        Get
            Return ((GetAsyncKeyState(CInt(Keys.ShiftKey)) And &H8000) <> 0)
        End Get
    End Property

    Public Function IsKeyPressed(ByVal key As Keys) As Boolean
        Return ((GetAsyncKeyState(CInt(key)) And &H8000) <> 0)
    End Function

    <DllImport("user32.dll")>
    Private Function GetKeyState(ByVal nVirtKey As Keys) As Short
    End Function

    Public Function GetState(ByVal key As Keys) As String
        Dim rtn As Short = GetKeyState(key)
        Dim downed As Boolean

        downed = rtn <> 0 And rtn <> 1

        If States.Keys.Contains(CInt(key)) Then
            If downed = False Then
                States.Remove(CInt(key))
                Return STATE_UP
            End If
            Return STATE_KEEP_DOWN
        Else
            If downed = True Then
                States.Add(CInt(key), True)
                Return STATE_DOWN
            End If
            Return STATE_KEEP_UP
        End If

    End Function

End Module
