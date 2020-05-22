Imports System.Runtime.InteropServices

Module DetectKey
    Public Enum State
        UP
        KEEP_UP
        DOWN
        KEEP_DOWN
    End Enum

    Private Declare Ansi Function GetAsyncKeyState Lib "user32" (ByVal vKey As Int32) As Int32

    Private States As New Dictionary(Of Integer, Boolean)()

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

    Public Function GetState(ByVal key As Keys) As State
        Dim rtn As Short = GetKeyState(key)
        Dim downed As Boolean

        downed = rtn <> 0 And rtn <> 1

        If States.Keys.Contains(CInt(key)) Then
            If downed = False Then
                States.Remove(CInt(key))
                Return State.UP
            End If
            Return State.KEEP_DOWN
        Else
            If downed = True Then
                States.Add(CInt(key), True)
                Return State.DOWN
            End If
            Return State.KEEP_UP
        End If

    End Function

End Module
