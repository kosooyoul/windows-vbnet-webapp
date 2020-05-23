Imports Newtonsoft.Json.Linq

Public Class WebAppForm
    Implements ILoadHandler

    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Public Const MOUSEEVENTF_ABSOLUTE As UInt32 = &HE8000
    Public Const MOUSEEVENTF_LEFTDOWN As UInt32 = &HE0002
    Public Const MOUSEEVENTF_LEFTUP As UInt32 = &HE0004
    Public Const MOUSEEVENTF_HWHEEL As UInt32 = &HE1000
    Public Const MOUSEEVENTF_MIDDLEDOWN As UInt32 = &HE0020
    Public Const MOUSEEVENTF_MIDDLEUP As UInt32 = &HE0040
    Public Const MOUSEEVENTF_MOVE As UInt32 = &HE0001
    Public Const MOUSEEVENTF_RIGHTDOWN As UInt32 = &HE0008
    Public Const MOUSEEVENTF_RIGHTUP As UInt32 = &HE0010
    Public Const MOUSEEVENTF_WHEEL As UInt32 = &HE0800
    Public Const MOUSEEVENTF_XDOWN As UInt32 = &HE0080
    Public Const MOUSEEVENTF_XUP As UInt32 = &HE0100

    Public CurrentCursor As System.Drawing.Point = New System.Drawing.Point

    Private testURL = "http://vr.ahyane.net/ed514c40ab33a2c4dec53c87afea2de6"
    Private appURL = "file:///" & IO.Path.GetFullPath(".\app\index.html")

    Private WithEvents Browser As ChromiumWebBrowser
    Private ExternalJavascript As JavascriptInterface.External

    Public Instance As WebAppForm

    Public Title As String

    Public TriggerWebCallback As Boolean = False
    Public TriggerWebCallbackOutput As String
    Public TriggerWebCallbackStr As String
    Public TriggerOpenFile As Boolean = False
    Public TriggerOpenFileCallbackStr As String
    Public MyOpenFileDialog As OpenFileDialog = New OpenFileDialog()
    Public PrevCursorX As Integer = 0
    Public PrevCursorY As Integer = 0

    Public Sub New()
        Instance = Me
        Me.Text = "Now Initializing..."

        InitializeComponent()

        CefSharp.Cef.Initialize(New CefSettings() With {.CachePath = "cache"})

        Browser = New ChromiumWebBrowser(appURL) With {
            .Dock = DockStyle.Fill,
            .BrowserSettings = New BrowserSettings() With {
                .ApplicationCache = CefState.Enabled,
                .FileAccessFromFileUrls = CefState.Disabled,
                .Javascript = CefState.Enabled,
                .LocalStorage = CefState.Enabled,
                .WebSecurity = CefState.Disabled,
                .JavascriptOpenWindows = CefState.Disabled,
                .JavascriptDomPaste = CefState.Disabled
            },
            .AllowDrop = False
        }

        ExternalJavascript = New JavascriptInterface.External(Me)
        Browser.RegisterJsObject("external", ExternalJavascript)
        Browser.MenuHandler = New CustomMenuHandler()
        Browser.DragHandler = New CustomDragHandler()
        Browser.LoadHandler = Me

        PrevCursorX = Cursor.Position.X
        PrevCursorY = Cursor.Position.Y

        MainPanel.Controls.Add(Browser)
    End Sub

    Public Function BrowserEval(script)
        Try
            Dim task = Browser.EvaluateScriptAsync(script)
            task.Wait()

            Dim response = task.Result
            If response.Success Then
                Return response.Result
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub BrowserReload()
        Browser.Reload()
    End Sub

    Private Function BrowserLoaded()
        Return BrowserEval("'0'") = "0"
    End Function

    Private Delegate Sub DoUpdateTitleDelegate()
    Private Delegate Sub DoSyncTitleDelegate()

    Private Sub DoUpdateTitle()
        If Me.InvokeRequired Then
            Me.Invoke(New DoUpdateTitleDelegate(AddressOf DoUpdateTitle))
        Else
            Me.Text = Me.Title
        End If
    End Sub

    Private Sub DoSyncTitle()
        If Me.InvokeRequired Then
            Me.Invoke(New DoSyncTitleDelegate(AddressOf DoSyncTitle))
        Else
            Me.Text = BrowserEval("document.title")
        End If
    End Sub

    Private Sub UpdateTitle(text As String)
        Me.Title = text

        Dim thread As System.Threading.Thread = New System.Threading.Thread(AddressOf DoUpdateTitle)
        thread.Start()
    End Sub

    Private Sub SyncTitle()
        Dim thread As System.Threading.Thread = New System.Threading.Thread(AddressOf DoSyncTitle)
        thread.Start()
    End Sub

    Public Sub OnLoadingStateChange(browserControl As IWebBrowser, loadingStateChangedArgs As LoadingStateChangedEventArgs) Implements ILoadHandler.OnLoadingStateChange

    End Sub

    Public Sub OnFrameLoadStart(browserControl As IWebBrowser, frameLoadStartArgs As FrameLoadStartEventArgs) Implements ILoadHandler.OnFrameLoadStart
        Me.UpdateTitle("Now Loading...")
    End Sub

    Public Sub OnFrameLoadEnd(browserControl As IWebBrowser, frameLoadEndArgs As FrameLoadEndEventArgs) Implements ILoadHandler.OnFrameLoadEnd
        CheckTriggerTimer.Enabled = True
        Me.SyncTitle()
    End Sub

    Public Sub OnLoadError(browserControl As IWebBrowser, loadErrorArgs As LoadErrorEventArgs) Implements ILoadHandler.OnLoadError
        Me.UpdateTitle("Error")
    End Sub

    Public Class CustomDragHandler
        Implements IDragHandler

        Public Sub OnDraggableRegionsChanged(browserControl As IWebBrowser, browser As IBrowser, regions As IList(Of DraggableRegion)) Implements IDragHandler.OnDraggableRegionsChanged

        End Sub

        Public Function OnDragEnter(browserControl As IWebBrowser, browser As IBrowser, dragData As IDragData, mask As DragOperationsMask) As Boolean Implements IDragHandler.OnDragEnter
            Return True
        End Function
    End Class

    Public Class CustomMenuHandler
        Implements CefSharp.IContextMenuHandler

        Private Sub IContextMenuHandler_OnBeforeContextMenu(browserControl As IWebBrowser, browser As IBrowser, frame As IFrame, parameters As IContextMenuParams, model As IMenuModel) Implements IContextMenuHandler.OnBeforeContextMenu
            model.Clear()
        End Sub

        Private Sub IContextMenuHandler_OnContextMenuDismissed(browserControl As IWebBrowser, browser As IBrowser, frame As IFrame) Implements IContextMenuHandler.OnContextMenuDismissed

        End Sub

        Private Function IContextMenuHandler_OnContextMenuCommand(browserControl As IWebBrowser, browser As IBrowser, frame As IFrame, parameters As IContextMenuParams, commandId As CefMenuCommand, eventFlags As CefEventFlags) As Boolean Implements IContextMenuHandler.OnContextMenuCommand
            Return False
        End Function

        Private Function IContextMenuHandler_RunContextMenu(browserControl As IWebBrowser, browser As IBrowser, frame As IFrame, parameters As IContextMenuParams, model As IMenuModel, callback As IRunContextMenuCallback) As Boolean Implements IContextMenuHandler.RunContextMenu
            Return False
        End Function
    End Class

    Public Sub WebCallback(output As String, callbackStr As String)
        TriggerWebCallback = True
        TriggerWebCallbackOutput = output
        TriggerWebCallbackStr = callbackStr
    End Sub

    Public Sub OpenFile(callbackStr As String)
        TriggerOpenFile = True
        TriggerOpenFileCallbackStr = callbackStr
    End Sub

    Public Sub MoveCursor(x As Integer, y As Integer)
        'Me.Cursor = New Cursor(Cursor.Current.Handle)
        Cursor.Position = New Point(x, y)
    End Sub

    Public Sub MouseClick(key As String, state As String)
        If key = "left" Then
            If state = "down" Then
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 1)
            ElseIf state = "up" Then
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 1)
            ElseIf state = "click" Then
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 1)
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 1)
            End If
        ElseIf key = "middle" Then
            If state = "down" Then
                mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 1)
            ElseIf state = "up" Then
                mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 1)
            ElseIf state = "click" Then
                mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 1)
                mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 1)
            End If
        ElseIf key = "right" Then
            If state = "down" Then
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 1)
            ElseIf state = "up" Then
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 1)
            ElseIf state = "click" Then
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 1)
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 1)
            End If
        End If
    End Sub

    Public Sub ClipCursor(x As Integer, y As Integer, w As Integer, h As Integer)
        'Me.Cursor = New Cursor(Cursor.Current.Handle)
        Cursor.Clip = New Rectangle(x, y, w, h) ', Me.Size)
    End Sub

    Private Sub CheckTriggerTimer_Tick(sender As Object, e As EventArgs) Handles CheckTriggerTimer.Tick
        'If DetectKey.GetState(Keys.ShiftKey) = DetectKey.STATE_KEEP_DOWN And DetectKey.GetState(Keys.Escape) = DetectKey.STATE_DOWN Then
        'Me.Activate()
        'End If

        'Key
        Dim key As Integer
        Dim state As String
        For key = 0 To 255
            state = DetectKey.GetState(key)

            If state = DetectKey.STATE_DOWN Or state = DetectKey.STATE_UP Then
                Dim result As JObject = New JObject()
                If key = 1 Then
                    result.Add("button", "left")
                    result.Add("type", state)
                    result.Add("x", Cursor.Position.X)
                    result.Add("y", Cursor.Position.Y)
                    Browser.ExecuteScriptAsync("window.onGlobalMouse(" & result.ToString() & ")")
                ElseIf key = 2 Then
                    result.Add("button", "right")
                    result.Add("type", state)
                    result.Add("x", Cursor.Position.X)
                    result.Add("y", Cursor.Position.Y)
                    Browser.ExecuteScriptAsync("window.onGlobalMouse(" & result.ToString() & ")")
                ElseIf key = 4 Then
                    result.Add("button", "middle")
                    result.Add("type", state)
                    result.Add("x", Cursor.Position.X)
                    result.Add("y", Cursor.Position.Y)
                    Browser.ExecuteScriptAsync("window.onGlobalMouse(" & result.ToString() & ")")
                Else
                    result.Add("key", key)
                    result.Add("state", state)
                    Browser.ExecuteScriptAsync("window.onGlobalKey(" & result.ToString() & ")")
                End If
            End If
        Next

        'Mouse
        If Cursor.Position.X <> PrevCursorX Or Cursor.Position.Y <> PrevCursorY Then
            Dim result As JObject = New JObject()
            result.Add("type", "move")
            result.Add("x", Cursor.Position.X)
            result.Add("y", Cursor.Position.Y)
            Browser.ExecuteScriptAsync("window.onGlobalMouse(" & result.ToString() & ")")

            PrevCursorX = Cursor.Position.X
            PrevCursorY = Cursor.Position.Y
        End If

        If TriggerWebCallback = True Then
            TriggerWebCallback = False

            Dim result As JObject = New JObject()
            result.Add("return", TriggerWebCallbackOutput)

            Browser.ExecuteScriptAsync(TriggerWebCallbackStr & "(" & result.ToString() & ")")
        End If

        If TriggerOpenFile = True Then
            TriggerOpenFile = False

            If MyOpenFileDialog.ShowDialog <> DialogResult.Cancel Then

                Dim result As JObject = New JObject()
                result.Add("return", MyOpenFileDialog.FileName.ToString())

                Browser.ExecuteScriptAsync(TriggerOpenFileCallbackStr & "(" & result.ToString() & ")")
            End If
        End If
    End Sub

End Class