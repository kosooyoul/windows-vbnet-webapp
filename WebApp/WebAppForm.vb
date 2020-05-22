Imports Newtonsoft.Json.Linq

Public Class WebAppForm

    Private testURL = "http://vr.ahyane.net/ed514c40ab33a2c4dec53c87afea2de6"
    Private appURL = "file:///" & IO.Path.GetFullPath(".\app\index.html")

    Private WithEvents Browser As ChromiumWebBrowser
    Private ExternalJavascript As JavascriptInterface.External

    Public Instance As WebAppForm

    Public TriggerWebCallback As Boolean = False
    Public TriggerWebCallbackOutput As String
    Public TriggerWebCallbackStr As String
    Public TriggerOpenFile As Boolean = False
    Public TriggerOpenFileCallbackStr As String
    Public MyOpenFileDialog As OpenFileDialog = New OpenFileDialog()

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
        SyncWebTimer.Enabled = True
    End Sub

    Private Function BrowserLoaded()
        Return BrowserEval("'0'") = "0"
    End Function

    Private Sub SyncWebTimer_Tick(sender As Object, e As EventArgs) Handles SyncWebTimer.Tick
        SyncWebTimer.Enabled = False

        If BrowserLoaded() = False Then
            Me.Text = "Now Loading..."
            SyncWebTimer.Enabled = True
            Exit Sub
        End If

        Me.Text = BrowserEval("document.title")
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

        Private Sub CheckTriggerTimer_Tick(sender As Object, e As EventArgs) Handles CheckTriggerTimer.Tick
        If DetectKey.GetState(Keys.ShiftKey) = DetectKey.State.KEEP_DOWN And DetectKey.GetState(Keys.Escape) = DetectKey.State.DOWN Then
            'Me.Text = "A DOWNED"
            Me.Activate()
        End If
        'If keyState = DetectKey.State.UP Then Me.Text = "A UPED"
        'If keyState = DetectKey.State.KEEP_DOWN Then Me.Text = "A KEEP DOWNED"
        'If keyState = DetectKey.State.KEEP_UP Then Me.Text = "A KEEP UPED"

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