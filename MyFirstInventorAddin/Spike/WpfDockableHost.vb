Imports System.Windows.Interop

Namespace iPropertiesController

    ' Hosts a WPF element in a Win32 child window so it can be handed to DockableWindow.AddChild,
    ' which only accepts an HWND. WPF elements have no HWND of their own; the HwndSource provides one.
    Friend NotInheritable Class WpfDockableHost
        Implements IDisposable

        Private Const WS_CHILD As Integer = &H40000000
        Private Const WS_VISIBLE As Integer = &H10000000
        Private Const WS_CLIPCHILDREN As Integer = &H2000000

        Private Const WM_GETDLGCODE As Integer = &H87
        Private Const DLGC_WANTARROWS As Integer = &H1
        Private Const DLGC_WANTTAB As Integer = &H2
        Private Const DLGC_WANTALLKEYS As Integer = &H4
        Private Const DLGC_WANTCHARS As Integer = &H80

        ' Toggled from the spike panel so the same session can compare behaviour with and without the hook.
        Public Shared Property InterceptDialogKeys As Boolean = True

        Private ReadOnly _source As HwndSource
        ' Held in a field so the delegate stays alive and RemoveHook gets the same instance.
        Private ReadOnly _hook As HwndSourceHook

        Public Sub New(content As System.Windows.UIElement, parentHwnd As IntPtr, width As Integer, height As Integer)
            Dim parameters As New HwndSourceParameters("iPropertiesControllerWpfHost", width, height) With {
                .ParentWindow = parentHwnd,
                .WindowStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN
            }
            _source = New HwndSource(parameters)
            _hook = AddressOf WndProc
            _source.AddHook(_hook)
            _source.RootVisual = content
        End Sub

        Public ReadOnly Property Handle As IntPtr
            Get
                Return _source.Handle
            End Get
        End Property

        ' Inventor's dialog-style message handling asks the focused window which keys it wants.
        ' This answer is required: without it (tested in Inventor 2027) WPF still receives KeyDown
        ' but never WM_CHAR, so no text can be typed and Backspace/Delete do nothing.
        Private Function WndProc(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
            If msg = WM_GETDLGCODE AndAlso InterceptDialogKeys Then
                handled = True
                Return New IntPtr(DLGC_WANTARROWS Or DLGC_WANTTAB Or DLGC_WANTALLKEYS Or DLGC_WANTCHARS)
            End If
            Return IntPtr.Zero
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _source.RemoveHook(_hook)
            _source.Dispose()
        End Sub

    End Class

End Namespace
