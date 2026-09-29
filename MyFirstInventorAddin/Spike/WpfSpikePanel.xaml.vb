Imports System.Collections.ObjectModel
Imports System.Windows.Media
Imports Serilog

Namespace iPropertiesController

    ' Throwaway spike panel: records which keys and characters actually arrive when the
    ' panel is hosted in an Inventor DockableWindow, so we can tell whether Inventor is
    ' stealing keystrokes before committing to a full WPF migration.
    Partial Public Class WpfSpikePanel

        Private ReadOnly _entries As New ObservableCollection(Of String)()
        Private _isDark As Boolean

        Public Sub New()
            InitializeComponent()
            KeyLog.ItemsSource = _entries
            AddEntry($"Panel loaded, hook = {WpfDockableHost.InterceptDialogKeys}")
        End Sub

        Public Sub ApplyTheme(dark As Boolean)
            _isDark = dark
            ' Same colours as iPropertiesAddInServer.SwitchTheme, so the two panels can be compared side by side.
            If dark Then
                SetBrush("PanelBack", 59, 68, 83)
                SetBrush("PanelFore", 225, 225, 225)
                SetBrush("ControlBack", 69, 79, 97)
                SetBrush("ControlHighlighted", 44, 52, 64)
            Else
                SetBrush("PanelBack", 240, 240, 240)
                SetBrush("PanelFore", 0, 0, 0)
                SetBrush("ControlBack", 255, 255, 255)
                SetBrush("ControlHighlighted", 225, 225, 225)
            End If
            AddEntry($"Theme applied: {If(dark, "dark", "light")}")
        End Sub

        Private Sub SetBrush(key As String, r As Byte, g As Byte, b As Byte)
            Resources(key) = New SolidColorBrush(Color.FromRgb(r, g, b))
        End Sub

        Private Sub Panel_PreviewKeyDown(sender As Object, e As System.Windows.Input.KeyEventArgs)
            Dim key = If(e.Key = System.Windows.Input.Key.System, e.SystemKey, e.Key)
            Dim modifiers = System.Windows.Input.Keyboard.Modifiers
            Dim chord = If(modifiers = System.Windows.Input.ModifierKeys.None, key.ToString(), $"{modifiers}+{key}")
            AddEntry($"KeyDown  {chord,-18} -> {DescribeTarget(e.OriginalSource)}")
        End Sub

        Private Sub Panel_PreviewTextInput(sender As Object, e As System.Windows.Input.TextCompositionEventArgs)
            AddEntry($"Text     '{e.Text.Replace(vbCr, "\r").Replace(vbTab, "\t")}' -> {DescribeTarget(e.OriginalSource)}")
        End Sub

        Private Sub Panel_SizeChanged(sender As Object, e As System.Windows.SizeChangedEventArgs)
            SizeText.Text = $"Panel size {ActualWidth:0} x {ActualHeight:0} (should track the dockable window)"
        End Sub

        Private Sub HookCheckBox_Changed(sender As Object, e As System.Windows.RoutedEventArgs)
            WpfDockableHost.InterceptDialogKeys = HookCheckBox.IsChecked.GetValueOrDefault()
            AddEntry($"WM_GETDLGCODE hook {If(WpfDockableHost.InterceptDialogKeys, "ON", "OFF")}")
        End Sub

        Private Sub ToggleTheme_Click(sender As Object, e As System.Windows.RoutedEventArgs)
            ApplyTheme(Not _isDark)
        End Sub

        Private Sub CopyLog_Click(sender As Object, e As System.Windows.RoutedEventArgs)
            System.Windows.Clipboard.SetText(String.Join(Environment.NewLine, _entries))
            AddEntry("Log copied to clipboard")
        End Sub

        Private Sub ClearLog_Click(sender As Object, e As System.Windows.RoutedEventArgs)
            _entries.Clear()
        End Sub

        Private Shared Function DescribeTarget(source As Object) As String
            Dim element = TryCast(source, System.Windows.FrameworkElement)
            If element Is Nothing Then Return source?.GetType().Name
            Return If(String.IsNullOrEmpty(element.Name), element.GetType().Name, element.Name)
        End Function

        Private Sub AddEntry(message As String)
            Log.Debug("WpfSpike: {Message}", message)
            _entries.Insert(0, $"{DateTime.Now:HH:mm:ss.fff}  {message}")
            If _entries.Count > 200 Then _entries.RemoveAt(_entries.Count - 1)
        End Sub

    End Class

End Namespace
