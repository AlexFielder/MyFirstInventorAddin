Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Interop

Namespace iPropertiesController

    ' WPF replacements for VB's MsgBox and InputBox, which on .NET 8 are implemented with
    ' Windows Forms (Microsoft.VisualBasic.Forms).
    Friend Module Dialogs

        Private Const DefaultTitle As String = "iProperties Controller"

        ' Same parameters and return values as MsgBox, so existing "= vbYes" checks keep working:
        ' MessageBoxResult and MsgBoxResult share their numeric values (OK=1, Cancel=2, Yes=6, No=7).
        Friend Function ShowMessage(prompt As String, Optional buttons As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional title As String = Nothing) As MsgBoxResult
            Dim wpfButtons As MessageBoxButton
            Select Case buttons And CType(7, MsgBoxStyle)
                Case MsgBoxStyle.YesNo
                    wpfButtons = MessageBoxButton.YesNo
                Case MsgBoxStyle.YesNoCancel
                    wpfButtons = MessageBoxButton.YesNoCancel
                Case MsgBoxStyle.OkCancel
                    wpfButtons = MessageBoxButton.OKCancel
                Case Else
                    wpfButtons = MessageBoxButton.OK
            End Select
            Dim result = MessageBox.Show(prompt, If(title, DefaultTitle), wpfButtons)
            Return CType(CInt(result), MsgBoxResult)
        End Function

        ' Same parameters as InputBox; returns an empty string when cancelled, as InputBox does.
        Friend Function ShowInputBox(prompt As String, Optional title As String = Nothing, Optional defaultResponse As String = "") As String
            Dim input As New TextBox With {.Text = defaultResponse, .Margin = New Thickness(0, 8, 0, 12)}
            Dim okButton As New Button With {.Content = "OK", .IsDefault = True, .Width = 75, .Margin = New Thickness(0, 0, 8, 0)}
            Dim cancelButton As New Button With {.Content = "Cancel", .IsCancel = True, .Width = 75}

            Dim buttonRow As New StackPanel With {.Orientation = Orientation.Horizontal, .HorizontalAlignment = HorizontalAlignment.Right}
            buttonRow.Children.Add(okButton)
            buttonRow.Children.Add(cancelButton)

            Dim layout As New StackPanel With {.Margin = New Thickness(12)}
            layout.Children.Add(New TextBlock With {.Text = prompt, .TextWrapping = TextWrapping.Wrap})
            layout.Children.Add(input)
            layout.Children.Add(buttonRow)

            Dim dialog As New Window With {
                .Title = If(title, DefaultTitle),
                .Content = layout,
                .Width = 400,
                .SizeToContent = SizeToContent.Height,
                .ResizeMode = ResizeMode.NoResize,
                .ShowInTaskbar = False,
                .WindowStartupLocation = WindowStartupLocation.CenterOwner
            }
            Dim helper As New WindowInteropHelper(dialog) With {.Owner = New IntPtr(AddinGlobal.InventorApp.MainFrameHWND)}

            AddHandler okButton.Click, Sub() dialog.DialogResult = True
            AddHandler dialog.Loaded, Sub()
                                          input.Focus()
                                          input.SelectAll()
                                      End Sub

            Return If(dialog.ShowDialog().GetValueOrDefault(), input.Text, String.Empty)
        End Function

    End Module

End Namespace
