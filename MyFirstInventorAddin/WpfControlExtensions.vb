Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Documents

Namespace iPropertiesController

    ' Small helpers that keep the panel code close to its Windows Forms original, where the
    ' add-in server shows, hides and recolours controls depending on the active document.
    Friend Module WpfControlExtensions

        <Extension>
        Friend Sub Show(element As UIElement)
            element.Visibility = Visibility.Visible
        End Sub

        <Extension>
        Friend Sub Hide(element As UIElement)
            element.Visibility = Visibility.Collapsed
        End Sub

        ' Returns text to the theme colour after it was highlighted (red = edited, blue = hover).
        ' TextBox, TextBlock and Button all share TextElement.ForegroundProperty.
        <Extension>
        Friend Sub ResetForeground(element As FrameworkElement)
            element.ClearValue(TextElement.ForegroundProperty)
        End Sub

    End Module

End Namespace
