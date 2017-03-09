Imports System.Windows.Forms
Imports log4net

Namespace MyFirstInventorAddin
    Friend Class ButtonActions
        Public Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ButtonActions))
        Friend Shared Sub Button1_Execute()
            log.Info("button clicked")
            MessageBox.Show("Hello World!")
        End Sub
    End Class
End Namespace
