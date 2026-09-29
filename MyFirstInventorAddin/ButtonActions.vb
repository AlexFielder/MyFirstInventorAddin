Namespace iPropertiesController

    Friend Class ButtonActions
        'Public Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ButtonActions))

        Friend Shared Sub Button1_Execute()
            'log.Info("button clicked")
            ShowMessage("Hello World!")
        End Sub

    End Class

End Namespace