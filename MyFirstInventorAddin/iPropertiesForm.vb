Imports System.Windows.Forms
Imports Inventor
Imports log4net
Imports MyFirstInventorAddin.yFirstInventorAddin

Public Class iPropertiesForm
    Private inventorApp As Inventor.Application
    Private localWindow As DockableWindow
    Private value As String

    Public ReadOnly log As ILog = LogManager.GetLogger(GetType(iPropertiesForm))

    Public Sub New(ByVal inventorApp As Inventor.Application, ByVal addinCLS As String, ByRef localWindow As DockableWindow)
        InitializeComponent()
        Me.inventorApp = inventorApp
        Me.value = addinCLS
        Me.localWindow = localWindow
        Dim uiMgr As UserInterfaceManager = inventorApp.UserInterfaceManager
        Dim myDockableWindow As DockableWindow = uiMgr.DockableWindows.Add(addinCLS, "MyFirstWindow", "My Add-in Dock")
        myDockableWindow.AddChild(Me.Handle)

        If Not myDockableWindow.IsCustomized = True Then
            myDockableWindow.DockingState = DockingStateEnum.kFloat
        End If

        Me.Dock = DockStyle.Fill
        Me.Visible = True
        localWindow = myDockableWindow
    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          TextBox1.Text,
                                                          "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Part Number Updated to: " + iPropPartNum)
            End If
        End If
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                          TextBox2.Text,
                                                          "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Description Updated to: " + iPropPartNum)
            End If
        End If
    End Sub
End Class