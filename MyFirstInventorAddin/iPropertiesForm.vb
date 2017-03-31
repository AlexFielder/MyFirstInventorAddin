Imports System.Windows.Forms
Imports Inventor
Imports log4net
Imports iPropertiesController.iPropertiesController

Public Class iPropertiesForm
    Private inventorApp As Inventor.Application
    Private localWindow As DockableWindow
    Private value As String

    Public ReadOnly log As ILog = LogManager.GetLogger(GetType(iPropertiesForm))

    Public Sub New(ByVal inventorApp As Inventor.Application, ByVal addinCLS As String, ByRef localWindow As DockableWindow)
        Try
            log.Debug("Loading iProperties Form")
            InitializeComponent()
            Me.inventorApp = inventorApp
            Me.value = addinCLS
            Me.localWindow = localWindow
            Dim uiMgr As UserInterfaceManager = inventorApp.UserInterfaceManager
            Dim myDockableWindow As DockableWindow = uiMgr.DockableWindows.Add(addinCLS, "iPropertiesControllerWindow", "My Add-in Dock")
            myDockableWindow.AddChild(Me.Handle)

            If Not myDockableWindow.IsCustomized = True Then
                'myDockableWindow.DockingState = DockingStateEnum.kFloat
                myDockableWindow.DockingState = DockingStateEnum.kDockLastKnown
            Else
                myDockableWindow.DockingState = DockingStateEnum.kFloat
            End If

            myDockableWindow.DisabledDockingStates = DockingStateEnum.kDockTop + DockingStateEnum.kDockBottom

            Me.Dock = DockStyle.Fill
            Me.Visible = True
            localWindow = myDockableWindow
            AddinGlobal.DockableList.Add(myDockableWindow)
        Catch ex As Exception
            log.Error(ex.Message)
        End Try
        log.Info("iProperties Form Loaded")

    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If TextBox1.Text = "Part Number" Then
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          "",
                                                          "")
                Else
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          TextBox1.Text,
                                                          "")
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Part Number Updated to: " + iPropPartNum)
                End If
            End If
        End If
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If TextBox2.Text = "Description" Then
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                          "",
                                                          "")
                Else
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                          TextBox2.Text,
                                                          "")
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Description Updated to: " + iPropPartNum)
                End If
            End If
        End If
    End Sub

    Private Sub TextBox3_Leave(sender As Object, e As EventArgs) Handles TextBox3.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If TextBox3.Text = "Stock Number" Then
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                          "",
                                                          "")
                Else
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                          TextBox3.Text,
                                                          "")
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Stock Number Updated to: " + iPropPartNum)
                End If
            End If
        End If
    End Sub

    Private Sub TextBox4_Leave(sender As Object, e As EventArgs) Handles TextBox4.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If TextBox4.Text = "Engineer" Then
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                          "",
                                                          "")
                Else
                    Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                          TextBox4.Text,
                                                          "")
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Engineer Updated to: " + iPropPartNum)
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'need to decide whether or not to leave our textbox.leave events as they are or change all to be driven by 
        'clicking this button.
        'If we do change, we need to error check the one time for activedocument and filename.length

        'For Each TxtBox As Windows.Forms.TextBox In Me.Controls
        '    ' this may or may not work because of the different types available within the controls collection.
        'Next

        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                inventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                'try these and see if they fire or not!
                TextBox1_Leave(sender, e)
                TextBox2_Leave(sender, e)
                TextBox3_Leave(sender, e)
                TextBox4_Leave(sender, e)
                TextBox7_Leave(sender, e)

                Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                Dim kgMass As Decimal = myMass / 1000
                Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                TextBox5.Text = myMass2 & " kg"
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + TextBox5.Text)

                Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                TextBox6.Text = myDensity2 & " g/cm^3"
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + TextBox6.Text)

                Label12.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")

            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Toggle 'Defer updates' on and off in a Drawing
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                    inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = False
                    'DrawingSettings.DeferUpdates = False
                    Label8.Text = "Drawing Updates Not Deferred"
                ElseIf iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                    inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = True
                    Label8.Text = "Drawing Updates Deferred"
                End If
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        inventorApp.ActiveDocument.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = DateTimePicker1.Value

    End Sub

    Private Sub TextBox7_Leave(sender As Object, e As EventArgs) Handles TextBox7.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForSummaryInformationEnum.kAuthorSummaryInformation,
                                                          TextBox7.Text,
                                                          "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Author Updated to: " + iPropPartNum)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim doc = inventorApp.ActiveDocument
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            oBOM.StructuredViewEnabled = True

            Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

            Dim oBOMRow As BOMRow

            For Each oBOMRow In oBOMView.BOMRows
                'Set a reference to the primary ComponentDefinition of the row
                Dim oCompDef As ComponentDefinition
                oCompDef = oBOMRow.ComponentDefinitions.Item(1)

                Dim CompFullDocumentName As String = oCompDef.Document.FullDocumentName
                Dim CompFileNameOnly As String
                Dim index As Integer = CompFullDocumentName.LastIndexOf("\")

                CompFileNameOnly = CompFullDocumentName.Substring(index + 1)

                'MessageBox.Show(CompFileNameOnly)

                Dim item As String
                item = oBOMRow.ItemNumber
                iProperties.SetorCreateCustomiProperty(oCompDef.Document, "#ITEM", item)
            Next
        Catch ex As Exception
            log.Error(ex.Message)
        End Try
    End Sub

    Private Sub TextBox5_Enter(sender As Object, e As EventArgs) Handles TextBox5.Enter
        Clipboard.SetText(TextBox5.Text)
        UpdateStatusBar("Mass Copied to Clipboard")
    End Sub
    Private Sub TextBox5_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox5.MouseClick
        TextBox5_Enter(sender, e)
    End Sub

    Private Sub TextBox6_Enter(sender As Object, e As EventArgs) Handles TextBox6.Enter
        Clipboard.SetText(TextBox6.Text)
        UpdateStatusBar("Density Copied to Clipboard")
    End Sub
    Private Sub TextBox6_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox6.MouseClick
        TextBox6_Enter(sender, e)
    End Sub

    Private Sub UpdateStatusBar(ByVal Message As String)
        AddinGlobal.InventorApp.StatusBarText = Message
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim MaterialTextBox As Inventor.TextBox = Nothing
        Try
            Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
            Dim oSheet As Sheet = oDrawDoc.ActiveSheet
            Dim oSheets As Sheets = Nothing
            Dim oViews As DrawingViews = Nothing
            Dim oScale As Double = Nothing
            Dim oViewCount As Integer = 0
            Dim oTitleBlock = oSheet.TitleBlock
            Dim oDWG As DrawingDocument = inventorApp.ActiveDocument

            Dim oSht As Sheet = oDWG.ActiveSheet

            Dim oView As DrawingView = Nothing
            Dim drawnDoc As Document = Nothing

            For i As Integer = 1 To oSht.DrawingViews.Count
                oView = oSht.DrawingViews(i)
                Exit For
            Next

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            prtMaterial = InputBox("leaving as 'Engineer' will bring through Engineer info from part, " &
                                   vbCrLf & "'PRT'or 'prt' will use part material, otherwise enter desired material info", "Material", "Engineer")

            MaterialTextBox = GetMaterialTextBox(oTitleBlock.Definition)
            Dim MaterialString As String = String.Empty
            MaterialString = prtMaterial
            If prtMaterial = "Engineer" Then
                MaterialString = iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                              "",
                                                              "")
            ElseIf UCase(prtMaterial) = "PRT" Then
                MaterialString = UCase(iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties,
                                                              "",
                                                              ""))
            Else
                MaterialString = prtMaterial
            End If
            oTitleBlock.SetPromptResultText(MaterialTextBox, MaterialString)

        Catch ex As Exception When MaterialTextBox Is Nothing
            UpdateStatusBar("No compatible drawing open!")
            log.Error(ex.Message)
        Catch ex As Exception
            log.Error(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
        Dim oSheet As Sheet = oDrawDoc.ActiveSheet
        Dim oSheets As Sheets = Nothing
        Dim oView As DrawingView = Nothing
        Dim oViews As DrawingViews = Nothing
        Dim oScale As Double = Nothing
        Dim oViewCount As Integer = 0

        Dim drawingDoc As DrawingDocument = TryCast(inventorApp.ActiveDocument, DrawingDocument)

        dwgScale = InputBox("If you leave as 'Scale From Sheet View' then it will use base view scale, otherwise enter scale to show", "Sheet Scale", "Scale From Sheet View")
        Dim oTitleBlock = oSheet.TitleBlock
        Dim scaleTextBox As Inventor.TextBox = GetScaleTextBox(oTitleBlock.Definition)
        Dim scaleString As String = String.Empty
        For Each viewX As DrawingView In oSheet.DrawingViews
            If (Not String.IsNullOrEmpty(viewX.ScaleString)) Then
                If dwgScale = "Scale From Sheet View" Then
                    scaleString = viewX.ScaleString
                Else
                    scaleString = dwgScale

                    Exit For
                End If

            End If
        Next
        oTitleBlock.SetPromptResultText(scaleTextBox, scaleString)
    End Sub

    Function GetMaterialTextBox(ByVal titleDef As TitleBlockDefinition) As Inventor.TextBox
        For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
            If (defText.Text = "<Material>" Or defText.Text = "Material") Then
                Return defText
            End If
        Next
        Return Nothing
    End Function

    Function GetScaleTextBox(ByVal titleDef As TitleBlockDefinition) As Inventor.TextBox
        For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
            If (defText.Text = "<Scale>" Or defText.Text = "Scale") Then
                Return defText
            End If
        Next
        Return Nothing
    End Function
End Class