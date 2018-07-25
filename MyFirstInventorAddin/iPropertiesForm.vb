Imports System.Windows.Forms
Imports Inventor
Imports iPropertiesController.iPropertiesController
Imports log4net

Public Class IPropertiesForm
    'Inherits Form
    Private inventorApp As Inventor.Application
    Private localWindow As DockableWindow
    Private value As String
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    Public ReadOnly log As ILog = LogManager.GetLogger(GetType(IPropertiesForm))

    Public Sub GetNewFilePaths()
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    CurrentPath = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)
                    NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)

                ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        'Do Nothing
                    Else
                        Dim oDrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                        Dim oSht As Sheet = oDrawDoc.ActiveSheet
                        Dim oView As DrawingView = Nothing

                        For Each view As DrawingView In oSht.DrawingViews
                            oView = view
                            Exit For
                        Next

                        CurrentPath = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)
                        RefCurrentPath = System.IO.Path.GetDirectoryName(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                        NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)

                        RefDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                        RefNewPath = RefCurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub AddReferences(ByVal odoc As Inventor.Document, ByVal selectedfile As String)
        Dim oleReference As ReferencedOLEFileDescriptor
        oleReference = odoc.ReferencedOLEFileDescriptors _
                    .Add(selectedfile, OLEDocumentTypeEnum.kOLEDocumentLinkObject)
        oleReference.BrowserVisible = True
        oleReference.Visible = False
        oleReference.DisplayName = System.IO.Path.GetFileName(selectedfile)
    End Sub

    Public CurrentPath As String = String.Empty
    Public NewPath As String = String.Empty
    Public RefNewPath As String = String.Empty
    Public RefDoc As Document = Nothing

    Public Sub New(ByVal inventorApp As Inventor.Application, ByVal addinCLS As String, ByRef localWindow As DockableWindow)
        Try
            log.Debug("Loading iProperties Form")
            InitializeComponent()

            'Me.KeyPreview = True
            Me.inventorApp = inventorApp
            Me.value = addinCLS
            Me.localWindow = localWindow
            Dim uiMgr As UserInterfaceManager = inventorApp.UserInterfaceManager
            Dim addinName As String = lbAddinName.Text
            Dim myDockableWindow As DockableWindow = uiMgr.DockableWindows.Add(addinCLS, "iPropertiesControllerWindow", "iProperties Controller " + addinName)
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

    Private Sub UpdateAllCommon()
        If Not inventorApp.ActiveDocument Is Nothing Then
            'If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
            inventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
            'try these and see if they fire or not!
            tbPartNumber_Leave(sender, e)
            tbDescription_Leave(sender, e)
            tbStockNumber_Leave(sender, e)
            tbEngineer_Leave(sender, e)
            tbDrawnBy_Leave(sender, e)
            tbRevNo_Leave(sender, e)
            tbComments_Leave(sender, e)
            tbNotes_Leave(sender, e)

            If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
                Dim AssyDoc As AssemblyDocument = Nothing
                AssyDoc = inventorApp.ActiveDocument
                If AssyDoc.SelectSet.Count = 1 Then
                    Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                    Dim selecteddoc As Document = compOcc.Definition.Document

                    Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(selecteddoc, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                    Dim kgMass As Decimal = myMass / 1000
                    Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                    tbMass.Text = myMass2 & " kg"
                    log.Debug(selecteddoc.FullFileName + " Mass Updated to: " + tbMass.Text)

                    Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(selecteddoc, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                    Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                    tbDensity.Text = myDensity2 & " g/cm^3"
                    log.Debug(selecteddoc.FullFileName + " Mass Updated to: " + tbDensity.Text)

                    Label12.Text = iProperties.GetorSetStandardiProperty(selecteddoc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")

                    AssyDoc.SelectSet.Select(compOcc)
                Else
                    Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                    Dim kgMass As Decimal = myMass / 1000
                    Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                    tbMass.Text = myMass2 & " kg"
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + tbMass.Text)

                    Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                    Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                    tbDensity.Text = myDensity2 & " g/cm^3"
                    log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + tbDensity.Text)

                    Label12.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                End If
            Else
                Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                Dim kgMass As Decimal = myMass / 1000
                Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                tbMass.Text = myMass2 & " kg"
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + tbMass.Text)

                Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                tbDensity.Text = myDensity2 & " g/cm^3"
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Mass Updated to: " + tbDensity.Text)

                Label12.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
            End If
            UpdateStatusBar("iProperties updated")
            'End If
            ErrorProvider1.Clear()
            If Me.ValidateChildren() Then
                ' continue on
            End If
        End If
    End Sub

    Private Sub CheckForDefaultAndUpdate(ByVal proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, ByVal propname As String, ByVal newPropValue As String)
        Dim iProp As String = String.Empty
        If TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then

            'Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
            'Dim oSht As Sheet = oDWG.ActiveSheet
            'Dim oView As DrawingView = Nothing
            'Dim drawnDoc As Document = Nothing


            'For Each view As DrawingView In oSht.DrawingViews
            '    oView = view
            '    Exit For
            'Next

            'drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            'If Not newPropValue = propname Then
            '    UpdateProperties(proptoUpdate, propname, newPropValue, iProp, drawnDoc)
            'End If
            If Not newPropValue = propname Then
                UpdateProperties(proptoUpdate, propname, newPropValue, iProp)
            End If

        ElseIf TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
                Dim AssyDoc As AssemblyDocument = Nothing
                AssyDoc = inventorApp.ActiveDocument
                If AssyDoc.SelectSet.Count = 1 Then
                    Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                    Dim selecteddoc As Document = compOcc.Definition.Document
                    If Not newPropValue = propname Then
                        UpdateProperties(proptoUpdate, propname, newPropValue, iProp, AssyDoc, selecteddoc)
                    End If
                    AssyDoc.SelectSet.Select(compOcc)
                ElseIf AssyDoc.SelectSet.Count = 0 Then
                    If Not newPropValue = propname Then
                        UpdateProperties(proptoUpdate, propname, newPropValue, iProp)
                    End If
                End If
            ElseIf TypeOf (inventorApp.ActiveDocument) Is PartDocument Then
                If inventorApp.ActiveEditObject IsNot Nothing Then
                    If Not newPropValue = propname Then
                        UpdateProperties(proptoUpdate, propname, newPropValue, iProp)
                    End If
                ElseIf inventorApp.ActiveDocument IsNot Nothing Then
                    If Not newPropValue = propname Then
                        UpdateProperties(proptoUpdate, propname, newPropValue, iProp)
                    End If
                End If
            ElseIf TypeOf (inventorApp.ActiveDocument) Is PresentationDocument Then
                Throw New NotImplementedException
        End If
    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String, drawnDoc As Document)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(drawnDoc, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(drawnDoc, proptoUpdate, newPropValue, "", True)
            log.Debug(inventorApp.ActiveDocument.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject, proptoUpdate, newPropValue, "", True)
            log.Debug(inventorApp.ActiveEditObject.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String, AssyDoc As AssemblyDocument, selecteddoc As Document)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(selecteddoc, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(selecteddoc, proptoUpdate, newPropValue, "", True)
            log.Debug(selecteddoc.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
            iPropertiesAddInServer.ShowOccurrenceProperties(AssyDoc)
        End If
    End Sub

    Private Sub SendSymbol(ByVal textbox As Object, symbol As String)
        Dim insertText = symbol

        'textbox.SelectedText = insertText
        Dim insertPos As Integer = textbox.SelectionStart

        textbox.Text = textbox.Text.Insert(insertPos, insertText)
        textbox.Focus()
        textbox.SelectionStart = insertPos + insertText.Length
    End Sub

    Private Sub tbStockNumber_Leave(sender As Object, e As EventArgs) Handles tbStockNumber.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            tbStockNumber.ForeColor = Drawing.Color.Black
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "Stock Number", tbStockNumber.Text)
        End If
    End Sub

    Private Sub tbEngineer_Leave(sender As Object, e As EventArgs) Handles tbEngineer.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            tbEngineer.ForeColor = Drawing.Color.Black
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "Engineer", tbEngineer.Text)
        End If
    End Sub

    Private Sub btUpdateAll_Click(sender As Object, e As EventArgs) Handles btUpdateAll.Click

        If Not inventorApp.ActiveDocument Is Nothing Then
            If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
                Dim assydoc As Document = Nothing
                assydoc = inventorApp.ActiveDocument
                If assydoc.SelectSet.Count = 1 Then
                    Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                    UpdateAllCommon()
                    assydoc.SelectSet.Select(compOcc)
                Else
                    UpdateAllCommon()
                End If
            Else
                UpdateAllCommon()
            End If
        End If
    End Sub

    Private Sub btDefer_Click(sender As Object, e As EventArgs) Handles btDefer.Click

        'Toggle 'Defer updates' on and off in a Drawing
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                    inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = False
                    'DrawingSettings.DeferUpdates = False
                    Label8.ForeColor = Drawing.Color.Green
                    Label8.Text = "Drawing Updates Not Deferred"
                    UpdateStatusBar("Drawing updates are no longer deferred")
                ElseIf iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                    inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = True
                    Label8.ForeColor = Drawing.Color.Red
                    Label8.Text = "Drawing Updates Deferred"
                    UpdateStatusBar("Drawing updates are now deferred")
                End If
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Not iPropertiesAddInServer.CheckReadOnly(AddinGlobal.InventorApp.ActiveDocument) Then
            If Not iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties, "", "") = DateTimePicker1.Value Then

                inventorApp.ActiveDocument.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = DateTimePicker1.Value
                UpdateStatusBar("Creation date updated to " + DateTimePicker1.Value)
            End If
        End If
    End Sub

    Private Sub tbDrawnBy_Leave(sender As Object, e As EventArgs) Handles tbDrawnBy.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then

            tbDrawnBy.ForeColor = Drawing.Color.Black

            Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForSummaryInformationEnum.kAuthorSummaryInformation,
                                                          tbDrawnBy.Text,
                                                          "")
            log.Debug(inventorApp.ActiveDocument.FullFileName + " Author Updated to: " + iPropPartNum)
            UpdateStatusBar("Author updated to " + iPropPartNum)

        End If
    End Sub

    Private Sub btITEM_Click(sender As Object, e As EventArgs) Handles btITEM.Click
        If TypeOf AddinGlobal.InventorApp.ActiveDocument Is DrawingDocument Then
            Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

            Dim oSht As Sheet = oDWG.ActiveSheet

            Dim oView As DrawingView = Nothing
            Dim drawnDoc As Document = Nothing

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            Dim doc = drawnDoc
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            oBOM.StructuredViewEnabled = True

            Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

            Dim oBOMRow As BOMRow

            For Each oBOMRow In oBOMView.BOMRows

                'Set a reference to the primary ComponentDefinition of the row
                Dim oCompDef As ComponentDefinition
                oCompDef = oBOMRow.ComponentDefinitions.Item(1)
                If oCompDef.Document.FullDocumentName.Contains("Content Center") Or oCompDef.Document.FullDocumentName.Contains("Bought Out") Then

                Else
                    Dim CompFullDocumentName As String = oCompDef.Document.FullDocumentName
                    Dim CompFileNameOnly As String
                    Dim index As Integer = CompFullDocumentName.LastIndexOf("\")

                    CompFileNameOnly = CompFullDocumentName.Substring(index + 1)

                    'MessageBox.Show(CompFileNameOnly)

                    Dim item As String
                    item = oBOMRow.ItemNumber
                    'iProperties.SetorCreateCustomiProperty(oCompDef.Document, "#ITEM", item)
                    iProperties.GetorSetStandardiProperty(
                                oCompDef.Document,
                                PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, item, "")
                End If
            Next
            oSht.Update()
        ElseIf TypeOf AddinGlobal.InventorApp.ActiveDocument Is AssemblyDocument Then
            Dim doc = inventorApp.ActiveEditDocument
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            oBOM.StructuredViewEnabled = True

            Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

            Dim oBOMRow As BOMRow

            For Each oBOMRow In oBOMView.BOMRows

                'Set a reference to the primary ComponentDefinition of the row
                Dim oCompDef As ComponentDefinition
                oCompDef = oBOMRow.ComponentDefinitions.Item(1)
                If oCompDef.Document.FullDocumentName.Contains("Content Center") Or oCompDef.Document.FullDocumentName.Contains("Bought Out") Then

                Else
                    Dim CompFullDocumentName As String = oCompDef.Document.FullDocumentName
                    Dim CompFileNameOnly As String
                    Dim index As Integer = CompFullDocumentName.LastIndexOf("\")

                    CompFileNameOnly = CompFullDocumentName.Substring(index + 1)

                    'MessageBox.Show(CompFileNameOnly)

                    Dim item As String
                    item = oBOMRow.ItemNumber
                    'iProperties.SetorCreateCustomiProperty(oCompDef.Document, "#ITEM", item)
                    iProperties.GetorSetStandardiProperty(
                                oCompDef.Document,
                                PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, item, "")
                End If
            Next
        End If
        UpdateStatusBar("BOM item numbers copied to #ITEM")
    End Sub

    Private Sub tbMass_Enter(sender As Object, e As EventArgs) Handles tbMass.Enter
        Clipboard.SetText(tbMass.Text)
        UpdateStatusBar("Mass copied to clipboard")
    End Sub

    Private Sub tbMass_MouseClick(sender As Object, e As MouseEventArgs) Handles tbMass.MouseClick
        tbMass_Enter(sender, e)
    End Sub

    Private Sub tbDensity_Enter(sender As Object, e As EventArgs) Handles tbDensity.Enter
        Clipboard.SetText(tbDensity.Text)
        UpdateStatusBar("Density copied to clipboard")
    End Sub

    Private Sub tbDensity_MouseClick(sender As Object, e As MouseEventArgs) Handles tbDensity.MouseClick
        tbDensity_Enter(sender, e)
    End Sub

    Private Sub UpdateStatusBar(ByVal Message As String)
        AddinGlobal.InventorApp.StatusBarText = Message
    End Sub

    Private Sub btShtMaterial_Click(sender As Object, e As EventArgs) Handles btShtMaterial.Click
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

            Dim oPromptEntry = Nothing

            Dim oCurrentSheet = Nothing
            oCurrentSheet = oDrawDoc.ActiveSheet.Name

            i = 1

            oTitleBlock = oSheet.TitleBlock
            oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
            For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
                Select Case oTextBox.Text
                    Case "<Material>"
                        oPromptEntry = oTitleBlock.GetResultText(oTextBox)
                End Select
            Next

            If oPromptEntry = "<Material>" Then
                oPromptText = "Engineer"
            ElseIf oPromptEntry = "" Then
                oPromptText = "Engineer"
            Else
                oPromptText = oPromptEntry
            End If


            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            prtMaterial = InputBox("leaving as 'Engineer' will bring through Engineer info from part, " &
                                   vbCrLf & "'PRT'or 'prt' will use part material, otherwise enter desired material info", "Material", oPromptText)

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
        UpdateStatusBar("Drawing material set")
    End Sub

    Private Sub btShtScale_Click(sender As Object, e As EventArgs) Handles btShtScale.Click
        Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
        Dim oSheet As Sheet = oDrawDoc.ActiveSheet
        Dim oSheets As Sheets = Nothing
        Dim oView As DrawingView = Nothing
        Dim oViews As DrawingViews = Nothing
        Dim oScale As Double = Nothing
        Dim oViewCount As Integer = 0
        Dim oTitleBlock = oSheet.TitleBlock
        Dim scaleTextBox As Inventor.TextBox = GetScaleTextBox(oTitleBlock.Definition)
        Dim scaleString As String = String.Empty

        Dim oPromptEntry = Nothing

        Dim oCurrentSheet = Nothing
        oCurrentSheet = oDrawDoc.ActiveSheet.Name

        i = 1

        oTitleBlock = oSheet.TitleBlock
        oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
        For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
            Select Case oTextBox.Text
                Case "<Scale>"
                    oPromptEntry = oTitleBlock.GetResultText(oTextBox)
            End Select
        Next

        If oPromptEntry = "<Scale>" Then
            oPromptText = "Scale from view"
        ElseIf oPromptEntry = "" Then
            oPromptText = "Scale from view"
        Else
            oPromptText = oPromptEntry
        End If

        Dim drawingDoc As DrawingDocument = TryCast(inventorApp.ActiveDocument, DrawingDocument)
        dwgScale = InputBox("If you leave as 'Scale from view' then it will use base view scale, otherwise enter scale to show", "Sheet Scale", oPromptText)

        For Each viewX As DrawingView In oSheet.DrawingViews
            If (Not String.IsNullOrEmpty(viewX.ScaleString)) Then
                If dwgScale = "Scale from view" Then
                    scaleString = viewX.ScaleString
                Else
                    scaleString = dwgScale

                    Exit For
                End If

            End If
        Next

        oTitleBlock.SetPromptResultText(scaleTextBox, scaleString)
        UpdateStatusBar("Drawing scale set")
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

    Function GetAuthorTextBox(ByVal titleDef As TitleBlockDefinition) As Inventor.TextBox
        For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
            If (defText.Text = "<AUTHOR>" Or defText.Text = "AUTHOR") Then
                Return defText
            End If
        Next
        Return Nothing
    End Function

    Private Sub btDiaEng_Click(sender As Object, e As EventArgs) Handles btDiaEng.Click
        SendSymbol(tbEngineer, "Ø")
    End Sub

    Private Sub btDegEng_Click(sender As Object, e As EventArgs) Handles btDegEng.Click
        SendSymbol(tbEngineer, "°")
    End Sub

    Private Sub btDiaDes_Click(sender As Object, e As EventArgs) Handles btDiaDes.Click
        SendSymbol(tbDescription, "Ø")
    End Sub

    Private Sub btDegDes_Click(sender As Object, e As EventArgs) Handles btDegDes.Click
        SendSymbol(tbDescription, "°")
    End Sub

    Public Sub AttachRefFile(ActiveDoc As Document, RefFile As String)
        FileNameHere = System.IO.Path.GetFileName(RefFile)
        AttachFile = MsgBox(FileNameHere & " File exported, attach it to main file as reference?", vbYesNo, "File Attach")
        If AttachFile = vbYes Then
            AddReferences(ActiveDoc, RefFile)
            UpdateStatusBar("File attached")
        Else
            'Do Nothing
        End If
    End Sub

    Private Sub btExpStp_Click(sender As Object, e As EventArgs) Handles btExpStp.Click
        Dim oDocu As Document = Nothing
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
            If CheckRef = vbYes Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                'GetNewFilePaths()
                ' Get the STEP translator Add-In.
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                Dim oSTEPTranslator As TranslatorAddIn
                oSTEPTranslator = inventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

                If oSTEPTranslator Is Nothing Then
                    MsgBox("Could not access STEP translator.")
                    Exit Sub
                End If

                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext
                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap
                If oSTEPTranslator.HasSaveCopyAsOptions(inventorApp.ActiveDocument, oContext, oOptions) Then
                    ' Set application protocol.
                    ' 2 = AP 203 - Configuration Controlled Design
                    ' 3 = AP 214 - Automotive Design
                    oOptions.Value("ApplicationProtocolType") = 3

                    ' Other options...
                    'oOptions.Value("Author") = ""
                    'oOptions.Value("Authorization") = ""
                    'oOptions.Value("Description") = ""
                    'oOptions.Value("Organization") = ""

                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = NewPath + "_R" + oRev + ".stp"

                    Call oSTEPTranslator.SaveCopyAs(inventorApp.ActiveDocument, oContext, oOptions, oData)
                    UpdateStatusBar("File saved as Step file")

                    AttachRefFile(inventorApp.ActiveDocument, oData.FileName)
                    'AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
                    'If AttachFile = vbYes Then
                    '    AddReferences(inventorApp.ActiveDocument, oData.FileName)
                    '    UpdateStatusBar("File attached")
                    'Else
                    '    'Do Nothing
                    'End If
                End If
            Else
                'Do nothing
            End If
        ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            If Not iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                'GetNewFilePaths()
                Dim oRev = iProperties.GetorSetStandardiProperty(
                            RefDoc,
                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                    UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
                Else

                    ' Get the STEP translator Add-In.
                    Dim oSTEPTranslator As TranslatorAddIn
                    oSTEPTranslator = inventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

                    If oSTEPTranslator Is Nothing Then
                        MsgBox("Could not access STEP translator.")
                        Exit Sub
                    End If

                    Dim oContext As TranslationContext
                    oContext = inventorApp.TransientObjects.CreateTranslationContext
                    Dim oOptions As NameValueMap
                    oOptions = inventorApp.TransientObjects.CreateNameValueMap
                    If oSTEPTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
                        ' Set application protocol.
                        ' 2 = AP 203 - Configuration Controlled Design
                        ' 3 = AP 214 - Automotive Design
                        oOptions.Value("ApplicationProtocolType") = 3

                        ' Other options...
                        'oOptions.Value("Author") = ""
                        'oOptions.Value("Authorization") = ""
                        'oOptions.Value("Description") = ""
                        'oOptions.Value("Organization") = ""

                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                        Dim oData As DataMedium
                        oData = inventorApp.TransientObjects.CreateDataMedium
                        oData.FileName = RefNewPath + "_R" + oRev + ".stp"

                        Call oSTEPTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                        UpdateStatusBar("Part/Assy file saved as Step file")
                        AttachRefFile(RefDoc, oData.FileName)
                    End If
                End If
            Else
                CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
                If CheckRef = vbYes Then
                    oDocu = inventorApp.ActiveDocument
                    oDocu.Save2(True)
                    'GetNewFilePaths()
                    Dim oRev = iProperties.GetorSetStandardiProperty(
                                RefDoc,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
                    Else

                        ' Get the STEP translator Add-In.
                        Dim oSTEPTranslator As TranslatorAddIn
                        oSTEPTranslator = inventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

                        If oSTEPTranslator Is Nothing Then
                            MsgBox("Could not access STEP translator.")
                            Exit Sub
                        End If

                        Dim oContext As TranslationContext
                        oContext = inventorApp.TransientObjects.CreateTranslationContext
                        Dim oOptions As NameValueMap
                        oOptions = inventorApp.TransientObjects.CreateNameValueMap
                        If oSTEPTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
                            ' Set application protocol.
                            ' 2 = AP 203 - Configuration Controlled Design
                            ' 3 = AP 214 - Automotive Design
                            oOptions.Value("ApplicationProtocolType") = 3

                            ' Other options...
                            'oOptions.Value("Author") = ""
                            'oOptions.Value("Authorization") = ""
                            'oOptions.Value("Description") = ""
                            'oOptions.Value("Organization") = ""

                            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                            Dim oData As DataMedium
                            oData = inventorApp.TransientObjects.CreateDataMedium
                            oData.FileName = RefNewPath + "_R" + oRev + ".stp"

                            Call oSTEPTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                            UpdateStatusBar("Part/Assy file saved as Step file")
                            AttachRefFile(RefDoc, oData.FileName)
                        End If
                    End If
                Else
                    'Do Nothing
                End If
            End If

        End If
    End Sub

    Private Sub btExpStl_Click(sender As Object, e As EventArgs) Handles btExpStl.Click
        Dim oDocu As Document = Nothing
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
            If CheckRef = vbYes Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                ' Get the STL translator Add-In.
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                    inventorApp.ActiveDocument,
                                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                Dim oSTLTranslator As TranslatorAddIn
                oSTLTranslator = inventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
                If oSTLTranslator Is Nothing Then
                    MsgBox("Could not access STL translator.")
                    Exit Sub
                End If

                Dim oDoc = inventorApp.ActiveDocument

                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext

                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap

                'Save Copy As Options
                'Name value Map
                'ExportUnits = 4
                'Resolution = 1
                'AllowMoveMeshNode = False
                'SurfaceDeviation = 60
                'NormalDeviation = 14
                'MaxEdgeLength = 100
                'AspectRatio = 40
                'ExportFileStructure = 0
                'OutputFileType = 0
                'ExportColor = True

                If oSTLTranslator.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

                    oOptions.Value("ExportUnits") = 5
                    ' Set output file type:
                    '   0 - binary,  1 - ASCII
                    oOptions.Value("OutputFileType") = 0
                    ' Set accuracy.
                    '   2 = Low,  1 = Medium,  0 = High
                    oOptions.Value("Resolution") = 0
                    'oOptions.Value("SurfaceDeviation") = 0.005
                    'oOptions.Value("NormalDeviation") = 10
                    'oOptions.Value("MaxEdgeLength") = 100
                    'oOptions.Value("AspectRatio") = 21.5
                    oOptions.Value("ExportColor") = True

                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = NewPath + "_R" + oRev + ".stl"

                    Call oSTLTranslator.SaveCopyAs(oDoc, oContext, oOptions, oData)
                    UpdateStatusBar("File saved as STL file")
                    AttachRefFile(inventorApp.ActiveDocument, oData.FileName)

                    ' The various names and values for the settings is described below.
                    'ExportUnits
                    '     2 - Inch
                    '     3 - Foot
                    '     4 - Centimeter
                    '     5 - Millimeter
                    '     6 - Meter
                    '     7 - Micron
                    '
                    'AllowMoveMeshNode (True Or False)
                    '
                    'ExportColor (True or False)  (Only used for Binary output file type)
                    '
                    'OutputFileType
                    '     0 - Binary
                    '     1 - ASCII
                    'ExportFileStructure (Only used for assemblies)
                    '     0 - One file
                    '     1 - One file per instance.
                    '
                    'Resolution
                    '     0 - High
                    '     1 - Medium
                    '     2 - Low
                    '     3 - Custom
                    '
                    '*** The following are only used for “Custom” resolution
                    '
                    'SurfaceDeviation
                    '     0 to 100 for a percentage between values computed
                    '     based on the size of the model.
                    'NormalDeviation
                    '     0 to 40 to indicate values 1 to 41
                    'MaxEdgeLength
                    '     0 to 100 for a percentage between values computed
                    '     based on the size of the model.
                    'AspectRatio
                    '     0 to 40 for values between 1.5 to 21.5 in 0.5 increments
                End If
            Else
                'Do Nothing
            End If
        ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            If Not iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                'GetNewFilePaths()
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                    RefDoc,
                                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                    UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
                Else
                    ' Get the STL translator Add-In.
                    Dim oSTLTranslator As TranslatorAddIn
                    oSTLTranslator = inventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
                    If oSTLTranslator Is Nothing Then
                        MsgBox("Could not access STL translator.")
                        Exit Sub
                    End If

                    Dim oContext As TranslationContext
                    oContext = inventorApp.TransientObjects.CreateTranslationContext

                    Dim oOptions As NameValueMap
                    oOptions = inventorApp.TransientObjects.CreateNameValueMap

                    '    Save Copy As Options:
                    '       Name Value Map:
                    '               ExportUnits = 4
                    '               Resolution = 1
                    '               AllowMoveMeshNode = False
                    '               SurfaceDeviation = 60
                    '               NormalDeviation = 14
                    '               MaxEdgeLength = 100
                    '               AspectRatio = 40
                    '               ExportFileStructure = 0
                    '               OutputFileType = 0
                    '               ExportColor = True

                    If oSTLTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then

                        oOptions.Value("ExportUnits") = 5
                        ' Set output file type:
                        '   0 - binary,  1 - ASCII
                        oOptions.Value("OutputFileType") = 0
                        ' Set accuracy.
                        '   2 = Low,  1 = Medium,  0 = High
                        oOptions.Value("Resolution") = 0
                        'oOptions.Value("SurfaceDeviation") = 0.005
                        'oOptions.Value("NormalDeviation") = 10
                        'oOptions.Value("MaxEdgeLength") = 100
                        'oOptions.Value("AspectRatio") = 21.5
                        oOptions.Value("ExportColor") = True

                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                        Dim oData As DataMedium
                        oData = inventorApp.TransientObjects.CreateDataMedium
                        oData.FileName = RefNewPath + "_R" + oRev + ".stl"

                        Call oSTLTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                        UpdateStatusBar("Part/Assy file saved as STL file")
                        AttachRefFile(RefDoc, oData.FileName)
                    End If
                End If
            Else
                CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
                If CheckRef = vbYes Then
                    oDocu = inventorApp.ActiveDocument
                    oDocu.Save2(True)
                    'GetNewFilePaths()
                    Dim oRev = iProperties.GetorSetStandardiProperty(
                                    RefDoc,
                                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
                    Else
                        ' Get the STL translator Add-In.
                        Dim oSTLTranslator As TranslatorAddIn
                        oSTLTranslator = inventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
                        If oSTLTranslator Is Nothing Then
                            MsgBox("Could not access STL translator.")
                            Exit Sub
                        End If

                        Dim oContext As TranslationContext
                        oContext = inventorApp.TransientObjects.CreateTranslationContext

                        Dim oOptions As NameValueMap
                        oOptions = inventorApp.TransientObjects.CreateNameValueMap

                        '    Save Copy As Options:
                        '       Name Value Map:
                        '               ExportUnits = 4
                        '               Resolution = 1
                        '               AllowMoveMeshNode = False
                        '               SurfaceDeviation = 60
                        '               NormalDeviation = 14
                        '               MaxEdgeLength = 100
                        '               AspectRatio = 40
                        '               ExportFileStructure = 0
                        '               OutputFileType = 0
                        '               ExportColor = True

                        If oSTLTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then

                            oOptions.Value("ExportUnits") = 5
                            ' Set output file type:
                            '   0 - binary,  1 - ASCII
                            oOptions.Value("OutputFileType") = 0
                            ' Set accuracy.
                            '   2 = Low,  1 = Medium,  0 = High
                            oOptions.Value("Resolution") = 0
                            'oOptions.Value("SurfaceDeviation") = 0.005
                            'oOptions.Value("NormalDeviation") = 10
                            'oOptions.Value("MaxEdgeLength") = 100
                            'oOptions.Value("AspectRatio") = 21.5
                            oOptions.Value("ExportColor") = True

                            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                            Dim oData As DataMedium
                            oData = inventorApp.TransientObjects.CreateDataMedium
                            oData.FileName = RefNewPath + "_R" + oRev + ".stl"

                            Call oSTLTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                            UpdateStatusBar("Part/Assy file saved as STL file")
                            AttachRefFile(RefDoc, oData.FileName)
                        End If
                    End If
                Else
                    'Do Nothing
                End If
            End If
        End If
    End Sub

    Private Sub btExpPdf_Click(sender As Object, e As EventArgs) Handles btExpPdf.Click
        Dim oDocu As Document = Nothing
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
            If CheckRef = vbYes Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                'GetNewFilePaths()
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                If Not inventorApp.SoftwareVersion.Major > 20 Then
                    inventorApp.StatusBarText = inventorApp.SoftwareVersion.Major
                    MessageBox.Show("3D PDF export not available in Inventor versions < 2017 release!")
                    Exit Sub
                End If
                ' Get the 3D PDF Add-In.
                Dim oPDFAddIn As ApplicationAddIn
                Dim oAddin As ApplicationAddIn
                For Each oAddin In inventorApp.ApplicationAddIns
                    If oAddin.ClassIdString = "{3EE52B28-D6E0-4EA4-8AA6-C2A266DEBB88}" Then
                        oPDFAddIn = oAddin
                        Exit For
                    End If
                Next

                If oPDFAddIn Is Nothing Then
                    MsgBox("Inventor 3D PDF Addin not loaded.")
                    Exit Sub
                End If

                Dim oPDFConvertor3D = oPDFAddIn.Automation

                'Set a reference to the active document (the document to be published).
                Dim oDocument As Document = inventorApp.ActiveDocument

                If oDocument.FileSaveCounter = 0 Then
                    MsgBox("You must save the document to continue...")
                    Return
                End If

                ' Create a NameValueMap objectfor all options...
                Dim oOptions As NameValueMap = inventorApp.TransientObjects.CreateNameValueMap
                Dim STEPFileOptions As NameValueMap = inventorApp.TransientObjects.CreateNameValueMap

                ' All Possible Options
                ' Export file name and location...
                oOptions.Value("FileOutputLocation") = NewPath + "_R" + oRev + ".pdf"
                ' Export annotations?
                oOptions.Value("ExportAnnotations") = 1
                ' Export work features?
                oOptions.Value("ExportWokFeatures") = 1
                ' Attach STEP file to 3D PDF?
                oOptions.Value("GenerateAndAttachSTEPFile") = True
                ' What quality (high quality takes longer to export)
                'oOptions.Value("VisualizationQuality") = AccuracyEnumVeryHigh
                oOptions.Value("VisualizationQuality") = AccuracyEnum.kHigh
                'oOptions.Value("VisualizationQuality") = AccuracyEnum.kMedium
                'oOptions.Value("VisualizationQuality") = AccuracyEnum.kLow
                ' Limit export to entities in selected view representation(s)
                oOptions.Value("LimitToEntitiesInDVRs") = True
                ' Open the 3D PDF when export is complete?
                oOptions.Value("ViewPDFWhenFinished") = False

                ' Export all properties?
                oOptions.Value("ExportAllProperties") = True
                ' OR - Set the specific properties to export
                '    Dim sProps(5) As String
                '    sProps(0) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Title"
                '    sProps(1) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Keywords"
                '    sProps(2) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Comments"
                '    sProps(3) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Description"
                '    sProps(4) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Stock Number"
                '    sProps(5) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Revision Number"

                'oOptions.Value("ExportProperties") = sProps

                ' Choose the export template based off the current document type
                If oDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    oOptions.Value("ExportTemplate") = "C:\Users\Public\Documents\Autodesk\Inventor 2017\Templates\Sample Part Template.pdf"
                Else
                    oOptions.Value("ExportTemplate") = "C:\Users\Public\Documents\Autodesk\Inventor 2017\Templates\Sample Assembly Template.pdf"
                End If

                ' Define a file to attach to the exported 3D PDF - note here I have picked an Excel spreadsheet
                ' You need to use the full path and filename - if it does not exist the file will not be attached.
                Dim oAttachedFiles As String() = {"C:\FileToAttach.xlsx"}
                oOptions.Value("AttachedFiles") = oAttachedFiles

                ' Set the design view(s) to export - note here I am exporting only the active design view (view representation)
                Dim sDesignViews(0) As String
                sDesignViews(0) = oDocument.ComponentDefinition.RepresentationsManager.ActiveDesignViewRepresentation.Name
                oOptions.Value("ExportDesignViewRepresentations") = sDesignViews

                'Publish document.
                Call oPDFConvertor3D.Publish(oDocument, oOptions)
                UpdateStatusBar("File saved as 3D pdf file")
                AttachRefFile(inventorApp.ActiveDocument, oOptions.Value("FileOutputLocation"))
            Else
                'Do Nothing
            End If
        Else
            If Not iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                'GetNewFilePaths()
                ' Get the PDF translator Add-In.
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                    inventorApp.ActiveDocument,
                                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                Dim PDFAddIn As TranslatorAddIn
                PDFAddIn = inventorApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

                'Set a reference to the active document (the document to be published).
                Dim oDocument As Document
                oDocument = inventorApp.ActiveDocument

                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext
                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                ' Create a NameValueMap object
                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap

                ' Create a DataMedium object
                Dim oDataMedium As DataMedium
                oDataMedium = inventorApp.TransientObjects.CreateDataMedium

                ' Check whether the translator has 'SaveCopyAs' options
                If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

                    ' Options for drawings...

                    oOptions.Value("All_Color_AS_Black") = 1

                    'oOptions.Value("Remove_Line_Weights") = 0
                    'oOptions.Value("Vector_Resolution") = 400
                    'oOptions.Value("Sheet_Range") = kPrintAllSheets
                    'oOptions.Value("Custom_Begin_Sheet") = 2
                    'oOptions.Value("Custom_End_Sheet") = 4

                End If

                'Set the destination file name
                oDataMedium.FileName = NewPath + "_R" + oRev + ".pdf"

                'Publish document.
                Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
                UpdateStatusBar("File saved as pdf file")
                AttachRefFile(inventorApp.ActiveDocument, oDataMedium.FileName)
            Else
                CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
                If CheckRef = vbYes Then
                    oDocu = inventorApp.ActiveDocument
                    oDocu.Save2(True)
                    'GetNewFilePaths()
                    ' Get the PDF translator Add-In.
                    Dim oRev = iProperties.GetorSetStandardiProperty(
                                        inventorApp.ActiveDocument,
                                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                    Dim PDFAddIn As TranslatorAddIn
                    PDFAddIn = inventorApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

                    'Set a reference to the active document (the document to be published).
                    Dim oDocument As Document
                    oDocument = inventorApp.ActiveDocument

                    Dim oContext As TranslationContext
                    oContext = inventorApp.TransientObjects.CreateTranslationContext
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    ' Create a NameValueMap object
                    Dim oOptions As NameValueMap
                    oOptions = inventorApp.TransientObjects.CreateNameValueMap

                    ' Create a DataMedium object
                    Dim oDataMedium As DataMedium
                    oDataMedium = inventorApp.TransientObjects.CreateDataMedium

                    ' Check whether the translator has 'SaveCopyAs' options
                    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

                        ' Options for drawings...

                        oOptions.Value("All_Color_AS_Black") = 1

                        'oOptions.Value("Remove_Line_Weights") = 0
                        'oOptions.Value("Vector_Resolution") = 400
                        'oOptions.Value("Sheet_Range") = kPrintAllSheets
                        'oOptions.Value("Custom_Begin_Sheet") = 2
                        'oOptions.Value("Custom_End_Sheet") = 4

                    End If

                    'Set the destination file name
                    oDataMedium.FileName = NewPath + "_R" + oRev + ".pdf"

                    'Publish document.
                    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
                    UpdateStatusBar("File saved as pdf file")
                    AttachRefFile(inventorApp.ActiveDocument, oDataMedium.FileName)
                Else
                    'Do Nothing
                End If
            End If
        End If
    End Sub

    Private Sub tbPartNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles tbPartNumber.KeyUp

        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbDescription.Focus()
                    assydoc.SelectSet.Select(compOcc)
                ElseIf e.KeyValue = Keys.Return Then
                    tbPartNumber_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbDescription.Focus()

                ElseIf e.KeyValue = Keys.Return Then
                    tbPartNumber_Leave(sender, e)
                End If
            End If
        Else
            If e.KeyValue = Keys.Tab Then
                tbDescription.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbPartNumber_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub tbStockNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles tbStockNumber.KeyUp
        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbEngineer.Focus()
                    assydoc.SelectSet.Select(compOcc)

                ElseIf e.KeyValue = Keys.Return Then
                    tbStockNumber_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbEngineer.Focus()

                ElseIf e.KeyValue = Keys.Return Then
                    tbStockNumber_Leave(sender, e)
                End If
            End If
        Else
            If e.KeyValue = Keys.Tab Then
                tbEngineer.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbStockNumber_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub tbEngineer_KeyUp(sender As Object, e As KeyEventArgs) Handles tbEngineer.KeyUp
        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbRevNo.Focus()
                    assydoc.SelectSet.Select(compOcc)
                ElseIf e.KeyValue = Keys.Return Then
                    tbEngineer_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbRevNo.Focus()
                ElseIf e.KeyValue = Keys.Return Then
                    tbEngineer_Leave(sender, e)
                End If
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is PartDocument Then
            If e.KeyValue = Keys.Tab Then
                tbRevNo.Focus()
            ElseIf e.KeyValue = Keys.Return Then
                tbEngineer_Leave(sender, e)
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then
            If e.KeyValue = Keys.Tab Then
                tbDrawnBy.Focus()
            ElseIf e.KeyValue = Keys.Return Then
                tbEngineer_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub btUpdateAll_KeyUp(sender As Object, e As KeyEventArgs) Handles btUpdateAll.KeyUp
        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Or Keys.Return Then
                    btUpdateAll_Click(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Or Keys.Return Then
                    btUpdateAll_Click(sender, e)
                End If
            End If
        Else
            If e.KeyValue = Keys.Tab Or Keys.Return Then
                btUpdateAll_Click(sender, e)
            End If
        End If

    End Sub

    Private Sub FileLocation_Click(sender As Object, e As EventArgs) Handles FileLocation.Click
        If inventorApp.ActiveEditObject IsNot Nothing Then
            Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
            Process.Start("explorer.exe", directoryPath)
        Else
            Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)
            Process.Start("explorer.exe", directoryPath)
        End If
    End Sub

    Private Sub FileLocation_MouseHover(sender As Object, e As EventArgs) Handles FileLocation.MouseHover
        FileLocation.ForeColor = Drawing.Color.Blue
    End Sub

    Private Sub FileLocation_MouseLeave(sender As Object, e As EventArgs) Handles FileLocation.MouseLeave
        FileLocation.ForeColor = Drawing.Color.Black
    End Sub

    Private Sub tbPartNumber_TextChanged(sender As Object, e As EventArgs) Handles tbPartNumber.TextChanged
        tbPartNumber.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbStockNumber_TextChanged(sender As Object, e As EventArgs) Handles tbStockNumber.TextChanged
        tbStockNumber.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbEngineer_TextChanged(sender As Object, e As EventArgs) Handles tbEngineer.TextChanged
        tbEngineer.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbDrawnBy_TextChanged(sender As Object, e As EventArgs) Handles tbDrawnBy.TextChanged
        tbDrawnBy.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbDescription_TextChanged(sender As Object, e As EventArgs) Handles tbDescription.TextChanged
        tbDescription.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbComments_TextChanged(sender As Object, e As EventArgs) Handles tbComments.TextChanged
        tbComments.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbNotes_TextChanged(sender As Object, e As EventArgs) Handles tbNotes.TextChanged
        tbNotes.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbDescription_Leave(sender As Object, e As EventArgs) Handles tbDescription.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            tbDescription.ForeColor = Drawing.Color.Black
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "Description", tbDescription.Text)

            Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
            Dim oSht As Sheet = oDWG.ActiveSheet
            Dim oView As DrawingView = Nothing

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next
            Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument

            Dim drawingDesc As String = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
            iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, drawingDesc, "")
        End If
    End Sub

    Private Sub btExpSat_Click(sender As Object, e As EventArgs) Handles btExpSat.Click
        Dim oDocu As Document = Nothing
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
            If CheckRef = vbYes Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                ' Get the SAT translator Add-In.
                Dim oSATTrans As TranslatorAddIn
                oSATTrans = inventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
                If oSATTrans Is Nothing Then
                    MsgBox("Could not access SAT translator.")
                    Exit Sub
                End If
                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext
                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap
                If oSATTrans.HasSaveCopyAsOptions(inventorApp.ActiveDocument, oContext, oOptions) Then
                    oOptions.Value("ExportUnits") = 5
                    oOptions.Value("IncludeSketches") = 0
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = NewPath + "_R" + oRev + ".sat"
                    Call oSATTrans.SaveCopyAs(inventorApp.ActiveDocument, oContext, oOptions, oData)
                    UpdateStatusBar("File saved as Sat file")
                    AttachRefFile(inventorApp.ActiveDocument, oData.FileName)
                End If
            Else
                'Do Nothing
            End If
        ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            If Not iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
                oDocu = inventorApp.ActiveDocument
                oDocu.Save2(True)
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                RefDoc,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                ' Get the SAT translator Add-In.
                Dim oSATTrans As TranslatorAddIn
                oSATTrans = inventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
                If oSATTrans Is Nothing Then
                    MsgBox("Could not access SAT translator.")
                    Exit Sub
                End If
                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext
                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap
                If oSATTrans.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
                    oOptions.Value("ExportUnits") = 5
                    oOptions.Value("IncludeSketches") = 0
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = RefNewPath + "_R" + oRev + ".sat"
                    Call oSATTrans.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                    UpdateStatusBar("Part/Assy file saved as Sat file")
                    AttachRefFile(RefDoc, oData.FileName)
                End If
            Else
                CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
                If CheckRef = vbYes Then
                    oDocu = inventorApp.ActiveDocument
                    oDocu.Save2(True)
                    Dim oRev = iProperties.GetorSetStandardiProperty(
                                    RefDoc,
                                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                    ' Get the SAT translator Add-In.
                    Dim oSATTrans As TranslatorAddIn
                    oSATTrans = inventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
                    If oSATTrans Is Nothing Then
                        MsgBox("Could not access SAT translator.")
                        Exit Sub
                    End If
                    Dim oContext As TranslationContext
                    oContext = inventorApp.TransientObjects.CreateTranslationContext
                    Dim oOptions As NameValueMap
                    oOptions = inventorApp.TransientObjects.CreateNameValueMap
                    If oSATTrans.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
                        oOptions.Value("ExportUnits") = 5
                        oOptions.Value("IncludeSketches") = 0
                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
                        Dim oData As DataMedium
                        oData = inventorApp.TransientObjects.CreateDataMedium
                        oData.FileName = RefNewPath + "_R" + oRev + ".sat"
                        Call oSATTrans.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                        UpdateStatusBar("Part/Assy file saved as Sat file")
                        AttachRefFile(RefDoc, oData.FileName)
                    End If
                Else
                    'Do Nothing
                End If
            End If
        End If

    End Sub

    Private Sub tbEngineer_Enter(sender As Object, e As EventArgs) Handles tbEngineer.Enter
        If tbEngineer.Text = "Engineer" Then
            tbEngineer.Clear()
            tbEngineer.Focus()
        End If
    End Sub

    Private Sub tbStockNumber_Enter(sender As Object, e As EventArgs) Handles tbStockNumber.Enter
        If tbStockNumber.Text = "Stock Number" Then
            tbStockNumber.Clear()
            tbStockNumber.Focus()
        End If
    End Sub

    Private Sub tbDescription_Enter(sender As Object, e As EventArgs) Handles tbDescription.Enter
        If tbDescription.Text = "Description" Then
            tbDescription.Clear()
            tbDescription.Focus()
        End If
    End Sub

    Private Sub tbPartNumber_Enter(sender As Object, e As EventArgs) Handles tbPartNumber.Enter
        If tbPartNumber.Text = "Part Number" Then
            tbPartNumber.Clear()
            tbPartNumber.Focus()
        End If
    End Sub

    Private Sub ModelFileLocation_MouseHover(sender As Object, e As EventArgs) Handles ModelFileLocation.MouseHover
        ModelFileLocation.ForeColor = Drawing.Color.Blue
    End Sub

    Private Sub ModelFileLocation_Click(sender As Object, e As EventArgs) Handles ModelFileLocation.Click
        Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

        Dim oSht As Sheet = oDWG.ActiveSheet

        Dim oView As DrawingView = Nothing
        Dim drawnDoc As Document = Nothing

        For Each view As DrawingView In oSht.DrawingViews
            oView = view
            Exit For
        Next

        ModelPath = System.IO.Path.GetDirectoryName(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
        Process.Start("explorer.exe", ModelPath)

        ModelPath = Nothing
    End Sub

    Private Sub ModelFileLocation_MouseLeave(sender As Object, e As EventArgs) Handles ModelFileLocation.MouseLeave
        ModelFileLocation.ForeColor = Drawing.Color.Black
    End Sub

    Private Sub tbRevNo_Leave(sender As Object, e As EventArgs) Handles tbRevNo.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then
                Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing

                For Each view As DrawingView In oSht.DrawingViews
                    oView = view
                    Exit For
                Next
                Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument
                tbRevNo.ForeColor = Drawing.Color.Black

                Dim drawingRev As String =
                    iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, tbRevNo.Text, "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Revision Updated to: " + drawingRev)
                UpdateStatusBar("Revision updated to " + drawingRev)


                Dim modelrev = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                If Not modelrev = drawingRev Then
                    modelrev = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, drawingRev, "")
                End If
            Else
                tbRevNo.ForeColor = Drawing.Color.Black

                Dim iPropRev As String =
                    iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, tbRevNo.Text, "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Revision Updated to: " + iPropRev)
                UpdateStatusBar("Revision updated to " + iPropRev)
            End If
        End If
    End Sub

    Private Sub tbRevNo_KeyUp(sender As Object, e As KeyEventArgs) Handles tbRevNo.KeyUp
        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbRevNo_Leave(sender, e)
                    btUpdateAll.Focus()
                    assydoc.SelectSet.Select(compOcc)

                ElseIf e.KeyValue = Keys.Return Then
                    tbRevNo_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbRevNo_Leave(sender, e)
                    btUpdateAll.Focus()

                ElseIf e.KeyValue = Keys.Return Then
                    tbRevNo_Leave(sender, e)
                End If
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is PartDocument Then
            If e.KeyValue = Keys.Tab Then
                tbRevNo_Leave(sender, e)
                btUpdateAll.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbRevNo_Leave(sender, e)
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then
            If e.KeyValue = Keys.Tab Then
                tbRevNo_Leave(sender, e)
                btUpdateAll.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbRevNo_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub tbRevNo_TextChanged(sender As Object, e As EventArgs) Handles tbRevNo.TextChanged
        tbRevNo.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbPartNumber_MouseClick(sender As Object, e As MouseEventArgs) Handles tbPartNumber.MouseClick
        If tbPartNumber.Text = "Part Number" Then
            tbPartNumber.Clear()
            tbPartNumber.Focus()
        End If
    End Sub

    Private Sub tbDescription_MouseClick(sender As Object, e As MouseEventArgs) Handles tbDescription.MouseClick
        If tbDescription.Text = "Description" Then
            tbDescription.Clear()
            tbDescription.Focus()
        End If
    End Sub

    Private Sub tbStockNumber_MouseClick(sender As Object, e As MouseEventArgs) Handles tbStockNumber.MouseClick
        If tbStockNumber.Text = "Stock Number" Then
            tbStockNumber.Clear()
            tbStockNumber.Focus()
        End If
    End Sub

    Private Sub tbEngineer_MouseClick(sender As Object, e As MouseEventArgs) Handles tbEngineer.MouseClick
        If tbEngineer.Text = "Engineer" Then
            tbEngineer.Clear()
            tbEngineer.Focus()
        End If
    End Sub

    Private Sub tbRevNo_Enter(sender As Object, e As EventArgs) Handles tbRevNo.Enter
        If tbRevNo.Text = "Revision Number" Then
            tbRevNo.Clear()
            tbRevNo.Focus()
        End If
    End Sub

    Private Sub tbRevNo_MouseClick(sender As Object, e As MouseEventArgs) Handles tbRevNo.MouseClick
        If tbRevNo.Text = "Revision Number" Then
            tbRevNo.Clear()
            tbRevNo.Focus()
        End If
    End Sub

    Private Sub btReNum_Click(sender As Object, e As EventArgs) Handles btReNum.Click
        If TypeOf AddinGlobal.InventorApp.ActiveDocument Is DrawingDocument Then
            Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

            Dim oSht As Sheet = oDWG.ActiveSheet

            Dim oView As DrawingView = Nothing
            Dim drawnDoc As Document = Nothing

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            Dim doc = drawnDoc
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            oBOM.StructuredViewEnabled = True

            Dim oStructuredBOMView As BOMView
            oStructuredBOMView = oBOM.BOMViews.Item("Structured")
            Call oStructuredBOMView.Renumber(1, 1)
            oSht.Update()
            Dim oPartsList As PartsList = oDWG.ActiveSheet.PartsLists.Item(1)
            oPartsList.Sort("ITEM", True)
        ElseIf TypeOf AddinGlobal.InventorApp.ActiveDocument Is AssemblyDocument Then
            Dim doc = AddinGlobal.InventorApp.ActiveDocument
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            oBOM.StructuredViewEnabled = True

            Dim oStructuredBOMView As BOMView
            oStructuredBOMView = oBOM.BOMViews.Item("Structured")
            Call oStructuredBOMView.Renumber(1, 1)
        End If
    End Sub

    Private Sub btCopyPN_Click(sender As Object, e As EventArgs) Handles btCopyPN.Click
        If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
            If tbPartNumber.Text.Length > 0 Then
                tbStockNumber.Text = tbPartNumber.Text
                tbStockNumber_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub btPipes_Click(sender As Object, e As EventArgs) Handles btPipes.Click
        'define the active document as an assembly file
        Dim oAsmDoc As AssemblyDocument
        oAsmDoc = inventorApp.ActiveDocument
        'oAsmName = oAsmDoc.FileName 'without extension

        oAsmName = System.IO.Path.GetFileNameWithoutExtension(oAsmDoc.FullDocumentName)

        If inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
            Exit Sub
        End If
        'get user input
        RUsure = MessageBox.Show(
        "This will create a STEP file for all components." _
        & vbLf & " " _
        & vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
        & vbLf & "This could take a while.", "Batch Output STEPs ", MessageBoxButtons.YesNo)
        If RUsure = vbNo Then
            Return
        Else
        End If
        '- - - - - - - - - - - - -STEP setup - - - - - - - - - - - -
        'oPath = oAsmDoc.Path
        ''get STEP target folder path
        'oFolder = oPath & "\" & oAsmName & " STEP Files"
        ''Check for the step folder and create it if it does not exist
        'If Not System.IO.Directory.Exists(oFolder) Then
        '    System.IO.Directory.CreateDirectory(oFolder)
        'End If


        ''- - - - - - - - - - - - -Assembly - - - - - - - - - - - -
        'oAsmDoc.Document.SaveAs(oFolder & "\" & oAsmName & (".stp"), True)

        '- - - - - - - - - - - - -Components - - - - - - - - - - - -
        'look at the files referenced by the assembly
        Dim oRefDocs As DocumentsEnumerator
        oRefDocs = oAsmDoc.AllReferencedDocuments
        Dim oRefDoc As Document
        Dim oRev = iProperties.GetorSetStandardiProperty(
                            inventorApp.ActiveDocument,
                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        'work the referenced models
        'oDocu = inventorApp.ActiveDocument
        oAsmDoc.Save2(True)
        For Each oRefDoc In oRefDocs
            If Not iPropertiesAddInServer.CheckReadOnly(oRefDoc) Then
                NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullDocumentName)

                'GetNewFilePaths()
                ' Get the STEP translator Add-In.

                Dim oSTEPTranslator As TranslatorAddIn
                oSTEPTranslator = inventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

                If oSTEPTranslator Is Nothing Then
                    MsgBox("Could not access STEP translator.")
                    Exit Sub
                End If

                Dim oContext As TranslationContext
                oContext = inventorApp.TransientObjects.CreateTranslationContext
                Dim oOptions As NameValueMap
                oOptions = inventorApp.TransientObjects.CreateNameValueMap
                If oSTEPTranslator.HasSaveCopyAsOptions(oRefDoc, oContext, oOptions) Then

                    ' Set application protocol.
                    ' 2 = AP 203 - Configuration Controlled Design
                    ' 3 = AP 214 - Automotive Design
                    oOptions.Value("ApplicationProtocolType") = 4

                    ' Other options...
                    'oOptions.Value("Author") = ""
                    'oOptions.Value("Authorization") = ""
                    'oOptions.Value("Description") = ""
                    'oOptions.Value("Organization") = ""

                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = NewPath + "_R" + oRev + ".stp"

                    Call oSTEPTranslator.SaveCopyAs(oRefDoc, oContext, oOptions, oData)
                    UpdateStatusBar("File saved as Step file")

                    AttachRefFile(oAsmDoc, oData.FileName)
                    'AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
                    'If AttachFile = vbYes Then
                    '    AddReferences(inventorApp.ActiveDocument, oData.FileName)
                    '    UpdateStatusBar("File attached")
                    'Else
                    '    'Do Nothing
                    'End If
                End If
            Else
                UpdateStatusBar("File skipped because it's read-only")
            End If
        Next

    End Sub

    Private Sub btCheckIn_Click(sender As Object, e As EventArgs) Handles btCheckIn.Click
        If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
            If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                Dim AssyDoc As AssemblyDocument = inventorApp.ActiveDocument
                If AssyDoc.SelectSet.Count = 1 Then
                    If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                        ' Get the CommandManager object. 
                        Dim oCommandMgr As CommandManager
                        oCommandMgr = inventorApp.CommandManager

                        ' Get control definition for the line command. 
                        Dim oControlDef As ControlDefinition
                        oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckin")
                        ' Execute the command. 
                        Call oControlDef.Execute()
                    End If
                Else
                    ' Get the CommandManager object. 
                    Dim oCommandMgr As CommandManager
                    oCommandMgr = inventorApp.CommandManager

                    ' Get control definition for the line command. 
                    Dim oControlDef As ControlDefinition
                    oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckinTop")
                    ' Execute the command. 
                    Call oControlDef.Execute()
                End If
            Else
                ' Get the CommandManager object. 
                Dim oCommandMgr As CommandManager
                oCommandMgr = inventorApp.CommandManager

                ' Get control definition for the line command. 
                Dim oControlDef As ControlDefinition
                oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckinTop")
                ' Execute the command. 
                Call oControlDef.Execute()
            End If
        End If
    End Sub

    Private Sub btCheckOut_Click(sender As Object, e As EventArgs) Handles btCheckOut.Click
        If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
            If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                Dim AssyDoc As AssemblyDocument = inventorApp.ActiveDocument
                If AssyDoc.SelectSet.Count = 1 Then
                    If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                        ' Get the CommandManager object. 
                        Dim oCommandMgr As CommandManager
                        oCommandMgr = inventorApp.CommandManager

                        ' Get control definition for the line command. 
                        Dim oControlDef As ControlDefinition
                        oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckout")
                        ' Execute the command. 
                        Call oControlDef.Execute()
                    End If
                Else
                    ' Get the CommandManager object. 
                    Dim oCommandMgr As CommandManager
                    oCommandMgr = inventorApp.CommandManager

                    ' Get control definition for the line command. 
                    Dim oControlDef As ControlDefinition
                    oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckoutTop")
                    ' Execute the command. 
                    Call oControlDef.Execute()
                End If
            Else
                ' Get the CommandManager object. 
                Dim oCommandMgr As CommandManager
                oCommandMgr = inventorApp.CommandManager

                ' Get control definition for the line command. 
                Dim oControlDef As ControlDefinition
                oControlDef = oCommandMgr.ControlDefinitions.Item("VaultCheckoutTop")
                ' Execute the command. 
                Call oControlDef.Execute()
            End If
        End If
    End Sub

    Private Sub btViewNames_Click(sender As Object, e As EventArgs) Handles btViewNames.Click
        Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
        Dim oSheet As Sheet = oDrawDoc.ActiveSheet
        Dim oSheets As Sheets = Nothing
        Dim oView As DrawingView = Nothing
        Dim oViews As DrawingViews = Nothing
        Dim oLabel As String = "DETAIL OF ITEM "
        Dim isoLabel As String = "ISOMETRIC VIEW"

        For Each oView In oSheet.DrawingViews
            If Not oView.ParentView Is Nothing Then
                'we're working on a child view and should get the parent view as an object
            Else
                If oView.IsFlatPatternView Then
                    oView.Name() = "FLAT PATTERN OF ITEM "
                    oView.ShowLabel() = True
                ElseIf oView.ReferencedFile.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    oView.Name() = oLabel
                    oView.ShowLabel() = True
                End If
            End If
        Next
    End Sub

    Private Sub tbPartNumber_Leave(sender As Object, e As EventArgs) Handles tbPartNumber.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            tbPartNumber.ForeColor = Drawing.Color.Black
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "Part Number", tbPartNumber.Text)

            'Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
            'Dim oSht As Sheet = oDWG.ActiveSheet
            'Dim oView As DrawingView = Nothing
            'Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument

            'For Each view As DrawingView In oSht.DrawingViews
            '    oView = view
            '    Exit For
            'Next

            'Dim drawingPN As String = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
            'iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, drawingPN, "")
        End If
    End Sub

    Private Sub btCopyPN_KeyUp(sender As Object, e As KeyEventArgs) Handles btCopyPN.KeyUp
        If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
            If e.KeyValue = Keys.Return Then
                If tbPartNumber.Text.Length > 0 Then
                    tbStockNumber.Text = tbPartNumber.Text
                    tbStockNumber_Leave(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub tbDescription_KeyUp(sender As Object, e As KeyEventArgs) Handles tbDescription.KeyUp
        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbDescription.SelectionStart = 0
                    tbStockNumber.Focus()
                    assydoc.SelectSet.Select(compOcc)

                ElseIf e.KeyValue = Keys.Return Then
                    tbDescription_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbDescription.SelectionStart = 0
                    tbStockNumber.Focus()

                ElseIf e.KeyValue = Keys.Return Then
                    tbDescription_Leave(sender, e)
                End If
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is PartDocument Then
            If e.KeyValue = Keys.Tab Then
                tbDescription.SelectionStart = 0
                tbStockNumber.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbDescription_Leave(sender, e)
            End If
        ElseIf TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then
            If e.KeyValue = Keys.Tab Then
                tbDescription.SelectionStart = 0
                tbEngineer.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbDescription_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub tbDrawnBy_KeyUp(sender As Object, e As KeyEventArgs) Handles tbDrawnBy.KeyUp
        If e.KeyValue = Keys.Tab Then
            tbRevNo.Focus()
        ElseIf e.KeyValue = Keys.Return Then
            tbDrawnBy_Leave(sender, e)
        End If
    End Sub

    Private Sub tbPartNumber_MouseHover(sender As Object, e As EventArgs) Handles tbPartNumber.MouseHover
        Dim partText As String = tbPartNumber.Text
        ToolTip1.Show(partText, tbPartNumber)
    End Sub

    Private Sub tbDescription_MouseHover(sender As Object, e As EventArgs) Handles tbDescription.MouseHover
        Dim descText As String = tbDescription.Text
        ToolTip1.Show(descText, tbDescription)
    End Sub

    Private Sub tbNotes_Leave(sender As Object, e As EventArgs) Handles tbNotes.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            iProperties.GetorSetStandardiProperty(
                     inventorApp.ActiveDocument,
                     PropertiesForDesignTrackingPropertiesEnum.kCatalogWebLinkDesignTrackingProperties, tbNotes.Text, "", True)
            tbNotes.ForeColor = Drawing.Color.Black
        End If
    End Sub

    Private Sub tbNotes_KeyUp(sender As Object, e As KeyEventArgs) Handles tbNotes.KeyUp
        If e.KeyValue = Keys.Return Then
            tbNotes_Leave(sender, e)
        End If
    End Sub

    Private Sub tbComments_Leave(sender As Object, e As EventArgs) Handles tbComments.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                  PropertiesForSummaryInformationEnum.kCommentsSummaryInformation,
                                                  tbComments.Text,
                                                  "",
                                                  True)
            tbComments.ForeColor = Drawing.Color.Black
        End If
    End Sub

    Private Sub tbComments_KeyUp(sender As Object, e As KeyEventArgs) Handles tbComments.KeyUp
        If e.KeyValue = Keys.Return Then
            tbComments_Leave(sender, e)
        End If
    End Sub

    Private Sub tbComments_Enter(sender As Object, e As EventArgs) Handles tbComments.Enter
        If tbComments.Text = "Comments" Then
            tbComments.Clear()
            tbComments.Focus()
        End If
        Dim hovText As String = "Comments"
        ToolTip1.Show(hovText, tbComments)
    End Sub

    Private Sub tbComments_MouseHover(sender As Object, e As EventArgs) Handles tbComments.MouseHover
        Dim hovText As String = "Comments"
        ToolTip1.Show(hovText, tbComments)
    End Sub

    Private Sub tbNotes_MouseHover(sender As Object, e As EventArgs) Handles tbNotes.MouseHover
        Dim hovText As String = "Notes"
        ToolTip1.Show(hovText, tbNotes)
    End Sub

    Private Sub tbNotes_Enter(sender As Object, e As EventArgs) Handles tbNotes.Enter
        Dim hovText As String = "Notes"
        ToolTip1.Show(hovText, tbNotes)
    End Sub

    'Public Shared Function GetFileName(path As String) As String
    'End Function
End Class
