Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor
Imports iPropertiesController.iPropertiesController
Imports log4net

Public Class IPropertiesForm
    'Inherits Form
    Private inventorApp As Inventor.Application
    'Private localWindow As DockableWindow
    Private value As String
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Public customMargin As Integer = 5
    Public customSize As Size = Me.ClientSize
    Public ReadOnly log As ILog = LogManager.GetLogger(GetType(IPropertiesForm))

    Public Sub GetNewFilePaths()
        If inventorApp.ActiveDocument IsNot Nothing Then
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
                        If oView IsNot Nothing Then
                            RefCurrentPath = System.IO.Path.GetDirectoryName(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                            NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)

                            RefDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                            RefNewPath = RefCurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub GetTheStuffs()
        tbPartNumber.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")

        tbDescription.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")

        tbRevNo.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        tbEngineer.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
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

    Public Sub New(ByVal inventorApp As Inventor.Application) ', ByVal addinCLS As String, ByRef localWindow As DockableWindow)
        Try
            log.Debug("Loading iProperties Form")
            InitializeComponent()

            'Me.KeyPreview = True
            Me.inventorApp = inventorApp
            Me.value = addinCLS
            'Me.localWindow = localWindow
            'Dim myDockableWindow As DockableWindow = uiMgr.DockableWindows.Add(addinCLS, "iPropertiesControllerWindow", "iProperties Controller " + addinName)
            'myDockableWindow.AddChild(Me.Handle)

            'If Not myDockableWindow.IsCustomized = True Then
            '    'myDockableWindow.DockingState = DockingStateEnum.kFloat
            '    myDockableWindow.DockingState = DockingStateEnum.kDockLastKnown
            'Else
            '    myDockableWindow.DockingState = DockingStateEnum.kFloat
            'End If

            'myDockableWindow.DisabledDockingStates = DockingStateEnum.kDockTop + DockingStateEnum.kDockBottom

            'Me.Dock = DockStyle.Fill
            'Me.Visible = True
            'localWindow = myDockableWindow
            'AddinGlobal.DockableList.Add(myDockableWindow)
        Catch ex As Exception
            log.Error(ex.Message)
        End Try
        log.Info("iProperties Form Loaded")

    End Sub

    Private Sub UpdateAllCommon()
        If inventorApp.ActiveDocument IsNot Nothing Then
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

            Dim drawnDoc As Document = Nothing

            drawnDoc = inventorApp.ActiveDocument

            If Not newPropValue = propname Then
                UpdateProperties(proptoUpdate, propname, newPropValue, iProp, drawnDoc)
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

    Private Sub CheckForDefaultAndUpdate(ByVal sumtoUpdate As PropertiesForSummaryInformationEnum, ByVal propname As String, ByVal newPropValue As String)
        Dim iProp As String = String.Empty


        If inventorApp.ActiveEditObject IsNot Nothing Then
            If Not newPropValue = propname Then
                UpdateProperties(sumtoUpdate, propname, newPropValue, iProp)
            End If
        ElseIf inventorApp.ActiveDocument IsNot Nothing Then
            If Not newPropValue = propname Then
                UpdateProperties(sumtoUpdate, propname, newPropValue, iProp)
            End If
        End If


    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String, drawnDoc As Document)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(drawnDoc, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(drawnDoc, proptoUpdate, newPropValue, "", True)
            'inventorApp.ActiveDocument.Save2(True)
            iPropertiesAddInServer.UpdateDisplayediProperties()
            log.Debug(inventorApp.ActiveDocument.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject, proptoUpdate, newPropValue, "", True)
            'inventorApp.ActiveDocument.Save2(True)
            iPropertiesAddInServer.UpdateDisplayediProperties(inventorApp.ActiveEditObject)
            log.Debug(inventorApp.ActiveEditObject.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub UpdateProperties(proptoUpdate As PropertiesForDesignTrackingPropertiesEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String, AssyDoc As AssemblyDocument, selecteddoc As Document)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(selecteddoc, proptoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(selecteddoc, proptoUpdate, newPropValue, "", True)
            'inventorApp.ActiveDocument.Save2(True)
            iPropertiesAddInServer.UpdateDisplayediProperties(selecteddoc)
            log.Debug(selecteddoc.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
            iPropertiesAddInServer.ShowOccurrenceProperties(AssyDoc)
        End If
    End Sub

    Private Sub UpdateProperties(sumtoUpdate As PropertiesForSummaryInformationEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, sumtoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, sumtoUpdate, newPropValue, "", True)
            'inventorApp.ActiveDocument.Save2(True)
            iPropertiesAddInServer.UpdateDisplayediProperties()
            log.Debug(inventorApp.ActiveDocument.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub UpdateProperties(sumtoUpdate As PropertiesForSummaryInformationEnum, propname As String, ByRef newPropValue As String, ByRef iProp As String, drawnDoc As Document)
        If Not newPropValue = iProperties.GetorSetStandardiProperty(drawnDoc, sumtoUpdate, "", "") Then
            iProp = iProperties.GetorSetStandardiProperty(drawnDoc, sumtoUpdate, newPropValue, "", True)
            'inventorApp.ActiveDocument.Save2(True)
            iPropertiesAddInServer.UpdateDisplayediProperties()
            log.Debug(inventorApp.ActiveDocument.FullFileName + propname + " Updated to: " + iProp)
            UpdateStatusBar(propname + " updated to " + iProp)
        End If
    End Sub

    Private Sub SendSymbol(ByVal textbox As Object, symbol As String)
        Dim insertText = symbol

        Dim insertPos As Integer = textbox.SelectionStart

        textbox.Text = textbox.Text.Insert(insertPos, insertText)
        textbox.Focus()
        textbox.SelectionStart = insertPos + insertText.Length
    End Sub

    Private Sub tbStockNumber_Leave(sender As Object, e As EventArgs) Handles tbStockNumber.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            tbStockNumber.ForeColor = AddinGlobal.ForeColour
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "Stock Number", tbStockNumber.Text)
        End If
    End Sub

    Private Sub tbEngineer_Leave(sender As Object, e As EventArgs) Handles tbEngineer.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            tbEngineer.ForeColor = AddinGlobal.ForeColour
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "Engineer", tbEngineer.Text)

            If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing

                For Each view As DrawingView In oSht.DrawingViews
                    oView = view
                    Exit For
                Next
                Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument

                Dim drawingEng As String = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                Dim iProp As String = String.Empty
                UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "Engineer", drawingEng, iProp, drawnDoc)
            End If
        End If
    End Sub

    Private Sub btUpdateAll_Click(sender As Object, e As EventArgs) Handles btUpdateAll.Click

        If inventorApp.ActiveDocument IsNot Nothing Then
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
        If inventorApp.ActiveDocument IsNot Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If TypeOf inventorApp.ActiveDocument Is DrawingDocument Then
                    Dim oSheet As Sheet = inventorApp.ActiveDocument.ActiveSheet
                    Dim oSheets = inventorApp.ActiveDocument.Sheets
                    inventorApp.ActiveDocument.Activate()
                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = False
                        'DrawingSettings.DeferUpdates = False
                        btDefer.BackColor = Drawing.Color.Green
                        btDefer.Text = "Drawing Updates Not Deferred"
                        UpdateStatusBar("Drawing updates are no longer deferred")
                    ElseIf iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                        inventorApp.ActiveDocument.DrawingSettings.DeferUpdates = True
                        btDefer.BackColor = Drawing.Color.Red
                        btDefer.Text = "Drawing Updates Deferred"
                        UpdateStatusBar("Drawing updates are now deferred")
                    End If
                End If
            Else
                btDefer.BackColor = Drawing.Color.Green
                btDefer.Text = "Drawing Updates Not Deferred"
                MessageBox.Show("Save file before deferring updates")
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
        If inventorApp.ActiveDocument IsNot Nothing Then

            tbDrawnBy.ForeColor = AddinGlobal.ForeColour

            CheckForDefaultAndUpdate(PropertiesForSummaryInformationEnum.kAuthorSummaryInformation, "", tbDrawnBy.Text)

        End If
    End Sub

    Private Sub btITEM_Click(sender As Object, e As EventArgs) Handles btITEM.Click

        'Dim oAsmDoc As AssemblyDocument
        'oAsmDoc = inventorApp.ActiveDocument

        'oAsmName = System.IO.Path.GetFileNameWithoutExtension(oAsmDoc.FullDocumentName)

        'Dim oCommandMgr As CommandManager
        'oCommandMgr = inventorApp.CommandManager

        ''Dim compOcc As ComponentOccurrence = oRefDoc
        ''Call inventorApp.ActiveDocument.SelectSet.Select(compOcc)

        'tube = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Pick Part")
        'oAsmDoc.SelectSet.Select(tube)

        'Call oCommandMgr.ControlDefinitions.Item("CADC:TpXmlReporter:BendingMachineCmd").Execute()


        If TypeOf AddinGlobal.InventorApp.ActiveDocument Is DrawingDocument Then
            Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

            Dim oSht As Sheet = oDWG.ActiveSheet

            Dim oView As DrawingView = Nothing
            Dim drDoc As Document = Nothing

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            Dim doc = drDoc
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

                    Dim iProp As String = String.Empty
                    Dim DrawnDoc = oCompDef.Document
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, "Authority", item, iProp, DrawnDoc)
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

                    Dim iProp As String = String.Empty
                    Dim DrawnDoc = oCompDef.Document
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, "Authority", item, iProp, DrawnDoc)
                End If
            Next
        End If
        UpdateStatusBar("BOM item numbers copied to #ITEM")
    End Sub

    Private Sub tbMass_Enter(sender As Object, e As EventArgs) Handles tbMass.Enter
        If Not tbMass.Text.Length = 0 Then
            Clipboard.SetText(tbMass.Text)
            UpdateStatusBar("Mass copied to clipboard")
        End If
    End Sub

    Private Sub tbMass_MouseClick(sender As Object, e As MouseEventArgs) Handles tbMass.MouseClick
        tbMass_Enter(sender, e)
    End Sub

    Private Sub tbDensity_Enter(sender As Object, e As EventArgs) Handles tbDensity.Enter
        If Not tbDensity.Text.Length = 0 Then
            Clipboard.SetText(tbDensity.Text)
            UpdateStatusBar("Density copied to ")
        End If
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
            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next
            Dim MaterialString As String = String.Empty
            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

            i = 1

            oTitleBlock = oSheet.TitleBlock
            oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
            For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
                Select Case oTextBox.Text
                    Case "<Material>"
                        oPromptEntry = oTitleBlock.GetResultText(oTextBox)
                End Select
            Next
            MaterialTextBox = GetMaterialTextBox(oTitleBlock.Definition)

            If TypeOf drawnDoc Is AssemblyDocument Then
                If oPromptEntry = "<Material>" Then
                    oPromptText = "SEE ABOVE"
                ElseIf oPromptEntry = "" Then
                    oPromptText = "SEE ABOVE"
                Else
                    oPromptText = oPromptEntry
                End If
                prtMaterial = InputBox("leaving as 'SEE ABOVE' will fill box with 'SEE ABOVE'" &
                                   vbCrLf & "otherwise you can alter this to suit needs", "Assembly", oPromptText)
                If prtMaterial = "SEE ABOVE" Then
                    MaterialString = "SEE ABOVE"
                Else
                    MaterialString = prtMaterial
                End If
            Else
                If oPromptEntry = "<Material>" Then
                    oPromptText = "Engineer"
                ElseIf oPromptEntry = "" Then
                    oPromptText = "Engineer"
                Else
                    oPromptText = oPromptEntry
                End If
                prtMaterial = InputBox("leaving as 'Engineer' will bring through Engineer info from part, " &
                                  vbCrLf & "'PRT'or 'prt' will use part material, otherwise enter desired material info", "Material", oPromptText)
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
            End If

            'MaterialString = prtMaterial

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

        For Each view As DrawingView In oSheet.DrawingViews
            oView = view
            Exit For
        Next

        If (Not String.IsNullOrEmpty(oView.ScaleString)) Then
            If dwgScale = "Scale from view" Then
                scaleString = oView.ScaleString
            Else
                scaleString = dwgScale
            End If

        End If

        'For Each viewX As DrawingView In oSheet.DrawingViews
        '    If (Not String.IsNullOrEmpty(viewX.ScaleString)) Then
        '        If dwgScale = "Scale from view" Then
        '            scaleString = viewX.ScaleString
        '        Else
        '            scaleString = dwgScale

        '            Exit For
        '        End If

        '    End If
        'Next

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
            If (defText.Text = "Scale" Or defText.Text = "<Scale>") Then
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
            If iPropertiesAddInServer.CheckReadOnly(ActiveDoc) Then
                MessageBox.Show("You can't attach things to read-only files! Your file has been exported but if you want it to be attached, check-out and try again.", "Warning", MessageBoxButtons.OK)
                Exit Sub
            End If
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
                If Not iPropertiesAddInServer.CheckReadOnly(oDocu) Then
                    oDocu.Save2(True)
                End If
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
                'Dim oRev = iProperties.GetorSetStandardiProperty(
                '            RefDoc,
                '            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
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
                If Not iPropertiesAddInServer.CheckReadOnly(oDocu) Then
                    oDocu.Save2(True)
                End If
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
                'Dim oRev = iProperties.GetorSetStandardiProperty(
                '            RefDoc,
                '            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                Dim oRev = iProperties.GetorSetStandardiProperty(
                                inventorApp.ActiveDocument,
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
                If Not iPropertiesAddInServer.CheckReadOnly(oDocu) Then
                    oDocu.Save2(True)
                End If
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

                If oPDFAddIn IsNot Nothing Then

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
                    MsgBox("Inventor 3D PDF Addin not loaded.")
                    Exit Sub
                End If
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

                'Get sheet names and set options depending on them
                Dim strSheetName As String
                Dim oSheet As Sheet = inventorApp.ActiveDocument.ActiveSheet
                Dim oDoc = inventorApp.ActiveDocument
                Dim oSheets = inventorApp.ActiveDocument.Sheets
                For Each oSheet In oDocu.Sheets
                    oSheet.Activate()

                    strSheetName = oSheet.Name
                    If strSheetName.Contains("Model") Then
                        Exit For
                    End If
                Next

                If strSheetName.Contains("Model") Then
                    oDoc.sheets.item("Sheet:1").Activate()
                    oDoc.Sheets.item("Model (AutoCAD)").ExcludeFromPrinting = True
                End If

                If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

                    ' Options for drawings...

                    oOptions.Value("All_Color_AS_Black") = 1

                    'oOptions.Value("Remove_Line_Weights") = 0
                    'oOptions.Value("Vector_Resolution") = 400
                    oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
                    'oOptions.Value("Custom_Begin_Sheet") = 2
                    'oOptions.Value("Custom_End_Sheet") = 4

                End If
                oDoc.sheets.item("Sheet:1").Activate()

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

                    'Get sheet names and set options depending on them
                    Dim strSheetName As String
                    Dim oSheet As Sheet = inventorApp.ActiveDocument.ActiveSheet
                    Dim oDoc = inventorApp.ActiveDocument
                    Dim oSheets = inventorApp.ActiveDocument.Sheets
                    For Each oSheet In oDocu.Sheets
                        oSheet.Activate()

                        strSheetName = oSheet.Name
                        If strSheetName.Contains("Model") Then
                            Exit For
                        End If
                    Next

                    If strSheetName.Contains("Model") Then
                        oDoc.sheets.item("Sheet:1").Activate()
                        oDoc.Sheets.item("Model (AutoCAD)").ExcludeFromPrinting = True
                    End If

                    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

                        ' Options for drawings...

                        oOptions.Value("All_Color_AS_Black") = 1

                        'oOptions.Value("Remove_Line_Weights") = 0
                        'oOptions.Value("Vector_Resolution") = 400
                        oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
                        'oOptions.Value("Custom_Begin_Sheet") = 2
                        'oOptions.Value("Custom_End_Sheet") = 4

                    End If
                    oDoc.sheets.item("Sheet:1").Activate()

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
        If AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName IsNot Nothing Then
            If AddinGlobal.InventorApp.ActiveEditObject IsNot Nothing Then
                If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                    Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                    If AssyDoc.SelectSet.Count = 1 Then
                        Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                        Dim def As ComponentDefinition
                        def = compOcc.Definition
                        selecteddoc = compOcc.Definition.Document
                        'Dim directoryPath As String = System.IO.Path.GetDirectoryName(selecteddoc.FullDocumentName)
                        'Process.Start("explorer.exe", directoryPath)
                        Dim Fpath As String = System.IO.Path.GetFullPath(selecteddoc.FulldocumentName)
                        Dim FilePath As String = "/select,""" & Fpath & """"
                        Process.Start("explorer.exe", FilePath)
                    Else
                        ' Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                        'Process.Start("explorer.exe", directoryPath)
                        Dim Fpath As String = System.IO.Path.GetFullPath(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                        Dim FilePath As String = "/select,""" & Fpath & """"
                        Process.Start("explorer.exe", FilePath)
                    End If
                Else
                    ' Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                    'Process.Start("explorer.exe", directoryPath)
                    Dim Fpath As String = System.IO.Path.GetFullPath(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                    Dim FilePath As String = "/select,""" & Fpath & """"
                    Process.Start("explorer.exe", FilePath)
                End If
            Else
                If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                    If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                        Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                        If AssyDoc.SelectSet.Count = 1 Then
                            Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                            Dim def As ComponentDefinition
                            def = compOcc.Definition
                            selecteddoc = compOcc.Definition.Document
                            'Dim directoryPath As String = System.IO.Path.GetDirectoryName(selecteddoc.FullDocumentName)
                            Dim Fpath As String = System.IO.Path.GetFullPath(selecteddoc.FulldocumentName)
                            Dim FilePath As String = "/select,""" & Fpath & """"
                            Process.Start("explorer.exe", FilePath)
                        Else
                            Dim Fpath As String = System.IO.Path.GetFullPath(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                            Dim FilePath As String = "/select,""" & Fpath & """"
                            Process.Start("explorer.exe", FilePath)
                            'Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                            'Process.Start("explorer.exe", directoryPath)
                        End If
                    Else
                        Dim Fpath As String = System.IO.Path.GetFullPath(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                        Dim FilePath As String = "/select,""" & Fpath & """"
                        Process.Start("explorer.exe", FilePath)
                        'Dim directoryPath As String = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName)
                        'Process.Start("explorer.exe", directoryPath)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub FileLocation_MouseHover(sender As Object, e As EventArgs) Handles FileLocation.MouseHover
        FileLocation.ForeColor = Drawing.Color.Blue
    End Sub

    Private Sub FileLocation_MouseLeave(sender As Object, e As EventArgs) Handles FileLocation.MouseLeave
        FileLocation.ForeColor = AddinGlobal.ForeColour
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
    Private Sub tbService_TextChanged(sender As Object, e As EventArgs) Handles tbService.TextChanged
        tbNotes.ForeColor = Drawing.Color.Red
    End Sub

    Private Sub tbDescription_Leave(sender As Object, e As EventArgs) Handles tbDescription.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            tbDescription.ForeColor = AddinGlobal.ForeColour
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "Description", tbDescription.Text)
            If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Dim iProp As String = String.Empty

                Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing

                For Each view As DrawingView In oSht.DrawingViews
                    oView = view
                    Exit For
                Next
                If oView IsNot Nothing Then
                    Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument
                    Dim drawingDoc As Document = inventorApp.ActiveDocument
                    Dim drawingDesc As String = tbDescription.Text

                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "Description", drawingDesc, iProp, drawingDoc)
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "Description", drawingDesc, iProp, drawnDoc)
                End If
            Else
                'CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "Description", tbDescription.Text)
            End If
        End If
    End Sub

    Private Sub tbService_Leave(sender As Object, e As EventArgs) Handles tbService.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            tbDescription.ForeColor = AddinGlobal.ForeColour
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kProjectDesignTrackingProperties, "Project", tbService.Text)
        End If
    End Sub

    Private Sub tbService_KeyUp(sender As Object, e As KeyEventArgs) Handles tbService.KeyUp

        If TypeOf (inventorApp.ActiveDocument) Is AssemblyDocument Then
            Dim assydoc As Document = Nothing
            assydoc = inventorApp.ActiveDocument
            If assydoc.SelectSet.Count = 1 Then
                Dim compOcc As ComponentOccurrence = assydoc.SelectSet(1)
                If e.KeyValue = Keys.Tab Then
                    tbRevNo.Focus()
                    assydoc.SelectSet.Select(compOcc)
                ElseIf e.KeyValue = Keys.Return Then
                    tbService_Leave(sender, e)
                    assydoc.SelectSet.Select(compOcc)
                End If
            Else
                If e.KeyValue = Keys.Tab Then
                    tbRevNo.Focus()

                ElseIf e.KeyValue = Keys.Return Then
                    tbService_Leave(sender, e)
                End If
            End If
        Else
            If e.KeyValue = Keys.Tab Then
                tbRevNo.Focus()

            ElseIf e.KeyValue = Keys.Return Then
                tbService_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub tbService_Enter(sender As Object, e As EventArgs) Handles tbService.Enter
        If tbService.Text = "Project" Then
            tbService.Clear()
            tbService.Focus()
        End If
    End Sub

    Private Sub tbService_MouseClick(sender As Object, e As EventArgs) Handles tbService.MouseClick
        If tbService.Text = "Project" Then
            tbService.Clear()
            tbService.Focus()
        End If
    End Sub

    Private Sub tbService_MouseHover(sender As Object, e As EventArgs) Handles tbService.MouseHover
        Dim descText As String = tbService.Text
        ToolTip1.Show(descText, tbService)
    End Sub

    Private Sub btExpSat_Click(sender As Object, e As EventArgs)
        Dim oDocu As Document = Nothing
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
            If CheckRef = vbYes Then
                oDocu = inventorApp.ActiveDocument
                If Not iPropertiesAddInServer.CheckReadOnly(oDocu) Then
                    oDocu.Save2(True)
                End If
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
                'Dim oRev = iProperties.GetorSetStandardiProperty(
                '            RefDoc,
                '            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
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
        If AddinGlobal.InventorApp.ActiveEditDocument.FullDocumentName IsNot Nothing Then
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
        End If
    End Sub

    Private Sub ModelFileLocation_MouseLeave(sender As Object, e As EventArgs) Handles ModelFileLocation.MouseLeave
        ModelFileLocation.ForeColor = AddinGlobal.ForeColour
    End Sub

    Private Sub tbRevNo_Leave(sender As Object, e As EventArgs) Handles tbRevNo.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            If TypeOf (inventorApp.ActiveDocument) Is DrawingDocument Then
                Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing
                Dim drawnDoc As Document = Nothing

                For Each view As DrawingView In oSht.DrawingViews
                    oView = view
                    Exit For
                Next

                drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                tbRevNo.ForeColor = AddinGlobal.ForeColour
                Dim drawingRev As String = tbRevNo.Text

                Dim iProp As String = String.Empty
                UpdateProperties(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "Revision", drawingRev, iProp, drawnDoc)
                UpdateProperties(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "Revision", drawingRev, iProp)
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Revision Updated to: " + drawingRev)
                UpdateStatusBar("Revision updated to " + drawingRev)

                'iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, drawingRev, "", True)

            Else
                tbRevNo.ForeColor = AddinGlobal.ForeColour

                Dim iPropRev As String = tbRevNo.Text

                Dim iProp As String = String.Empty
                UpdateProperties(PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "Revision", iPropRev, iProp)


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
        If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
            If tbPartNumber.Text.Length > 0 Then
                tbStockNumber.Text = tbPartNumber.Text
                tbStockNumber_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub btPipes_Click(sender As Object, e As EventArgs) Handles btPipes.Click
        'define the active document as an assembly file
        If Not inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Please run this rule from the assembly file.", "Vikoma Notice")
            Exit Sub
        End If

        Dim oAsmDoc As AssemblyDocument
        oAsmDoc = inventorApp.ActiveDocument
        'oAsmName = oAsmDoc.FileName 'without extension

        oAsmName = System.IO.Path.GetFileNameWithoutExtension(oAsmDoc.FullDocumentName)

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
        Dim oRev = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        Dim oAssyName As String = tbStockNumber.Text
        'work the referenced models
        'oDocu = inventorApp.ActiveDocument
        If Not iPropertiesAddInServer.CheckReadOnly(oAsmDoc) Then
            oAsmDoc.Save2(True)
        End If

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
                Dim item As String = oBOMRow.ItemNumber

                Dim iProp As String = String.Empty
                Dim DrawnDoc = oCompDef.Document
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, "Authority", item, iProp, DrawnDoc)
                End If
        Next

        For Each oRefDoc In oRefDocs
            If Not oRefDoc.FullFileName.Contains("Route") Then
                If oRefDoc.FullFileName.Contains("pisweep") Then
                    If Not iPropertiesAddInServer.CheckReadOnly(oRefDoc) Then
                        'Dim pisWeep As String = UCase(System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullDocumentName))
                        ''Dim DeleteThese As Char() = {"P"c, "I"c, "S"c, "W"c, "E"c, "P"c, "."c}
                        'pisWeep = pisWeep.Replace("PISWEEP.", "")

                        'iProperties.GetorSetStandardiProperty(oRefDoc, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, pisWeep, "", True)

                        'Dim NewPath As String = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullDocumentName)
                        'NewPath = NewPath.Replace("pisweep.", "")

                        Dim itemNum As String = iProperties.GetorSetStandardiProperty(oRefDoc, PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties)
                        Dim itemNo As String
                        If itemNum < 10 Then
                            itemNo = "0" & itemNum
                        Else
                            itemNo = itemNum
                        End If
                        Dim NewFileName As String = System.IO.Path.GetDirectoryName(oRefDoc.FullDocumentName) & "\"
                        Dim NewPath As String = NewFileName & oAssyName & "-" & itemNo
                        Dim pisWeep As String = oAssyName & "-" & itemNo

                        iProperties.GetorSetStandardiProperty(oRefDoc, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, pisWeep, "", True)

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
                            Dim RefName As String = pisWeep + "_R" + oRev + ".stp"
                            Call oSTEPTranslator.SaveCopyAs(oRefDoc, oContext, oOptions, oData)
                            UpdateStatusBar("File saved as Step file")

                            'AttachRefFile(oAsmDoc, oData.FileName)
                            AttachFile = MsgBox("File " & RefName & " exported, attach it to main file as reference?", vbYesNo, "File Attach")
                            If AttachFile = vbYes Then
                                AddReferences(inventorApp.ActiveDocument, oData.FileName)
                                UpdateStatusBar("File attached")
                            Else
                                'Do Nothing
                            End If
                        End If
                    Else
                        UpdateStatusBar("File skipped because it's read-only")
                    End If
                Else
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
                End If
            End If
        Next

    End Sub

    Private Sub btCheckIn_Click(sender As Object, e As EventArgs) Handles btCheckIn.Click
        If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
            'If Not (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
            '    Dim PartNo As String = tbPartNumber.Text
            '    Dim StockNo As String = tbStockNumber.Text
            '    If Not PartNo = StockNo Then
            '        stockNum = MsgBox("Your Stock Number and Part Number are different, is this OK?", vbYesNo, "Stock/Part Number Check")
            '        If stockNum = vbNo Then
            '            Exit Sub
            '        End If
            '    End If
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
        'End If
    End Sub

    Private Sub btCheckOut_Click(sender As Object, e As EventArgs) Handles btCheckOut.Click
        If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
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
            Me.btCheckOut.Hide()
            Me.btCheckIn.Show()
        End If
    End Sub

    Private Sub btViewNames_Click(sender As Object, e As EventArgs) Handles btViewNames.Click
        Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
        Dim oSheet As Sheet = oDrawDoc.ActiveSheet
        Dim oSheets As Sheets = oDrawDoc.Sheets
        Dim oView As DrawingView = Nothing
        Dim oViews As DrawingViews = Nothing
        Dim oLabel As String = "DETAIL OF ITEM "
        Dim isoLabel As String = "ISOMETRIC VIEW"

        Dim strSheetName As String

        For Each oSheet In oDrawDoc.Sheets

            strSheetName = oSheet.Name
            If strSheetName.Contains("Sheet:1") Then
                Exit For
            Else
                oDrawDoc.Sheets.Item("Sheet:1").Activate()
                Exit For
            End If
        Next

        'If Not strSheetName.Contains("Sheet:1") Then
        '    oDrawDoc.Sheets.Item("Sheet:1").Activate()
        'End If

        For Each view As DrawingView In oSheet.DrawingViews
                oView = view
            Exit For
        Next

        Dim drDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument

        Dim oAssyDef As AssemblyComponentDefinition = drDoc.ComponentDefinition
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

                    Dim iProp As String = String.Empty
                    Dim DrawnDoc = oCompDef.Document
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kAuthorityDesignTrackingProperties, "Authority", item, iProp, DrawnDoc)
                End If
            Next
        oSheet.Update()

        UpdateStatusBar("BOM item numbers copied to #ITEM")

        For Each oSheet In oDrawDoc.Sheets
            oSheet.Activate()
            For Each oView In oSheet.DrawingViews
                If oView.ParentView IsNot Nothing Then
                    'we're working on a child view and should get the parent view as an object
                Else
                    If oView.IsFlatPatternView Then
                        oView.Name() = "FLAT PATTERN OF ITEM "
                        oView.ShowLabel() = True
                    ElseIf oView.ReferencedFile.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        'assembly object added but it currently also pulls in the main assembly too.
                        'Need to find a way to not show the name label of the main referenced doc.
                        oView.Name() = oLabel
                        oView.ShowLabel() = True
                    End If
                End If
            Next
        Next
        'oDrawDoc.Sheets.Item("Sheet:1").Activate()
    End Sub

    Private Sub tbPartNumber_Leave(sender As Object, e As EventArgs) Handles tbPartNumber.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            tbPartNumber.ForeColor = AddinGlobal.ForeColour
            CheckForDefaultAndUpdate(PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "Part Number", tbPartNumber.Text)

            If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Dim oDWG As DrawingDocument = inventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing

                For Each view As DrawingView In oSht.DrawingViews
                    oView = view
                    Exit For
                Next
                If oView.ReferencedDocumentDescriptor.ReferencedDocument IsNot Nothing Then
                    Dim drawnDoc As Document = oView.ReferencedDocumentDescriptor.ReferencedDocument

                    Dim drawingPN As String = tbPartNumber.Text

                    Dim iProp As String = String.Empty
                    UpdateProperties(PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "Part Number", drawingPN, iProp, drawnDoc)
                End If
            End If
        End If
    End Sub

    Private Sub btCopyPN_KeyUp(sender As Object, e As KeyEventArgs) Handles btCopyPN.KeyUp
        If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
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
        If inventorApp.ActiveDocument IsNot Nothing Then
            iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kCatalogWebLinkDesignTrackingProperties, tbNotes.Text, "", True)
            tbNotes.ForeColor = AddinGlobal.ForeColour
        End If
    End Sub

    Private Sub tbNotes_KeyUp(sender As Object, e As KeyEventArgs) Handles tbNotes.KeyUp
        If e.KeyValue = Keys.Return Then
            tbNotes_Leave(sender, e)
        End If
    End Sub

    Private Sub tbComments_Leave(sender As Object, e As EventArgs) Handles tbComments.Leave
        If inventorApp.ActiveDocument IsNot Nothing Then
            iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, tbComments.Text, "", True)
            tbComments.ForeColor = AddinGlobal.ForeColour
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
        Dim hovText As String = tbComments.Text
        ToolTip1.Show(hovText, tbComments)
    End Sub

    Private Sub tbNotes_MouseHover(sender As Object, e As EventArgs) Handles tbNotes.MouseHover
        Dim hovText As String = tbNotes.Text
        ToolTip1.Show(hovText, tbNotes)
    End Sub

    Private Sub tbNotes_Enter(sender As Object, e As EventArgs) Handles tbNotes.Enter
        Dim hovText As String = "Notes"
        ToolTip1.Show(hovText, tbNotes)
    End Sub

    ''Validates a string of alpha characters
    'Function CheckForAlphaCharacters(ByVal StringToCheck As String)
    '    For i = 0 To StringToCheck.Length - 1
    '        If Not Char.IsLetter(StringToCheck.Chars(i)) Then
    '            Return False
    '        End If
    '    Next

    '    Return True 'Return true if all elements are characters
    'End Function

    Private Sub btRevision_Click(sender As Object, e As EventArgs) Handles btRevision.Click
        Dim oDoc As Document = inventorApp.ActiveDocument
        Dim oChange As String = String.Empty
        Dim oRow As RevisionTableRow
        Dim oRows As RevisionTableRows
        Dim oSheet As Sheet = inventorApp.ActiveDocument.ActiveSheet
        Dim oInput As String = String.Empty
        Dim oSheetSize As DrawingSheetSizeEnum = oSheet.Size
        Dim oRevSelect As String = String.Empty
        Dim oDocStyles As Inventor.DrawingStylesManager = oDoc.StylesManager
        Dim oRevStyle As RevisionTableStyle = Nothing

        Dim oCreation1 As DateTime = iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                     PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties,
                                                     "", "")
        Dim oCreation As String = oCreation1.ToString("dd/MM/yyyy")

        Dim oRevTable As RevisionTable = Nothing

        Dim Rev1date As String = String.Empty
        Dim rev1rev As String = String.Empty

        Dim oNumberRev As String = String.Empty

        'oDoc.PropertySets.Item("Inventor Summary Information").Item("Author").Value ="ELC"
        oNumberRev = UCase(InputBox("Input revision letter/number, leave blank for new revision.", "REV", ""))
        oChange = UCase(InputBox("Input change number if any?", "ECN", ""))
        oInput = UCase(InputBox("What did you change?", "CHANGE", "INTRODUCED"))
        If oInput = "" Then
            Exit Sub
        End If

        If oDoc.ActiveSheet.RevisionTables.Count = 0 Then
            For Each oSheet In oDoc.Sheets
                oSheet.Activate()
                Dim oTG As TransientGeometry = inventorApp.TransientGeometry
                Dim pt As Point2d = oTG.CreatePoint2d(1, 2.56642)
                If oNumberRev = "" Then
                    oSheet.RevisionTables.Add2(pt, False, True, False, "1", , )
                Else
                    oSheet.RevisionTables.Add2(pt, False, True, False, oNumberRev, , )
                End If
                If oSheetSize = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
                    oRevStyle = oDocStyles.RevisionTableStyles.Item("VIKOMA A3")
                ElseIf oSheetSize = DrawingSheetSizeEnum.kA2DrawingSheetSize Then
                    oRevStyle = oDocStyles.RevisionTableStyles.Item("VIKOMA A2")
                ElseIf oSheetSize = DrawingSheetSizeEnum.kA1DrawingSheetSize Then
                    oRevStyle = oDocStyles.RevisionTableStyles.Item("VIKOMA A1")
                ElseIf oSheetSize = DrawingSheetSizeEnum.kA0DrawingSheetSize Then
                    oRevStyle = oDocStyles.RevisionTableStyles.Item("VIKOMA A0")
                End If
                oRevTable = oDoc.ActiveSheet.RevisionTables.Item(1)
                oRow = oRevTable.RevisionTableRows.Item(oRevTable.RevisionTableRows.Count)
                Dim oCell3 As RevisionTableCell = oRow.Item(3)
                'Set it equal to the the current date        
                'oCell4.Text= "UPDATED TO BOLTER LIB DWG"
                oCell3.Text = oInput
                oRevTable = oDoc.ActiveSheet.RevisionTables.Item(1)
                oRow1 = oRevTable.RevisionTableRows.Item(1)
                Rev1date = oRow1.Item(4).Text
                oRevTable.Style = oRevStyle

                If Not iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties, "", "") = Rev1date Then

                    inventorApp.ActiveDocument.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = Rev1date
                    UpdateStatusBar("Creation date updated to " + Rev1date)
                End If
            Next
        Else
            oSheet.Activate()
            oRevTable = oDoc.ActiveSheet.RevisionTables.Item(1)

            oRow1 = oRevTable.RevisionTableRows.Item(1)

            Rev1date = oRow1.Item(4).Text
            rev1rev = oRow1.Item(1).Text


            ' Make sure we have the active row
            If rev1rev = String.Empty Then
                Dim oCell1 As RevisionTableCell = oRow1.Item(1)
                '                    'Set it equal to the user name on the open application 

                If oNumberRev = "" Then
                    oCell1.Text = "1"
                Else
                    oCell1.Text = oNumberRev
                End If

                Dim oCell2 As RevisionTableCell = oRow1.Item(2)
                '                    'Set it equal to the user name on the open application        
                oCell2.Text = oChange

                Dim oCell3 As RevisionTableCell = oRow1.Item(3)
                'Set it equal to the the current date        
                'oCell4.Text= "UPDATED TO BOLTER LIB DWG"
                oCell3.Text = oInput

                Dim oCell4 As RevisionTableCell = oRow1.Item(4)
                'Set it equal to the the current date        
                oCell4.Text = DateTime.Now.ToString("d")

            ElseIf rev1rev IsNot String.Empty Then
                'If Rev1date <> oCreation Then
                'If rev1rev = "1" Or rev1rev = "A" Then
                '    'For Each oSheet In oDoc.Sheets

                '    oRevTable = oDoc.ActiveSheet.RevisionTables.Item(1)
                '    oRows = oRevTable.RevisionTableRows
                '    oRow = oRevTable.RevisionTableRows.Item(oRevTable.RevisionTableRows.Count)

                '    For Each oRow In oRows
                '        If oRow.IsActiveRow Then
                '            Dim oCell1 As RevisionTableCell = oRow.Item(1)
                '            '                    'Set it equal to the user name on the open application 
                '            If oNumberRev = "" Then
                '                oCell1.Text = "1"
                '            Else
                '                oCell1.Text = oNumberRev
                '            End If

                '            Dim oCell2 As RevisionTableCell = oRow.Item(2)
                '            '                    'Set it equal to the user name on the open application        
                '            oCell2.Text = oChange

                '            Dim oCell3 As RevisionTableCell = oRow.Item(3)
                '            'Set it equal to the the current date        
                '            'oCell4.Text= "UPDATED TO BOLTER LIB DWG"
                '            oCell3.Text = oInput

                '            Dim oCell4 As RevisionTableCell = oRow.Item(4)
                '            'Set it equal to the the current date        
                '            oCell4.Text = DateTime.Now.ToString("d")

                '        Else
                '            oRow.Delete() 'deletes rev tags also
                '        End If
                '    Next

                'End If
                'Else
                oRevTable.RevisionTableRows.Add()
                    oRow = oRevTable.RevisionTableRows.Item(oRevTable.RevisionTableRows.Count)
                    If oRow.IsActiveRow Then
                        Dim oCell1 As RevisionTableCell = oRow.Item(1)
                        '                    'Set it equal to the user name on the open application 
                        If oNumberRev = "" Then
                            OldRev = oRevTable.RevisionTableRows.Item(oRevTable.RevisionTableRows.Count - 1).Item(1).Text
                            If Char.IsLetter(OldRev) Then
                                oCell1.Text = "1"
                            Else
                                oCell1.Text = oRevTable.RevisionTableRows.Item(oRevTable.RevisionTableRows.Count - 1).Item(1).Text + 1
                            End If
                        Else
                            oCell1.Text = oNumberRev
                        End If

                        Dim oCell2 As RevisionTableCell = oRow.Item(2)
                        '                    'Set it equal to the user name on the open application        
                        oCell2.Text = oChange

                        Dim oCell3 As RevisionTableCell = oRow.Item(3)
                        'Set it equal to the the current date        
                        'oCell4.Text= "UPDATED TO BOLTER LIB DWG"
                        oCell3.Text = oInput

                        Dim oCell4 As RevisionTableCell = oRow.Item(4)
                        'Set it equal to the the current date        
                        oCell4.Text = DateTime.Now.ToString("d")
                    End If
                End If
        End If
        'End If

    End Sub

    Private Sub btExpDXF_Click(sender As Object, e As EventArgs) Handles btExpDXF.Click
        If TypeOf inventorApp.ActiveDocument IsNot PartDocument Then

        Else
            'Set a reference to the active document (the document to be published).
            Dim oDoc As Document = inventorApp.ActiveDocument

            Dim oCompDef As SheetMetalComponentDefinition
            oCompDef = oDoc.ComponentDefinition
            If oCompDef.HasFlatPattern = False Then
                oCompDef.Unfold()

            Else
                oCompDef.FlatPattern.Edit()
            End If

            ' Get the DataIO object.
            Dim oDataIO As DataIO = oDoc.ComponentDefinition.DataIO

            ' Build the string that defines the format of the DXF file.
            Dim sOut As String = "FLAT PATTERN DXF?AcadVersion=2000&OuterProfileLayer=IV_INTERIOR_PROFILES&BendUpLayerColor=255;0;0&BendDownLayerColor=255;0;0&BendLayerColor=255;0;0&InvisibleLayers=IV_TANGENT;IV_ROLL_TANGENT"
            Dim oRev As String = iProperties.GetorSetStandardiProperty(oDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation)
            Dim fileFolder As String = System.IO.Path.GetDirectoryName(oDoc.FullFileName)
            Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullDocumentName)
            Dim sFname As String = fileFolder & "\" & fileName & "_R" & oRev & ".dxf"

            'Export the DXF and fold the model back up
            oCompDef.DataIO.WriteDataToFile(sOut, sFname)
            Dim oSMDef As SheetMetalComponentDefinition
            oSMDef = oDoc.ComponentDefinition
            oSMDef.FlatPattern.ExitEdit()

            AttachRefFile(oDoc, sFname)
        End If
    End Sub

    Private Sub btAttachFile_Click(sender As Object, e As EventArgs) Handles btAttachFile.Click
        Dim oDoc As Document = inventorApp.ActiveDocument
        Dim oFileDlg As Inventor.FileDialog = Nothing
        inventorApp.CreateFileDialog(oFileDlg)
        oFileDlg.Filter = "All Files (*.*)|*.*"
        oFileDlg.FilterIndex = 1
        oFileDlg.DialogTitle = "Select file to attach"
        oFileDlg.InitialDirectory = "C:\VAULT WORKING FOLDER\Designs"
        oFileDlg.CancelError = True
        On Error Resume Next
        oFileDlg.ShowOpen()
        If Err.Number <> 0 Then

        ElseIf oFileDlg.FileName <> "" Then
            AddReferences(oDoc, oFileDlg.FileName)
            MsgBox("File attached")
        End If
    End Sub

    Private Sub btFrame_Click(sender As Object, e As EventArgs) Handles btFrame.Click
        If TypeOf AddinGlobal.InventorApp.ActiveDocument Is DrawingDocument Then
            Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

            Dim oSht As Sheet = oDWG.ActiveSheet

            Dim oView As DrawingView = Nothing
            Dim drDoc As Document = Nothing

            For Each view As DrawingView In oSht.DrawingViews
                oView = view
                Exit For
            Next

            drDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
            Dim oAsm As AssemblyDocument = drDoc

            Dim doc = drDoc
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            If Not oBOM.StructuredViewEnabled = True Then
                oBOM.StructuredViewEnabled = True
            End If

            Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

            Dim oBOMRow As BOMRow

            For Each oBOMRow In oBOMView.BOMRows

                'Set a reference to the primary ComponentDefinition of the row
                Dim oCompDef As ComponentDefinition
                oCompDef = oBOMRow.ComponentDefinitions.Item(1)
                If oCompDef.Document.FullDocumentName.Contains("Content Center") Or oCompDef.Document.FullDocumentName.Contains("Bought Out") Then

                Else
                    Dim CompFullDocumentName As String = oCompDef.Document.FullDocumentName
                    Dim CompFileNameOnly As String = System.IO.Path.GetFileNameWithoutExtension(CompFullDocumentName)

                    Dim item As Integer
                    item = oBOMRow.ItemNumber

                    Dim DrawnDoc = oCompDef.Document
                    Dim oAsmName As String
                    If oAsm.DisplayName.Contains(".iam") Then
                        oAsmName = RemoveCharacter(oAsm.DisplayName, ".iam")
                    Else
                        oAsmName = oAsm.DisplayName
                    End If

                    If DrawnDoc.DocumentInterests.HasInterest("{AC211AE0-A7A5-4589-916D-81C529DA6D17}") _ 'Frame generator component
                        AndAlso DrawnDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject _ 'Part
                        AndAlso DrawnDoc.IsModifiable _  'Modifiable (not reference skeleton)
                        AndAlso oAsm.ComponentDefinition.Occurrences.AllReferencedOccurrences(DrawnDoc).Count > 0 Then 'Exists in assembly (not derived base component)

                        Dim invDesignInfo As PropertySet
                        Dim CorrectName As String
                        If oAsm.DisplayName.Contains(".iam") Then
                            If item < 10 Then
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-0" & item
                            Else
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-" & item
                            End If
                        Else
                            If item < 10 Then
                                CorrectName = DrawnDoc.DisplayName & "-0" & item
                            Else
                                CorrectName = DrawnDoc.DisplayName & "-" & item
                            End If
                        End If
                        invDesignInfo = DrawnDoc.PropertySets.Item("Design Tracking Properties")
                        invDesignInfo.Item("Part Number").Value = CorrectName
                    ElseIf DrawnDoc.DisplayName.Contains(oAsmName) AndAlso DrawnDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject _ 'Part
                        AndAlso DrawnDoc.IsModifiable _  'Modifiable (not reference skeleton)
                        AndAlso oAsm.ComponentDefinition.Occurrences.AllReferencedOccurrences(DrawnDoc).Count > 0 Then 'Exists in assembly (not derived base component)

                        Dim invDesignInfo As PropertySet
                        Dim CorrectName As String
                        If oAsm.DisplayName.Contains(".iam") Then
                            If item < 10 Then
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-0" & item
                            Else
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-" & item
                            End If
                        Else
                            If item < 10 Then
                                CorrectName = DrawnDoc.DisplayName & "-0" & item
                            Else
                                CorrectName = DrawnDoc.DisplayName & "-" & item
                            End If
                        End If
                        invDesignInfo = DrawnDoc.PropertySets.Item("Design Tracking Properties")
                        invDesignInfo.Item("Part Number").Value = CorrectName
                    End If
                End If
            Next
            oSht.Update()
        ElseIf TypeOf AddinGlobal.InventorApp.ActiveDocument Is AssemblyDocument Then
            Dim doc = inventorApp.ActiveEditDocument
            Dim oAsm As AssemblyDocument = inventorApp.ActiveDocument
            Dim oAssyDef As AssemblyComponentDefinition = doc.ComponentDefinition
            Dim oBOM As BOM = oAssyDef.BOM

            If Not oBOM.StructuredViewEnabled = True Then
                oBOM.StructuredViewEnabled = True
            End If

            Dim oBOMView As BOMView = oBOM.BOMViews.Item("Structured")

            Dim oBOMRow As BOMRow

            For Each oBOMRow In oBOMView.BOMRows

                'Set a reference to the primary ComponentDefinition of the row
                Dim oCompDef As ComponentDefinition
                oCompDef = oBOMRow.ComponentDefinitions.Item(1)
                If oCompDef.Document.FullDocumentName.Contains("Content Center") Or oCompDef.Document.FullDocumentName.Contains("Bought Out") Then

                Else
                    Dim CompFullDocumentName As String = oCompDef.Document.FullDocumentName
                    Dim CompFileNameOnly As String = System.IO.Path.GetFileNameWithoutExtension(CompFullDocumentName)

                    Dim item As Integer
                    item = oBOMRow.ItemNumber

                    Dim DrawnDoc = oCompDef.Document
                    Dim oAsmName As String
                    If oAsm.DisplayName.Contains(".iam") Then
                        oAsmName = RemoveCharacter(oAsm.DisplayName, ".iam")
                    Else
                        oAsmName = oAsm.DisplayName
                    End If

                    If DrawnDoc.DocumentInterests.HasInterest("{AC211AE0-A7A5-4589-916D-81C529DA6D17}") _ 'Frame generator component
                        AndAlso DrawnDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject _ 'Part
                        AndAlso DrawnDoc.IsModifiable _  'Modifiable (not reference skeleton)
                        AndAlso oAsm.ComponentDefinition.Occurrences.AllReferencedOccurrences(DrawnDoc).Count > 0 Then 'Exists in assembly (not derived base component)

                        Dim invDesignInfo As PropertySet
                        Dim CorrectName As String
                        If oAsm.DisplayName.Contains(".iam") Then
                            If item < 10 Then
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-0" & item
                            Else
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-" & item
                            End If
                        Else
                            If item < 10 Then
                                CorrectName = DrawnDoc.DisplayName & "-0" & item
                            Else
                                CorrectName = DrawnDoc.DisplayName & "-" & item
                            End If
                        End If
                        invDesignInfo = DrawnDoc.PropertySets.Item("Design Tracking Properties")
                        invDesignInfo.Item("Part Number").Value = CorrectName
                    ElseIf DrawnDoc.DisplayName.Contains(oAsmName) AndAlso DrawnDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject _ 'Part
                    AndAlso DrawnDoc.IsModifiable _  'Modifiable (not reference skeleton)
                    AndAlso oAsm.ComponentDefinition.Occurrences.AllReferencedOccurrences(DrawnDoc).Count > 0 Then 'Exists in assembly (not derived base component)

                        Dim invDesignInfo As PropertySet
                        Dim CorrectName As String
                        If oAsm.DisplayName.Contains(".iam") Then
                            If item < 10 Then
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-0" & item
                            Else
                                CorrectName = RemoveCharacter(oAsm.DisplayName, ".iam") & "-" & item
                            End If
                        Else
                            If item < 10 Then
                                CorrectName = DrawnDoc.DisplayName & "-0" & item
                            Else
                                CorrectName = DrawnDoc.DisplayName & "-" & item
                            End If
                        End If
                        invDesignInfo = DrawnDoc.PropertySets.Item("Design Tracking Properties")
                        invDesignInfo.Item("Part Number").Value = CorrectName
                    End If
                End If
            Next
        End If
        'Dim oAsm As AssemblyDocument = inventorApp.ActiveDocument

        'Dim oDoc As Document

        'For Each oDoc In oAsm.AllReferencedDocuments
        '    If oDoc.DocumentInterests.HasInterest("{AC211AE0-A7A5-4589-916D-81C529DA6D17}") _ 'Frame generator component
        '        AndAlso oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject _ 'Part
        '        AndAlso oDoc.IsModifiable _  'Modifiable (not reference skeleton)
        '        AndAlso oAsm.ComponentDefinition.Occurrences.AllReferencedOccurrences(oDoc).Count > 0 Then 'Exists in assembly (not derived base component)

        '        Dim invDesignInfo As PropertySet
        '        Dim CorrectName As String
        '        If oDoc.DisplayName.Contains(".ipt") Then
        '            CorrectName = RemoveCharacter(oDoc.DisplayName, ".ipt")
        '        Else
        '            CorrectName = oDoc.DisplayName
        '        End If
        '        invDesignInfo = oDoc.PropertySets.Item("Design Tracking Properties")
        '            'set the member part number text to be the same as the display name
        '            invDesignInfo.Item("Part Number").Value = CorrectName
        '        End If
        'Next
    End Sub

    Function RemoveCharacter(ByVal stringToCleanUp As String, ByVal characterToRemove As String)
        ' replace the target with nothing
        ' Replace() returns a new String and does not modify the current one
        Return stringToCleanUp.Replace(characterToRemove, "")
    End Function

    'Public Shared Function GetFileName(path As String) As String
    'End Function
End Class
