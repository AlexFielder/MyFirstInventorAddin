Imports System.Windows.Forms
Imports Inventor
Imports log4net
Imports iPropertiesController.iPropertiesController
Imports System.Collections.Generic
Imports System.Drawing

Public Class iPropertiesForm

    'Dim rs As New Resizer

    'Private Sub iPropertiesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    rs.FindAllControls(Me)

    'End Sub

    'Private Sub iPropertiesForm_Resize(sender As Object, e As EventArgs) Handles Me.Resize
    '    rs.ResizeAllControls(Me)
    'End Sub

    Private inventorApp As Inventor.Application
    Private localWindow As DockableWindow
    Private value As String

    Public ReadOnly log As ILog = LogManager.GetLogger(GetType(iPropertiesForm))

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
                        CurrentPath = System.IO.Path.GetDirectoryName(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)
                        NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(AddinGlobal.InventorApp.ActiveDocument.FullDocumentName)
                        Dim oDrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                        Dim oSht As Sheet = oDrawDoc.ActiveSheet
                        Dim oView As DrawingView = Nothing
                        'Dim RefDoc As Document = Nothing

                        'For i As Integer = 1 To oSht.DrawingViews.Count
                        '    oView = oSht.DrawingViews(i)
                        '    Exit For
                        'Next

                        For Each view As DrawingView In oSht.DrawingViews
                            oView = view
                            Exit For
                        Next

                        RefDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                        RefNewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
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

    Private Sub tbPartNumber_Leave(sender As Object, e As EventArgs) Handles tbPartNumber.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    tbPartNumber.ForeColor = Drawing.Color.Black
                    Dim oDrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    Dim oSheets As Sheets = Nothing
                    Dim oViews As DrawingViews = Nothing
                    Dim oScale As Double = Nothing
                    Dim oViewCount As Integer = 0
                    Dim oTitleBlock = oSheet.TitleBlock
                    Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

                    Dim oSht As Sheet = oDWG.ActiveSheet

                    Dim oView As DrawingView = Nothing
                    Dim drawnDoc As Document = Nothing

                    For Each view As DrawingView In oSht.DrawingViews
                        oView = view
                        Exit For
                    Next

                    drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                    If tbPartNumber.Text = "Part Number" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(drawnDoc,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          tbPartNumber.Text,
                                                          "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Part Number Updated to: " + iPropPartNum)
                        UpdateStatusBar("Part Number updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject IsNot Nothing Then

                    If tbPartNumber.Text = "Part Number" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                              PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          tbPartNumber.Text,
                                                          "")
                        log.Debug(inventorApp.ActiveEditObject.FullFileName + " Part Number Updated to: " + iPropPartNum)
                        UpdateStatusBar("Part Number updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject Is Nothing Then
                    If tbPartNumber.Text = "Part Number" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                              PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                                                          tbPartNumber.Text,
                                                          "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Part Number Updated to: " + iPropPartNum)
                        UpdateStatusBar("Part Number updated to " + iPropPartNum)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tbDescription_Leave(sender As Object, e As EventArgs) Handles tbDescription.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    tbDescription.ForeColor = Drawing.Color.Black
                    Dim oDrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    Dim oSheets As Sheets = Nothing
                    Dim oViews As DrawingViews = Nothing
                    Dim oScale As Double = Nothing
                    Dim oViewCount As Integer = 0
                    Dim oTitleBlock = oSheet.TitleBlock
                    Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

                    Dim oSht As Sheet = oDWG.ActiveSheet

                    Dim oView As DrawingView = Nothing
                    Dim drawnDoc As Document = Nothing

                    For Each view As DrawingView In oSht.DrawingViews
                        oView = view
                        Exit For
                    Next

                    drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                    If tbDescription.Text = "Description" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              tbDescription.Text,
                                                              "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Description Updated to: " + iPropPartNum)
                        UpdateStatusBar("Description updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject IsNot Nothing Then
                    If tbDescription.Text = "Description" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              tbDescription.Text,
                                                              "")
                        log.Debug(inventorApp.ActiveEditObject.FullFileName + " Description Updated to: " + iPropPartNum)
                        UpdateStatusBar("Description updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject Is Nothing Then
                    If tbDescription.Text = "Description" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                              PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                                                              tbDescription.Text,
                                                              "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Description Updated to: " + iPropPartNum)
                        UpdateStatusBar("Description updated to " + iPropPartNum)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tbStockNumber_Leave(sender As Object, e As EventArgs) Handles tbStockNumber.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                tbStockNumber.ForeColor = Drawing.Color.Black
                If inventorApp.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    'Do Nothing
                Else
                    If inventorApp.ActiveEditObject IsNot Nothing Then
                        If tbStockNumber.Text = "Stock Number" Then
                            Dim iPropPartNum As String =
                            iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                                  PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                                  "",
                                                                  "")
                        Else
                            Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                              PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                              tbStockNumber.Text,
                                                              "")
                            log.Debug(inventorApp.ActiveEditObject.FullFileName + " Stock Number Updated to: " + iPropPartNum)
                            UpdateStatusBar("Stock Number updated to " + iPropPartNum)
                        End If
                    ElseIf inventorApp.ActiveEditObject Is Nothing Then
                        If tbStockNumber.Text = "Stock Number" Then
                            Dim iPropPartNum As String =
                            iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                                  PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                                  "",
                                                                  "")
                        Else
                            Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                              PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties,
                                                              tbStockNumber.Text,
                                                              "")
                            log.Debug(inventorApp.ActiveDocument.FullFileName + " Stock Number Updated to: " + iPropPartNum)
                            UpdateStatusBar("Stock Number updated to " + iPropPartNum)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tbEngineer_Leave(sender As Object, e As EventArgs) Handles tbEngineer.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If inventorApp.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    tbEngineer.ForeColor = Drawing.Color.Black
                    Dim oDrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                    Dim oSheet As Sheet = oDrawDoc.ActiveSheet
                    Dim oSheets As Sheets = Nothing
                    Dim oViews As DrawingViews = Nothing
                    Dim oScale As Double = Nothing
                    Dim oViewCount As Integer = 0
                    Dim oTitleBlock = oSheet.TitleBlock
                    Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

                    Dim oSht As Sheet = oDWG.ActiveSheet

                    Dim oView As DrawingView = Nothing
                    Dim drawnDoc As Document = Nothing

                    For Each view As DrawingView In oSht.DrawingViews
                        oView = view
                        Exit For
                    Next

                    drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                    If tbEngineer.Text = "Engineer" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(drawnDoc,
                                                              PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                              tbEngineer.Text,
                                                              "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Engineer Updated to: " + iPropPartNum)
                        UpdateStatusBar("Engineer updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject IsNot Nothing Then
                    If tbEngineer.Text = "Engineer" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                              PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveEditObject,
                                                          PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                          tbEngineer.Text,
                                                          "")
                        log.Debug(inventorApp.ActiveEditObject.FullFileName + " Engineer Updated to: " + iPropPartNum)
                        UpdateStatusBar("Engineer updated to " + iPropPartNum)
                    End If
                ElseIf inventorApp.ActiveEditObject Is Nothing Then
                    If tbEngineer.Text = "Engineer" Then
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                              PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                              "",
                                                              "")
                    Else
                        Dim iPropPartNum As String =
                        iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties,
                                                          tbEngineer.Text,
                                                          "")
                        log.Debug(inventorApp.ActiveDocument.FullFileName + " Engineer Updated to: " + iPropPartNum)
                        UpdateStatusBar("Engineer updated to " + iPropPartNum)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btUpdateAll_Click(sender As Object, e As EventArgs) Handles btUpdateAll.Click

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
                tbPartNumber_Leave(sender, e)
                tbDescription_Leave(sender, e)
                tbStockNumber_Leave(sender, e)
                tbEngineer_Leave(sender, e)
                tbDrawnBy_Leave(sender, e)

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
                UpdateStatusBar("iProperties updated")
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

        inventorApp.ActiveDocument.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = DateTimePicker1.Value
        UpdateStatusBar("Creation date updated to " + DateTimePicker1.Value)
    End Sub

    Private Sub tbDrawnBy_Leave(sender As Object, e As EventArgs) Handles tbDrawnBy.Leave
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                Dim iPropPartNum As String =
                    iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument,
                                                          PropertiesForSummaryInformationEnum.kAuthorSummaryInformation,
                                                          tbDrawnBy.Text,
                                                          "")
                log.Debug(inventorApp.ActiveDocument.FullFileName + " Author Updated to: " + iPropPartNum)
                UpdateStatusBar("Author updated to " + iPropPartNum)
            End If
        End If
    End Sub

    Private Sub btITEM_Click(sender As Object, e As EventArgs) Handles btITEM.Click
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
            For Each oSheet In oDrawDoc.Sheets
                'i = i+1
                inventorApp.ActiveDocument.Sheets.Item(i).Activate
                oTitleBlock = oSheet.TitleBlock
                oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
                For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
                    Select Case oTextBox.Text
                        Case "<Material>"
                            oPromptEntry = oTitleBlock.GetResultText(oTextBox)
                    End Select
                Next
            Next

            If oPromptEntry = "<Material>" Then
                oPromptText = "Engineer"
            ElseIf oPromptEntry = "" Then
                oPromptText = "Engineer"
            Else
                oPromptText = oPromptEntry
            End If

            'For i As Integer = 1 To oSht.DrawingViews.Count
            '    oView = oSht.DrawingViews(i)
            '    Exit For
            'Next

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
        For Each oSheet In oDrawDoc.Sheets
            'i = i+1
            inventorApp.ActiveDocument.Sheets.Item(i).Activate
            oTitleBlock = oSheet.TitleBlock
            oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
            For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
                Select Case oTextBox.Text
                    Case "<Scale>"
                        oPromptEntry = oTitleBlock.GetResultText(oTextBox)
                End Select
            Next
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
        'Dim scaleTextBox As Inventor.TextBox = GetScaleTextBox(oTitleBlock.Definition)
        'Dim scaleString As String = scaleTextBox.Text

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

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim insertText = "Ø"
        Dim insertPos As Integer = tbEngineer.SelectionStart
        Dim focusPoint = insertPos + insertText.Length
        If tbEngineer.Text = "Engineer" Then
            tbEngineer.Text = insertText
            tbEngineer.Focus()
            tbEngineer.Select(insertPos + insertText.Length, 0)
        Else
            tbEngineer.Text = tbEngineer.Text.Insert(insertPos, insertText)
            tbEngineer.Focus()
            tbEngineer.Select(insertPos + insertText.Length, 0)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim insertText = "°"
        Dim insertPos As Integer = tbEngineer.SelectionStart
        If tbEngineer.Text = "Engineer" Then
            tbEngineer.Text = insertText
            tbEngineer.Focus()
            tbEngineer.Select(insertPos + insertText.Length, 0)
        Else
            tbEngineer.Text = tbEngineer.Text.Insert(insertPos, insertText)
            tbEngineer.Focus()
            tbEngineer.Select(insertPos + insertText.Length, 0)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim insertText = "°"
        Dim insertPos As Integer = tbDescription.SelectionStart
        If tbDescription.Text = "Description" Then
            tbDescription.Text = insertText
            tbDescription.Focus()
            tbDescription.Select(insertPos + insertText.Length, 0)
        Else
            tbDescription.Text = tbDescription.Text.Insert(insertPos, insertText)
            tbDescription.Focus()
            tbDescription.Select(insertPos + insertText.Length, 0)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim insertText = "Ø"
        Dim insertPos As Integer = tbDescription.SelectionStart
        If tbDescription.Text = "Description" Then
            tbDescription.Text = insertText
            tbDescription.Focus()
            tbDescription.Select(insertPos + insertText.Length, 0)
        Else
            tbDescription.Text = tbDescription.Text.Insert(insertPos, insertText)
            tbDescription.Focus()
            tbDescription.Select(insertPos + insertText.Length, 0)
        End If
    End Sub

    Public Sub AttachRefFile(ActiveDoc As Document, RefFile As String)
        AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
        If AttachFile = vbYes Then
            AddReferences(ActiveDoc, RefFile)
            UpdateStatusBar("File attached")
        Else
            'Do Nothing
        End If
    End Sub

    Private Sub btExpStp_Click(sender As Object, e As EventArgs) Handles btExpStp.Click
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
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
        ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
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
        End If
    End Sub

    Private Sub btExpStl_Click(sender As Object, e As EventArgs) Handles btExpStl.Click
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            'GetNewFilePaths()
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

            If oSTLTranslator.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

                ' Set output file type:
                '   0 - binary,  1 - ASCII
                oOptions.Value("OutputFileType") = 0
                oOptions.Value("ExportUnits") = 5
                ' Set accuracy.
                '   2 = High,  1 = Medium,  0 = Low
                'oOptions.Value("Resolution") = 2
                oOptions.Value("SurfaceDeviation") = 0.005
                oOptions.Value("NormalDeviation") = 10
                oOptions.Value("MaxEdgeLength") = 100
                oOptions.Value("AspectRatio") = 21.5
                oOptions.Value("ExportColor") = True

                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                Dim oData As DataMedium
                oData = inventorApp.TransientObjects.CreateDataMedium
                oData.FileName = NewPath + "_R" + oRev + ".stl"

                Call oSTLTranslator.SaveCopyAs(oDoc, oContext, oOptions, oData)
                UpdateStatusBar("File saved as STL file")
                AttachRefFile(inventorApp.ActiveDocument, oData.FileName)
            End If
        ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
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

                    ' Set accuracy.
                    '   2 = High,  1 = Medium,  0 = Low
                    oOptions.Value("Resolution") = 2

                    ' Set output file type:
                    '   0 - binary,  1 - ASCII
                    oOptions.Value("OutputFileType") = 0

                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

                    Dim oData As DataMedium
                    oData = inventorApp.TransientObjects.CreateDataMedium
                    oData.FileName = RefNewPath + "_R" + oRev + ".stl"

                    Call oSTLTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
                    UpdateStatusBar("Part/Assy file saved as STL file")
                    AttachRefFile(RefDoc, oData.FileName)
                End If
            End If
        End If
    End Sub

    Private Sub btExpPdf_Click(sender As Object, e As EventArgs) Handles btExpPdf.Click
        If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
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
        End If
    End Sub

    Private Sub tbPartNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbPartNumber.KeyPress
        If e.KeyChar = Chr(9) Then
            tbDescription.Focus()
        ElseIf e.KeyChar = Chr(13) Then
            tbPartNumber_Leave(sender, e)
        End If
    End Sub

    Private Sub tbDescription_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescription.KeyPress
        If e.KeyChar = Chr(9) Then
            tbStockNumber.Focus()
            'ElseIf e.KeyChar = Chr(33 And 38) Then
            '    tbPartNumber.Focus()
        ElseIf e.KeyChar = Chr(13) Then
            tbDescription_Leave(sender, e)
        End If
    End Sub

    Private Sub tbStockNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbStockNumber.KeyPress
        If e.KeyChar = Chr(9) Then
            tbEngineer.Focus()
            'ElseIf e.KeyChar = Chr(33 And 38) Then
            '    tbDescription.Focus()
        ElseIf e.KeyChar = Chr(13) Then
            tbStockNumber_Leave(sender, e)
        End If
    End Sub

    Private Sub tbEngineer_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbEngineer.KeyPress
        If e.KeyChar = Chr(9) Then
            btUpdateAll.Focus()
            'ElseIf e.KeyChar = Chr(33 And 38) Then
            '    tbStockNumber.Focus()
        ElseIf e.KeyChar = Chr(13) Then
            tbEngineer_Leave(sender, e)
        End If
    End Sub

    Private Sub btUpdateAll_KeyPress(sender As Object, e As KeyPressEventArgs) Handles btUpdateAll.KeyPress
        If e.KeyChar = Chr(9) Then
            btUpdateAll_Click(sender, e)
            'ElseIf e.KeyChar = Chr(33 And 38) Then
            '    tbEngineer.Focus()
        ElseIf e.KeyChar = Chr(13) Then
            btUpdateAll_Click(sender, e)
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

    Private Sub tbDescription_TextChanged(sender As Object, e As EventArgs) Handles tbDescription.TextChanged
        tbDescription.ForeColor = Drawing.Color.Red
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

    'Private Sub btUpdateAssy_Click(sender As Object, e As EventArgs)
    '    If RefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
    '        inventorApp.CommandManager.ControlDefinitions.Item("AssemblyGlobalUpdateCmd").Execute()
    '        inventorApp.CommandManager.ControlDefinitions.Item("AppGlobalUpdateWrapperCmd").Execute()
    '    Else
    '        UpdateStatusBar("Can't update a part!! Come on, you know this!")
    '    End If
    'End Sub
End Class

'-------------------------------------------------------------------------------
' Resizer
' This class is used to dynamically resize and reposition all controls on a form.
' Container controls are processed recursively so that all controls on the form
' are handled.
'
' Usage:
'  Resizing functionality requires only three lines of code on a form:
'
'  1. Create a form-level reference to the Resize class:
'     Dim myResizer as Resizer
'
'  2. In the Form_Load event, call the  Resizer class FIndAllControls method:
'     myResizer.FindAllControls(Me)
'
'  3. In the Form_Resize event, call the  Resizer class ResizeAllControls method:
'     myResizer.ResizeAllControls(Me)
'
'-------------------------------------------------------------------------------
'Public Class Resizer

'    '----------------------------------------------------------
'    ' ControlInfo
'    ' Structure of original state of all processed controls
'    '----------------------------------------------------------
'    Private Structure ControlInfo
'        Public name As String
'        Public parentName As String
'        Public leftOffsetPercent As Double
'        Public topOffsetPercent As Double
'        Public heightPercent As Double
'        Public originalHeight As Integer
'        Public originalWidth As Integer
'        Public widthPercent As Double
'        Public originalFontSize As Single
'    End Structure

'    '-------------------------------------------------------------------------
'    ' ctrlDict
'    ' Dictionary of (control name, control info) for all processed controls
'    '-------------------------------------------------------------------------
'    Private ctrlDict As Dictionary(Of String, ControlInfo) = New Dictionary(Of String, ControlInfo)

'    '----------------------------------------------------------------------------------------
'    ' FindAllControls
'    ' Recursive function to process all controls contained in the initially passed
'    ' control container and store it in the Control dictionary
'    '----------------------------------------------------------------------------------------
'    Public Sub FindAllControls(thisCtrl As Control)

'        '-- If the current control has a parent, store all original relative position
'        '-- and size information in the dictionary.
'        '-- Recursively call FindAllControls for each control contained in the
'        '-- current Control
'        For Each ctl As Control In thisCtrl.Controls
'            Try
'                If Not IsNothing(ctl.Parent) Then
'                    Dim parentHeight = ctl.Parent.Height
'                    Dim parentWidth = ctl.Parent.Width

'                    Dim c As New ControlInfo
'                    c.name = ctl.Name
'                    c.parentName = ctl.Parent.Name
'                    c.topOffsetPercent = Convert.ToDouble(ctl.Top) / Convert.ToDouble(parentHeight)
'                    c.leftOffsetPercent = Convert.ToDouble(ctl.Left) / Convert.ToDouble(parentWidth)
'                    c.heightPercent = Convert.ToDouble(ctl.Height) / Convert.ToDouble(parentHeight)
'                    c.widthPercent = Convert.ToDouble(ctl.Width) / Convert.ToDouble(parentWidth)
'                    c.originalFontSize = ctl.Font.Size
'                    c.originalHeight = ctl.Height
'                    c.originalWidth = ctl.Width
'                    ctrlDict.Add(c.name, c)
'                End If

'            Catch ex As Exception
'                Debug.Print(ex.Message)
'            End Try

'            If ctl.Controls.Count > 0 Then
'                FindAllControls(ctl)
'            End If

'        Next '-- For Each

'    End Sub

'    '----------------------------------------------------------------------------------------
'    ' ResizeAllControls
'    ' Recursive function to resize and reposition all controls contained in the Control
'    ' dictionary
'    '----------------------------------------------------------------------------------------
'    Public Sub ResizeAllControls(thisCtrl As Control)

'        Dim fontRatioW As Single
'        Dim fontRatioH As Single
'        Dim fontRatio As Single
'        Dim f As Font

'        '-- Resize and reposition all controls in the passed control
'        For Each ctl As Control In thisCtrl.Controls
'            Try
'                If Not IsNothing(ctl.Parent) Then
'                    Dim parentHeight = ctl.Parent.Height
'                    Dim parentWidth = ctl.Parent.Width

'                    Dim c As New ControlInfo

'                    Dim ret As Boolean = False
'                    Try
'                        '-- Get the current control's info from the control info dictionary
'                        ret = ctrlDict.TryGetValue(ctl.Name, c)

'                        '-- If found, adjust the current control based on control relative
'                        '-- size and position information stored in the dictionary
'                        If (ret) Then
'                            '-- Size
'                            ctl.Width = Int(parentWidth * c.widthPercent)
'                            ctl.Height = Int(parentHeight * c.heightPercent)

'                            '-- Position
'                            ctl.Top = Int(parentHeight * c.topOffsetPercent)
'                            ctl.Left = Int(parentWidth * c.leftOffsetPercent)

'                            '-- Font
'                            f = ctl.Font
'                            fontRatioW = ctl.Width / c.originalWidth
'                            fontRatioH = ctl.Height / c.originalHeight
'                            fontRatio = (fontRatioW +
'                            fontRatioH) / 2 '-- average change in control Height and Width
'                            ctl.Font = New Font(f.FontFamily,
'                            c.originalFontSize * fontRatio, f.Style)

'                        End If
'                    Catch
'                    End Try
'                End If
'            Catch ex As Exception
'            End Try

'            '-- Recursive call for controls contained in the current control
'            If ctl.Controls.Count > 0 Then
'                ResizeAllControls(ctl)
'            End If

'        Next '-- For Each
'    End Sub

'End Class