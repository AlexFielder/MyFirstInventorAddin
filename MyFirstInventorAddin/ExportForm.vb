Imports System.Windows.Forms
Imports Inventor
Imports iPropertiesController.iPropertiesController
Imports log4net

Public Class ExportForm
    Private inventorApp As Inventor.Application
    Public Shared iPropsForm As IPropertiesForm = Nothing
    Private localWindow As DockableWindow
    Private value As String

    Public CurrentPath As String = String.Empty
    Public NewPath As String = String.Empty
    Public RefNewPath As String = String.Empty
    Public RefDoc As Document = Nothing

    'Private Sub UpdateStatusBar(ByVal Message As String)
    '    AddinGlobal.InventorApp.StatusBarText = Message
    'End Sub

    'Public Sub AttachRefFile(ActiveDoc As Document, RefFile As String)
    '    AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
    '    If AttachFile = vbYes Then
    '        AddReferences(ActiveDoc, RefFile)
    '        UpdateStatusBar("File attached")
    '    Else
    '        'Do Nothing
    '    End If
    'End Sub

    'Public Sub AddReferences(ByVal odoc As Inventor.Document, ByVal selectedfile As String)
    '    Dim oleReference As ReferencedOLEFileDescriptor
    '    oleReference = odoc.ReferencedOLEFileDescriptors _
    '                .Add(selectedfile, OLEDocumentTypeEnum.kOLEDocumentLinkObject)
    '    oleReference.BrowserVisible = True
    '    oleReference.Visible = False
    '    oleReference.DisplayName = System.IO.Path.GetFileName(selectedfile)
    'End Sub

    Private Sub btPipes_Click(sender As Object, e As EventArgs) Handles btExpPipes.Click
        iPropsForm.btPipes(sender, e)
        ''define the active document as an assembly file
        'Dim oAsmDoc As AssemblyDocument
        'oAsmDoc = AddinGlobal.InventorApp.ActiveDocument
        ''oAsmName = oAsmDoc.FileName 'without extension

        'oAsmName = System.IO.Path.GetFileNameWithoutExtension(oAsmDoc.FullDocumentName)

        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        '    MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
        '    Exit Sub
        'End If
        ''get user input
        'RUsure = MessageBox.Show(
        '"This will create a STEP file for all components." _
        '& vbLf & " " _
        '& vbLf & "Are you sure you want to create STEP Drawings for all of the assembly components?" _
        '& vbLf & "This could take a while.", "Batch Output STEPs ", MessageBoxButtons.YesNo)
        'If RUsure = vbNo Then
        '    Return
        'Else
        'End If
        ''- - - - - - - - - - - - -STEP setup - - - - - - - - - - - -
        ''oPath = oAsmDoc.Path
        '''get STEP target folder path
        ''oFolder = oPath & "\" & oAsmName & " STEP Files"
        '''Check for the step folder and create it if it does not exist
        ''If Not System.IO.Directory.Exists(oFolder) Then
        ''    System.IO.Directory.CreateDirectory(oFolder)
        ''End If


        '''- - - - - - - - - - - - -Assembly - - - - - - - - - - - -
        ''oAsmDoc.Document.SaveAs(oFolder & "\" & oAsmName & (".stp"), True)

        ''- - - - - - - - - - - - -Components - - - - - - - - - - - -
        ''look at the files referenced by the assembly
        'Dim oRefDocs As DocumentsEnumerator
        'oRefDocs = oAsmDoc.AllReferencedDocuments
        'Dim oRefDoc As Document
        'Dim oRev = iProperties.GetorSetStandardiProperty(
        '                    AddinGlobal.InventorApp.ActiveDocument,
        '                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        ''work the referenced models
        'oDocu = AddinGlobal.InventorApp.ActiveDocument
        'oDocu.Save2(True)
        'For Each oRefDoc In oRefDocs
        '    If Not iPropertiesAddInServer.CheckReadOnly(oRefDoc) Then
        '        NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oRefDoc.FullDocumentName)

        '        'GetNewFilePaths()
        '        ' Get the STEP translator Add-In.

        '        Dim oSTEPTranslator As TranslatorAddIn
        '        oSTEPTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

        '        If oSTEPTranslator Is Nothing Then
        '            MsgBox("Could not access STEP translator.")
        '            Exit Sub
        '        End If

        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '        If oSTEPTranslator.HasSaveCopyAsOptions(oRefDoc, oContext, oOptions) Then

        '            ' Set application protocol.
        '            ' 2 = AP 203 - Configuration Controlled Design
        '            ' 3 = AP 214 - Automotive Design
        '            oOptions.Value("ApplicationProtocolType") = 3

        '            ' Other options...
        '            'oOptions.Value("Author") = ""
        '            'oOptions.Value("Authorization") = ""
        '            'oOptions.Value("Description") = ""
        '            'oOptions.Value("Organization") = ""

        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '            Dim oData As DataMedium
        '            oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '            oData.FileName = NewPath + "_R" + oRev + ".stp"

        '            Call oSTEPTranslator.SaveCopyAs(oRefDoc, oContext, oOptions, oData)
        '            UpdateStatusBar("File saved as Step file")

        '            AttachRefFile(oAsmDoc, oData.FileName)
        '            'AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
        '            'If AttachFile = vbYes Then
        '            '    AddReferences(inventorApp.ActiveDocument, oData.FileName)
        '            '    UpdateStatusBar("File attached")
        '            'Else
        '            '    'Do Nothing
        '            'End If
        '        End If
        '    Else
        '        UpdateStatusBar("File skipped because it's read-only")
        '    End If
        'Next
        'CloseIt()
    End Sub

    Private Sub btExpPdf_Click(sender As Object, e As EventArgs) Handles btExpPdf.Click
        iPropsForm.btExpPdf(sender, e)
        'Dim oDocu As Document = Nothing
        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '    CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '    If CheckRef = vbYes Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        'GetNewFilePaths()
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        If Not AddinGlobal.InventorApp.SoftwareVersion.Major > 20 Then
        '            AddinGlobal.InventorApp.StatusBarText = AddinGlobal.InventorApp.SoftwareVersion.Major
        '            MessageBox.Show("3D PDF export not available in Inventor versions < 2017 release!")
        '            Exit Sub
        '        End If
        '        ' Get the 3D PDF Add-In.
        '        Dim oPDFAddIn As ApplicationAddIn
        '        Dim oAddin As ApplicationAddIn
        '        For Each oAddin In AddinGlobal.InventorApp.ApplicationAddIns
        '            If oAddin.ClassIdString = "{3EE52B28-D6E0-4EA4-8AA6-C2A266DEBB88}" Then
        '                oPDFAddIn = oAddin
        '                Exit For
        '            End If
        '        Next

        '        If oPDFAddIn Is Nothing Then
        '            MsgBox("Inventor 3D PDF Addin not loaded.")
        '            Exit Sub
        '        End If

        '        Dim oPDFConvertor3D = oPDFAddIn.Automation

        '        'Set a reference to the active document (the document to be published).
        '        Dim oDocument As Document = AddinGlobal.InventorApp.ActiveDocument

        '        If oDocument.FileSaveCounter = 0 Then
        '            MsgBox("You must save the document to continue...")
        '            Return
        '        End If

        '        ' Create a NameValueMap objectfor all options...
        '        Dim oOptions As NameValueMap = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '        Dim STEPFileOptions As NameValueMap = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '        ' All Possible Options
        '        ' Export file name and location...
        '        oOptions.Value("FileOutputLocation") = NewPath + "_R" + oRev + ".pdf"
        '        ' Export annotations?
        '        oOptions.Value("ExportAnnotations") = 1
        '        ' Export work features?
        '        oOptions.Value("ExportWokFeatures") = 1
        '        ' Attach STEP file to 3D PDF?
        '        oOptions.Value("GenerateAndAttachSTEPFile") = True
        '        ' What quality (high quality takes longer to export)
        '        'oOptions.Value("VisualizationQuality") = AccuracyEnumVeryHigh
        '        oOptions.Value("VisualizationQuality") = AccuracyEnum.kHigh
        '        'oOptions.Value("VisualizationQuality") = AccuracyEnum.kMedium
        '        'oOptions.Value("VisualizationQuality") = AccuracyEnum.kLow
        '        ' Limit export to entities in selected view representation(s)
        '        oOptions.Value("LimitToEntitiesInDVRs") = True
        '        ' Open the 3D PDF when export is complete?
        '        oOptions.Value("ViewPDFWhenFinished") = False

        '        ' Export all properties?
        '        oOptions.Value("ExportAllProperties") = True
        '        ' OR - Set the specific properties to export
        '        '    Dim sProps(5) As String
        '        '    sProps(0) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Title"
        '        '    sProps(1) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Keywords"
        '        '    sProps(2) = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}:Comments"
        '        '    sProps(3) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Description"
        '        '    sProps(4) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Stock Number"
        '        '    sProps(5) =    "{32853F0F-3444-11D1-9E93-0060B03C1CA6}:Revision Number"

        '        'oOptions.Value("ExportProperties") = sProps

        '        ' Choose the export template based off the current document type
        '        If oDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '            oOptions.Value("ExportTemplate") = "C:\Users\Public\Documents\Autodesk\Inventor 2017\Templates\Sample Part Template.pdf"
        '        Else
        '            oOptions.Value("ExportTemplate") = "C:\Users\Public\Documents\Autodesk\Inventor 2017\Templates\Sample Assembly Template.pdf"
        '        End If

        '        ' Define a file to attach to the exported 3D PDF - note here I have picked an Excel spreadsheet
        '        ' You need to use the full path and filename - if it does not exist the file will not be attached.
        '        Dim oAttachedFiles As String() = {"C:\FileToAttach.xlsx"}
        '        oOptions.Value("AttachedFiles") = oAttachedFiles

        '        ' Set the design view(s) to export - note here I am exporting only the active design view (view representation)
        '        Dim sDesignViews(0) As String
        '        sDesignViews(0) = oDocument.ComponentDefinition.RepresentationsManager.ActiveDesignViewRepresentation.Name
        '        oOptions.Value("ExportDesignViewRepresentations") = sDesignViews

        '        'Publish document.
        '        Call oPDFConvertor3D.Publish(oDocument, oOptions)
        '        UpdateStatusBar("File saved as 3D pdf file")
        '        AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oOptions.Value("FileOutputLocation"))
        '    Else
        '        CloseIt()
        '        'Do Nothing
        '    End If
        'Else
        '    If Not iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        'GetNewFilePaths()
        '        ' Get the PDF translator Add-In.
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                            AddinGlobal.InventorApp.ActiveDocument,
        '                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        Dim PDFAddIn As TranslatorAddIn
        '        PDFAddIn = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

        '        'Set a reference to the active document (the document to be published).
        '        Dim oDocument As Document
        '        oDocument = AddinGlobal.InventorApp.ActiveDocument

        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '        ' Create a NameValueMap object
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '        ' Create a DataMedium object
        '        Dim oDataMedium As DataMedium
        '        oDataMedium = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium

        '        ' Check whether the translator has 'SaveCopyAs' options
        '        If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

        '            ' Options for drawings...

        '            oOptions.Value("All_Color_AS_Black") = 1

        '            'oOptions.Value("Remove_Line_Weights") = 0
        '            'oOptions.Value("Vector_Resolution") = 400
        '            'oOptions.Value("Sheet_Range") = kPrintAllSheets
        '            'oOptions.Value("Custom_Begin_Sheet") = 2
        '            'oOptions.Value("Custom_End_Sheet") = 4

        '        End If

        '        'Set the destination file name
        '        oDataMedium.FileName = NewPath + "_R" + oRev + ".pdf"

        '        'Publish document.
        '        Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
        '        UpdateStatusBar("File saved as pdf file")
        '        AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oDataMedium.FileName)
        '    Else
        '        CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '        If CheckRef = vbYes Then
        '            oDocu = AddinGlobal.InventorApp.ActiveDocument
        '            oDocu.Save2(True)
        '            'GetNewFilePaths()
        '            ' Get the PDF translator Add-In.
        '            Dim oRev = iProperties.GetorSetStandardiProperty(
        '                                AddinGlobal.InventorApp.ActiveDocument,
        '                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '            Dim PDFAddIn As TranslatorAddIn
        '            PDFAddIn = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

        '            'Set a reference to the active document (the document to be published).
        '            Dim oDocument As Document
        '            oDocument = AddinGlobal.InventorApp.ActiveDocument

        '            Dim oContext As TranslationContext
        '            oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '            ' Create a NameValueMap object
        '            Dim oOptions As NameValueMap
        '            oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '            ' Create a DataMedium object
        '            Dim oDataMedium As DataMedium
        '            oDataMedium = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium

        '            ' Check whether the translator has 'SaveCopyAs' options
        '            If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

        '                ' Options for drawings...

        '                oOptions.Value("All_Color_AS_Black") = 1

        '                'oOptions.Value("Remove_Line_Weights") = 0
        '                'oOptions.Value("Vector_Resolution") = 400
        '                'oOptions.Value("Sheet_Range") = kPrintAllSheets
        '                'oOptions.Value("Custom_Begin_Sheet") = 2
        '                'oOptions.Value("Custom_End_Sheet") = 4

        '            End If

        '            'Set the destination file name
        '            oDataMedium.FileName = NewPath + "_R" + oRev + ".pdf"

        '            'Publish document.
        '            Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
        '            UpdateStatusBar("File saved as pdf file")
        '            AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oDataMedium.FileName)
        '        Else
        '            'Do Nothing
        '        End If
        '    End If
        'End If
        'CloseIt()
    End Sub

    Private Sub btExpStp_Click(sender As Object, e As EventArgs) Handles btExpStp.Click
        iPropsForm.btExpStp(sender, e)
        'Dim oDocu As Document = Nothing
        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '    CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '    If CheckRef = vbYes Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        'GetNewFilePaths()
        '        ' Get the STEP translator Add-In.
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        Dim oSTEPTranslator As TranslatorAddIn
        '        oSTEPTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

        '        If oSTEPTranslator Is Nothing Then
        '            MsgBox("Could not access STEP translator.")
        '            Exit Sub
        '        End If

        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '        If oSTEPTranslator.HasSaveCopyAsOptions(AddinGlobal.InventorApp.ActiveDocument, oContext, oOptions) Then
        '            ' Set application protocol.
        '            ' 2 = AP 203 - Configuration Controlled Design
        '            ' 3 = AP 214 - Automotive Design
        '            oOptions.Value("ApplicationProtocolType") = 3

        '            ' Other options...
        '            'oOptions.Value("Author") = ""
        '            'oOptions.Value("Authorization") = ""
        '            'oOptions.Value("Description") = ""
        '            'oOptions.Value("Organization") = ""

        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '            Dim oData As DataMedium
        '            oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '            oData.FileName = NewPath + "_R" + oRev + ".stp"

        '            Call oSTEPTranslator.SaveCopyAs(AddinGlobal.InventorApp.ActiveDocument, oContext, oOptions, oData)
        '            UpdateStatusBar("File saved as Step file")

        '            AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oData.FileName)
        '            'AttachFile = MsgBox("File exported, attach it to main file as reference?", vbYesNo, "File Attach")
        '            'If AttachFile = vbYes Then
        '            '    AddReferences(AddinGlobal.inventorApp.ActiveDocument, oData.FileName)
        '            '    UpdateStatusBar("File attached")
        '            'Else
        '            '    'Do Nothing
        '            'End If
        '        End If
        '    Else
        '        'Do nothing
        '    End If
        'ElseIf AddinGlobal.inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        '    If Not iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        'GetNewFilePaths()
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                    RefDoc,
        '                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
        '            UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
        '        Else

        '            ' Get the STEP translator Add-In.
        '            Dim oSTEPTranslator As TranslatorAddIn
        '            oSTEPTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

        '            If oSTEPTranslator Is Nothing Then
        '                MsgBox("Could not access STEP translator.")
        '                Exit Sub
        '            End If

        '            Dim oContext As TranslationContext
        '            oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '            Dim oOptions As NameValueMap
        '            oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '            If oSTEPTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
        '                ' Set application protocol.
        '                ' 2 = AP 203 - Configuration Controlled Design
        '                ' 3 = AP 214 - Automotive Design
        '                oOptions.Value("ApplicationProtocolType") = 3

        '                ' Other options...
        '                'oOptions.Value("Author") = ""
        '                'oOptions.Value("Authorization") = ""
        '                'oOptions.Value("Description") = ""
        '                'oOptions.Value("Organization") = ""

        '                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '                Dim oData As DataMedium
        '                oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '                oData.FileName = RefNewPath + "_R" + oRev + ".stp"

        '                Call oSTEPTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '                UpdateStatusBar("Part/Assy file saved as Step file")
        '                AttachRefFile(RefDoc, oData.FileName)
        '            End If
        '        End If
        '    Else
        '        CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '        If CheckRef = vbYes Then
        '            oDocu = AddinGlobal.InventorApp.ActiveDocument
        '            oDocu.Save2(True)
        '            'GetNewFilePaths()
        '            Dim oRev = iProperties.GetorSetStandardiProperty(
        '                        RefDoc,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '            If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
        '                UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
        '            Else

        '                ' Get the STEP translator Add-In.
        '                Dim oSTEPTranslator As TranslatorAddIn
        '                oSTEPTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

        '                If oSTEPTranslator Is Nothing Then
        '                    MsgBox("Could not access STEP translator.")
        '                    Exit Sub
        '                End If

        '                Dim oContext As TranslationContext
        '                oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '                Dim oOptions As NameValueMap
        '                oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '                If oSTEPTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
        '                    ' Set application protocol.
        '                    ' 2 = AP 203 - Configuration Controlled Design
        '                    ' 3 = AP 214 - Automotive Design
        '                    oOptions.Value("ApplicationProtocolType") = 3

        '                    ' Other options...
        '                    'oOptions.Value("Author") = ""
        '                    'oOptions.Value("Authorization") = ""
        '                    'oOptions.Value("Description") = ""
        '                    'oOptions.Value("Organization") = ""

        '                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '                    Dim oData As DataMedium
        '                    oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '                    oData.FileName = RefNewPath + "_R" + oRev + ".stp"

        '                    Call oSTEPTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '                    UpdateStatusBar("Part/Assy file saved as Step file")
        '                    AttachRefFile(RefDoc, oData.FileName)
        '                End If
        '            End If
        '        Else
        '            CloseIt()
        '            'Do Nothing
        '        End If
        '    End If

        'End If
        'CloseIt()
    End Sub

    Private Sub btExpSat_Click(sender As Object, e As EventArgs) Handles btExpSat.Click
        iPropsForm.btExpSat(sender, e)
        'Dim oDocu As Document = Nothing
        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '    CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '    If CheckRef = vbYes Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        '        ' Get the SAT translator Add-In.
        '        Dim oSATTrans As TranslatorAddIn
        '        oSATTrans = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
        '        If oSATTrans Is Nothing Then
        '            MsgBox("Could not access SAT translator.")
        '            Exit Sub
        '        End If
        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '        If oSATTrans.HasSaveCopyAsOptions(AddinGlobal.InventorApp.ActiveDocument, oContext, oOptions) Then
        '            oOptions.Value("ExportUnits") = 5
        '            oOptions.Value("IncludeSketches") = 0
        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        '            Dim oData As DataMedium
        '            oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '            oData.FileName = NewPath + "_R" + oRev + ".sat"
        '            Call oSATTrans.SaveCopyAs(AddinGlobal.InventorApp.ActiveDocument, oContext, oOptions, oData)
        '            UpdateStatusBar("File saved as Sat file")
        '            AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oData.FileName)
        '        End If
        '    Else
        '        'Do Nothing
        '    End If
        'ElseIf AddinGlobal.inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        '    If Not iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                        RefDoc,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        '        ' Get the SAT translator Add-In.
        '        Dim oSATTrans As TranslatorAddIn
        '        oSATTrans = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
        '        If oSATTrans Is Nothing Then
        '            MsgBox("Could not access SAT translator.")
        '            Exit Sub
        '        End If
        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '        If oSATTrans.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
        '            oOptions.Value("ExportUnits") = 5
        '            oOptions.Value("IncludeSketches") = 0
        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        '            Dim oData As DataMedium
        '            oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '            oData.FileName = RefNewPath + "_R" + oRev + ".sat"
        '            Call oSATTrans.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '            UpdateStatusBar("Part/Assy file saved as Sat file")
        '            AttachRefFile(RefDoc, oData.FileName)
        '        End If
        '    Else
        '        CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '        If CheckRef = vbYes Then
        '            oDocu = AddinGlobal.InventorApp.ActiveDocument
        '            oDocu.Save2(True)
        '            Dim oRev = iProperties.GetorSetStandardiProperty(
        '                            RefDoc,
        '                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
        '            ' Get the SAT translator Add-In.
        '            Dim oSATTrans As TranslatorAddIn
        '            oSATTrans = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{89162634-02B6-11D5-8E80-0010B541CD80}")
        '            If oSATTrans Is Nothing Then
        '                MsgBox("Could not access SAT translator.")
        '                Exit Sub
        '            End If
        '            Dim oContext As TranslationContext
        '            oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '            Dim oOptions As NameValueMap
        '            oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap
        '            If oSATTrans.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then
        '                oOptions.Value("ExportUnits") = 5
        '                oOptions.Value("IncludeSketches") = 0
        '                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        '                Dim oData As DataMedium
        '                oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '                oData.FileName = RefNewPath + "_R" + oRev + ".sat"
        '                Call oSATTrans.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '                UpdateStatusBar("Part/Assy file saved as Sat file")
        '                AttachRefFile(RefDoc, oData.FileName)
        '            End If
        '        Else
        '            CloseIt()
        '            'Do Nothing
        '        End If
        '    End If
        'End If
        'CloseIt()
    End Sub

    Private Sub btExpStl_Click(sender As Object, e As EventArgs) Handles btExpStl.Click
        iPropsForm.btExpStl(sender, e)
        'Dim oDocu As Document = Nothing
        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '    CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '    If CheckRef = vbYes Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        ' Get the STL translator Add-In.
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                            AddinGlobal.InventorApp.ActiveDocument,
        '                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        Dim oSTLTranslator As TranslatorAddIn
        '        oSTLTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
        '        If oSTLTranslator Is Nothing Then
        '            MsgBox("Could not access STL translator.")
        '            Exit Sub
        '        End If

        '        Dim oDoc = AddinGlobal.InventorApp.ActiveDocument

        '        Dim oContext As TranslationContext
        '        oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext

        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '        'Save Copy As Options
        '        'Name value Map
        '        'ExportUnits = 4
        '        'Resolution = 1
        '        'AllowMoveMeshNode = False
        '        'SurfaceDeviation = 60
        '        'NormalDeviation = 14
        '        'MaxEdgeLength = 100
        '        'AspectRatio = 40
        '        'ExportFileStructure = 0
        '        'OutputFileType = 0
        '        'ExportColor = True

        '        If oSTLTranslator.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

        '            oOptions.Value("ExportUnits") = 5
        '            ' Set output file type:
        '            '   0 - binary,  1 - ASCII
        '            oOptions.Value("OutputFileType") = 0
        '            ' Set accuracy.
        '            '   2 = Low,  1 = Medium,  0 = High
        '            oOptions.Value("Resolution") = 0
        '            'oOptions.Value("SurfaceDeviation") = 0.005
        '            'oOptions.Value("NormalDeviation") = 10
        '            'oOptions.Value("MaxEdgeLength") = 100
        '            'oOptions.Value("AspectRatio") = 21.5
        '            oOptions.Value("ExportColor") = True

        '            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '            Dim oData As DataMedium
        '            oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '            oData.FileName = NewPath + "_R" + oRev + ".stl"

        '            Call oSTLTranslator.SaveCopyAs(oDoc, oContext, oOptions, oData)
        '            UpdateStatusBar("File saved as STL file")
        '            AttachRefFile(AddinGlobal.InventorApp.ActiveDocument, oData.FileName)

        '            ' The various names and values for the settings is described below.
        '            'ExportUnits
        '            '     2 - Inch
        '            '     3 - Foot
        '            '     4 - Centimeter
        '            '     5 - Millimeter
        '            '     6 - Meter
        '            '     7 - Micron
        '            '
        '            'AllowMoveMeshNode (True Or False)
        '            '
        '            'ExportColor (True or False)  (Only used for Binary output file type)
        '            '
        '            'OutputFileType
        '            '     0 - Binary
        '            '     1 - ASCII
        '            'ExportFileStructure (Only used for assemblies)
        '            '     0 - One file
        '            '     1 - One file per instance.
        '            '
        '            'Resolution
        '            '     0 - High
        '            '     1 - Medium
        '            '     2 - Low
        '            '     3 - Custom
        '            '
        '            '*** The following are only used for “Custom” resolution
        '            '
        '            'SurfaceDeviation
        '            '     0 to 100 for a percentage between values computed
        '            '     based on the size of the model.
        '            'NormalDeviation
        '            '     0 to 40 to indicate values 1 to 41
        '            'MaxEdgeLength
        '            '     0 to 100 for a percentage between values computed
        '            '     based on the size of the model.
        '            'AspectRatio
        '            '     0 to 40 for values between 1.5 to 21.5 in 0.5 increments
        '        End If
        '    Else
        '        'Do Nothing
        '    End If
        'ElseIf AddinGlobal.inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        '    If Not iProperties.GetorSetStandardiProperty(
        '                        AddinGlobal.InventorApp.ActiveDocument,
        '                        PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "") = "Revision Number" Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        'GetNewFilePaths()
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                            RefDoc,
        '                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
        '            UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
        '        Else
        '            ' Get the STL translator Add-In.
        '            Dim oSTLTranslator As TranslatorAddIn
        '            oSTLTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
        '            If oSTLTranslator Is Nothing Then
        '                MsgBox("Could not access STL translator.")
        '                Exit Sub
        '            End If

        '            Dim oContext As TranslationContext
        '            oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext

        '            Dim oOptions As NameValueMap
        '            oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '            '    Save Copy As Options:
        '            '       Name Value Map:
        '            '               ExportUnits = 4
        '            '               Resolution = 1
        '            '               AllowMoveMeshNode = False
        '            '               SurfaceDeviation = 60
        '            '               NormalDeviation = 14
        '            '               MaxEdgeLength = 100
        '            '               AspectRatio = 40
        '            '               ExportFileStructure = 0
        '            '               OutputFileType = 0
        '            '               ExportColor = True

        '            If oSTLTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then

        '                oOptions.Value("ExportUnits") = 5
        '                ' Set output file type:
        '                '   0 - binary,  1 - ASCII
        '                oOptions.Value("OutputFileType") = 0
        '                ' Set accuracy.
        '                '   2 = Low,  1 = Medium,  0 = High
        '                oOptions.Value("Resolution") = 0
        '                'oOptions.Value("SurfaceDeviation") = 0.005
        '                'oOptions.Value("NormalDeviation") = 10
        '                'oOptions.Value("MaxEdgeLength") = 100
        '                'oOptions.Value("AspectRatio") = 21.5
        '                oOptions.Value("ExportColor") = True

        '                oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '                Dim oData As DataMedium
        '                oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '                oData.FileName = RefNewPath + "_R" + oRev + ".stl"

        '                Call oSTLTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '                UpdateStatusBar("Part/Assy file saved as STL file")
        '                AttachRefFile(RefDoc, oData.FileName)
        '            End If
        '        End If
        '    Else
        '        CheckRef = MsgBox("Have you checked the model revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '        If CheckRef = vbYes Then
        '            oDocu = AddinGlobal.InventorApp.ActiveDocument
        '            oDocu.Save2(True)
        '            'GetNewFilePaths()
        '            Dim oRev = iProperties.GetorSetStandardiProperty(
        '                            RefDoc,
        '                            PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '            If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
        '                UpdateStatusBar("Cannot export model whilst drawing updates are deferred")
        '            Else
        '                ' Get the STL translator Add-In.
        '                Dim oSTLTranslator As TranslatorAddIn
        '                oSTLTranslator = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")
        '                If oSTLTranslator Is Nothing Then
        '                    MsgBox("Could not access STL translator.")
        '                    Exit Sub
        '                End If

        '                Dim oContext As TranslationContext
        '                oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext

        '                Dim oOptions As NameValueMap
        '                oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '                '    Save Copy As Options:
        '                '       Name Value Map:
        '                '               ExportUnits = 4
        '                '               Resolution = 1
        '                '               AllowMoveMeshNode = False
        '                '               SurfaceDeviation = 60
        '                '               NormalDeviation = 14
        '                '               MaxEdgeLength = 100
        '                '               AspectRatio = 40
        '                '               ExportFileStructure = 0
        '                '               OutputFileType = 0
        '                '               ExportColor = True

        '                If oSTLTranslator.HasSaveCopyAsOptions(RefDoc, oContext, oOptions) Then

        '                    oOptions.Value("ExportUnits") = 5
        '                    ' Set output file type:
        '                    '   0 - binary,  1 - ASCII
        '                    oOptions.Value("OutputFileType") = 0
        '                    ' Set accuracy.
        '                    '   2 = Low,  1 = Medium,  0 = High
        '                    oOptions.Value("Resolution") = 0
        '                    'oOptions.Value("SurfaceDeviation") = 0.005
        '                    'oOptions.Value("NormalDeviation") = 10
        '                    'oOptions.Value("MaxEdgeLength") = 100
        '                    'oOptions.Value("AspectRatio") = 21.5
        '                    oOptions.Value("ExportColor") = True

        '                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '                    Dim oData As DataMedium
        '                    oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '                    oData.FileName = RefNewPath + "_R" + oRev + ".stl"

        '                    Call oSTLTranslator.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '                    UpdateStatusBar("Part/Assy file saved as STL file")
        '                    AttachRefFile(RefDoc, oData.FileName)
        '                End If
        '            End If
        '        Else
        '            CloseIt()
        '            'Do Nothing
        '        End If
        '    End If
        'End If
        'CloseIt()
    End Sub

    Private Sub btExpDWF_Click(sender As Object, e As EventArgs) Handles btExpDWF.Click
        iPropertiesAddInServer.btExpDWF(sender, e)

        'Dim oDocu As Document = Nothing

        'If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        '    CheckRef = MsgBox("Have you checked the revision number matches the drawing revision?", vbYesNo, "Rev. Check")
        '    If CheckRef = vbYes Then
        '        oDocu = AddinGlobal.InventorApp.ActiveDocument
        '        oDocu.Save2(True)
        '        ' Get the STL translator Add-In.
        '        Dim oRev = iProperties.GetorSetStandardiProperty(
        '                                AddinGlobal.InventorApp.ActiveDocument,
        '                                PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

        '        ' Get the DWF translator Add-In.
        '        Dim DWFAddIn As TranslatorAddIn
        '        DWFAddIn = AddinGlobal.InventorApp.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        '        'Set a reference to the active document (the document to be published).
        '        Dim oDoc = AddinGlobal.InventorApp.ActiveDocument

        '        Dim oContext = AddinGlobal.InventorApp.TransientObjects.CreateTranslationContext
        '        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        '        ' Create a NameValueMap object
        '        Dim oOptions As NameValueMap
        '        oOptions = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '        ' Create a DataMedium object
        '        Dim oDataMedium As DataMedium
        '        oDataMedium = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium

        '        ' Check whether the translator has 'SaveCopyAs' options
        '        If DWFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then

        '            oOptions.Value("Launch_Viewer") = 1

        '            ' Other options...
        '            'oOptions.Value("Publish_All_Component_Props") = 1
        '            'oOptions.Value("Publish_All_Physical_Props") = 1
        '            'oOptions.Value("Password") = 0

        '            If TypeOf oDoc Is DrawingDocument Then

        '                ' Drawing options
        '                oOptions.Value("Publish_Mode") = kCustomDWFPublish
        '                oOptions.Value("Publish_All_Sheets") = 1

        '                ' The specified sheets will be ignored if
        '                ' the option "Publish_All_Sheets" is True (1)
        '                Dim oSheets As NameValueMap
        '                oSheets = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '                ' Publish the first sheet AND its 3D model
        '                Dim oSheet1Options As NameValueMap
        '                oSheet1Options = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '                oSheet1Options.Add("Name", "Sheet:1")
        '                oSheet1Options.Add("3DModel", True)
        '                oSheets.Value("Sheet1") = oSheet1Options

        '                ' Publish the third sheet but NOT its 3D model
        '                Dim oSheet3Options As NameValueMap
        '                oSheet3Options = AddinGlobal.InventorApp.TransientObjects.CreateNameValueMap

        '                oSheet3Options.Add("Name", "Sheet:3")
        '                oSheet3Options.Add("3DModel", False)

        '                oSheets.Value("Sheet2") = oSheet3Options

        '                'Set the sheet options object in the oOptions NameValueMap
        '                oOptions.Value("Sheets") = oSheets
        '            End If

        '        End If

        '        Dim oData As DataMedium
        '        oData = AddinGlobal.InventorApp.TransientObjects.CreateDataMedium
        '        oData.FileName = RefNewPath + "_R" + oRev + ".dwf"

        '        DWFAddIn.SaveCopyAs(RefDoc, oContext, oOptions, oData)
        '        UpdateStatusBar("Part/Assy file saved as DWF file")
        '    Else
        '        CloseIt()
        '    End If

        'End If
        'CloseIt()
    End Sub

    Public Sub CloseIt()
        Me.Visible = False
    End Sub
End Class