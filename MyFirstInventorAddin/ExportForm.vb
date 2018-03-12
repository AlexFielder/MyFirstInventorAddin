Imports System.Windows.Forms
Imports Inventor
Imports iPropertiesController.iPropertiesController
Imports log4net

Public Class ExportForm
    Public Shared myiPropsForm As IPropertiesForm = Nothing
    Public inventorApp As Inventor.Application = AddinGlobal.InventorApp

    Public Sub GetNewFilePaths()
        If Not inventorApp.ActiveDocument Is Nothing Then
            If inventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                If inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    CurrentPath = System.IO.Path.GetDirectoryName(inventorApp.ActiveDocument.FullDocumentName)
                    NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(inventorApp.ActiveDocument.FullDocumentName)

                ElseIf inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    If iProperties.GetorSetStandardiProperty(inventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        'Do Nothing
                    Else
                        Dim oDrawDoc As DrawingDocument = inventorApp.ActiveDocument
                        Dim oSht As Sheet = oDrawDoc.ActiveSheet
                        Dim oView As DrawingView = Nothing

                        For Each view As DrawingView In oSht.DrawingViews
                            oView = view
                            Exit For
                        Next

                        CurrentPath = System.IO.Path.GetDirectoryName(inventorApp.ActiveDocument.FullDocumentName)
                        RefCurrentPath = System.IO.Path.GetDirectoryName(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                        NewPath = CurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(inventorApp.ActiveDocument.FullDocumentName)

                        RefDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                        RefNewPath = RefCurrentPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)
                    End If
                End If
            End If
        End If
    End Sub

    Public CurrentPath As String = String.Empty
    Public NewPath As String = String.Empty
    Public RefNewPath As String = String.Empty
    Public RefDoc As Document = Nothing

    Public Sub CloseIt()
        Me.Visible = False
    End Sub

    Private Sub UpdateStatusBar(ByVal Message As String)
        inventorApp.StatusBarText = Message
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

    Public Sub AddReferences(ByVal odoc As Inventor.Document, ByVal selectedfile As String)
        Dim oleReference As ReferencedOLEFileDescriptor
        oleReference = odoc.ReferencedOLEFileDescriptors _
                    .Add(selectedfile, OLEDocumentTypeEnum.kOLEDocumentLinkObject)
        oleReference.BrowserVisible = True
        oleReference.Visible = False
        oleReference.DisplayName = System.IO.Path.GetFileName(selectedfile)
    End Sub

    Private Sub ExpStp_Click(sender As Object, e As EventArgs) Handles ExpStp.Click
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
        CloseIt()
    End Sub
End Class