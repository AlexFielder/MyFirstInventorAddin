Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports log4net

Namespace iPropertiesController

    <ProgIdAttribute("iPropertiesController.StandardAddInServer"),
    GuidAttribute("e691af34-cb32-4296-8cca-5d1027a27c72")>
    Public Class iPropertiesAddInServer
        Implements Inventor.ApplicationAddInServer

        'some events objects we might need later
        Private WithEvents m_uiEvents As UserInterfaceEvents

        Private WithEvents m_UserInputEvents As UserInputEvents
        Private WithEvents m_AppEvents As ApplicationEvents

        ' new unused - at this point - event objects
        Private WithEvents m_DocEvents As DocumentEvents

        Private WithEvents m_AssemblyEvents As AssemblyEvents
        Private WithEvents m_PartEvents As PartEvents
        Private WithEvents m_ModelingEvents As ModelingEvents
        Private WithEvents m_StyleEvents As StyleEvents

        Private WithEvents m_FileAccesEvents As FileAccessEvents

        Private thisAssembly As Assembly = Assembly.GetExecutingAssembly()
        Private thisAssemblyPath As String = String.Empty
        Public thisVersion As Version = Nothing
        Public Shared attribute As GuidAttribute = Nothing
        Public Shared myiPropsForm As IPropertiesForm = Nothing
        Public Property InventorAppQuitting As Boolean = False

        'we can set the following to false if we don't want the file to save:
        Public AllowFileToSave As Boolean = True
        Public AllowFileToSaveAs As Boolean = True

        Private logHelper As Log4NetFileHelper.Log4NetFileHelper = New Log4NetFileHelper.Log4NetFileHelper()
        Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(iPropertiesAddInServer))

        'Private WithEvents m_sampleButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) Implements ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            AddinGlobal.InventorApp = addInSiteObject.Application
            'new versioning display method borrowed from here: https://stackoverflow.com/a/826850/572634
            thisVersion = Assembly.GetExecutingAssembly().GetName().Version
            Dim buildDate As DateTime = New DateTime(2000, 1, 1).AddDays(thisVersion.Build).AddSeconds(thisVersion.Revision * 2)
            AddinGlobal.DisplayableVersion = $"{thisVersion}"

            Dim uiMgr As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
            attribute = DirectCast(thisAssembly.GetCustomAttributes(GetType(GuidAttribute), True)(0), GuidAttribute)
            Try

                AddinGlobal.GetAddinClassId(Me.GetType())
                'store our Addin path.
                thisAssemblyPath = IO.Path.GetDirectoryName(thisAssembly.Location)
                ' Connect to the user-interface events to handle a ribbon reset.
                m_uiEvents = AddinGlobal.InventorApp.UserInterfaceManager.UserInterfaceEvents
                'Connect to the Application Events to handle document opening/switching for our iProperties dockable Window.
                m_AppEvents = AddinGlobal.InventorApp.ApplicationEvents
                m_UserInputEvents = AddinGlobal.InventorApp.CommandManager.UserInputEvents
                m_StyleEvents = AddinGlobal.InventorApp.StyleEvents

                AddHandler m_AppEvents.OnOpenDocument, AddressOf Me.m_ApplicationEvents_OnOpenDocument
                AddHandler m_AppEvents.OnActivateDocument, AddressOf Me.m_ApplicationEvents_OnActivateDocument
                AddHandler m_AppEvents.OnSaveDocument, AddressOf Me.m_ApplicationEvents_OnSaveDocument
                AddHandler m_AppEvents.OnQuit, AddressOf Me.m_ApplicationEvents_OnQuit
                AddHandler m_AppEvents.OnActivateView, AddressOf Me.m_ApplicationEvents_OnActivateView

                AddHandler m_UserInputEvents.OnActivateCommand, AddressOf Me.m_UserInputEvents_OnActivateCommand
                AddHandler m_UserInputEvents.OnTerminateCommand, AddressOf Me.m_UserInputEvents_OnTerminateCommand
                'you can add extra handlers like this - if you uncomment the next line Visual Studio will prompt you to create the method:
                'AddHandler m_AssemblyEvents.OnNewOccurrence, AddressOf Me.m_AssemblyEvents_NewOcccurrence
                If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                    m_DocEvents = AddinGlobal.InventorApp.ActiveDocument.DocumentEvents
                    AddHandler m_DocEvents.OnChangeSelectSet, AddressOf Me.m_DocumentEvents_OnChangeSelectSet
                End If

                AddHandler m_StyleEvents.OnActivateStyle, AddressOf Me.m_StyleEvents_OnActivateStyle

                AddHandler m_AppEvents.OnNewEditObject, AddressOf Me.m_ApplicationEvents_OnNewEditObject

                AddHandler m_AppEvents.OnCloseDocument, AddressOf Me.m_ApplicationEvents_OnCloseDocument

                'start our logger.
                logHelper.Init()
                logHelper.AddFileLogging(IO.Path.Combine(thisAssemblyPath, "iPropertiesController.log"))
                logHelper.AddFileLogging("C:\Logs\MyLogFile.txt", Core.Level.All, True)
                logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", Core.Level.All, True)
                log.Debug("Loading My First Inventor Addin")
                ' TODO: Add button definitions.

                ' Sample to illustrate creating a button definition.
                'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
                'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
                'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
                'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

                Dim icon1 As New Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("iPropertiesController.addin.ico"))
                'Change it if necessary but make sure it's embedded.
                Dim button1 As New InventorButton("Button 1", "MyVBInventorAddin.Button_" & Guid.NewGuid().ToString(), "Button 1 description", "Button 1 tooltip", icon1, icon1,
                    CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
                button1.SetBehavior(True, True, True)
                button1.Execute = AddressOf ButtonActions.Button1_Execute

                ' Add to the user interface, if it's the first time.
                If firstTime Then
                    AddToUserInterface(button1)
                    'add our userform to a new DockableWindow
                    Dim localWindow As DockableWindow = Nothing
                    myiPropsForm = New IPropertiesForm(AddinGlobal.InventorApp)
                    'deal with Inventor's Dark theme:
                    If IsInventorUsingDarkTheme() Then
                        SwitchTheme(myiPropsForm, True)
                    Else
                        SwitchTheme(myiPropsForm)
                    End If
                    'custom sizing
                    myiPropsForm.tbDrawnBy.Width = (myiPropsForm.Size.Width * 0.7) - 2 * myiPropsForm.customMargin - myiPropsForm.tbDrawnBy.Location.X
                    myiPropsForm.Show()
                    Window = uiMgr.DockableWindows.Add(attribute.Value, "iPropertiesControllerWindow", "iProperties Controller " + AddinGlobal.DisplayableVersion)
                    Window.AddChild(myiPropsForm.Handle)

                    'If Not Window.IsCustomized = True Then
                    '    'myDockableWindow.DockingState = DockingStateEnum.kFloat
                    '    Window.DockingState = DockingStateEnum.kDockLastKnown
                    'Else
                    '    Window.DockingState = DockingStateEnum.kFloat
                    'End If

                    Window.DisabledDockingStates = DockingStateEnum.kDockTop + DockingStateEnum.kDockBottom
                    Window.ShowVisibilityCheckBox = True
                    Window.ShowTitleBar = True
                    Window.SetMinimumSize(440, 285)
                    myiPropsForm.Dock = DockStyle.Fill
                    Window.Visible = True
                    'localWindow = myDockableWindow
                    AddinGlobal.DockableList.Add(Window)
                    'Window = localWindow

                End If
                log.Info("Loaded My First Inventor Add-in")
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
        End Sub

        Private Sub SwitchTheme(ByRef myiPropsForm As IPropertiesForm, Optional DarkTheme As Boolean = False)
            If DarkTheme Then
                AddinGlobal.BackColour = Drawing.Color.FromArgb(59, 68, 83)
                AddinGlobal.ForeColour = Drawing.Color.FromArgb(255, 255, 255)
                AddinGlobal.ControlBackColour = Drawing.Color.FromArgb(69, 79, 97)
                AddinGlobal.ControlHighlightedColour = Drawing.Color.FromArgb(44, 52, 64)
            Else
                AddinGlobal.BackColour = Drawing.Color.FromArgb(240, 240, 240)
                AddinGlobal.ForeColour = Drawing.Color.FromArgb(0, 0, 0)
                AddinGlobal.ControlBackColour = Drawing.Color.FromArgb(255, 255, 255)
                AddinGlobal.ControlHighlightedColour = Drawing.Color.FromArgb(225, 255, 255)
            End If
            myiPropsForm.BackColor = AddinGlobal.BackColour
            For Each FormControl As Control In myiPropsForm.Controls
                Select Case TypeName(FormControl)
                    Case "Button"
                        Dim Btn As Button = FormControl
                        If Not Btn.Name = btDefer Then
                            Btn.BackColor = AddinGlobal.ControlBackColour
                            Btn.ForeColor = AddinGlobal.ForeColour
                            'Btn.FlatAppearance.BorderColor = AddinGlobal.ForeColour
                        End If
                    Case "Label"
                        Dim Lbl As Label = FormControl
                        Lbl.BackColor = AddinGlobal.BackColour
                        Lbl.ForeColor = AddinGlobal.ForeColour
                    Case "TextBox"
                        Dim TxtBox As Windows.Forms.TextBox = FormControl
                        TxtBox.BackColor = AddinGlobal.ControlBackColour
                        TxtBox.ForeColor = AddinGlobal.ForeColour
                    Case "DateTimePicker"
                        Dim DateTimePickR As DateTimePicker = FormControl
                        DateTimePickR.CalendarMonthBackground = AddinGlobal.ControlBackColour
                        DateTimePickR.CalendarForeColor = AddinGlobal.ForeColour
                End Select

            Next
        End Sub

        Private Function IsInventorUsingDarkTheme() As Boolean
            Dim oThemeManager As ThemeManager = AddinGlobal.InventorApp.ThemeManager
            If oThemeManager.Themes.Item(1).Name = oThemeManager.ActiveTheme.Name Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Sub m_UserInputEvents_OnTerminateCommand(CommandName As String, Context As NameValueMap)
            Dim oDoc As Document = AddinGlobal.InventorApp.ActiveDocument
            If TypeOf oDoc Is DrawingDocument Then
                Dim oDWG = AddinGlobal.InventorApp.ActiveDocument
                DocumentToPulliPropValuesFrom = oDWG
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing
                Dim drawnDoc As Document = Nothing
                Dim MaterialString As String = String.Empty
                Dim DrawDesc As String = iProperties.GetorSetStandardiProperty(oDWG, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")

                If DrawDesc = String.Empty Then

                    If CommandName = "DrawingBaseViewCmd" Then

                        For Each view As DrawingView In oSht.DrawingViews
                            oView = view
                            Exit For
                        Next

                        If oView IsNot Nothing Then

                            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

                            myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                            myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                            myiPropsForm.tbDescription.ForeColor = AddinGlobal.ForeColour
                            myiPropsForm.tbPartNumber.ForeColor = AddinGlobal.ForeColour

                            iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, myiPropsForm.tbDescription.Text, "", True)
                            iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, myiPropsForm.tbPartNumber.Text, "", True)

                        End If
                    End If
                End If
            Else
                If TypeOf oDoc Is PartDocument Then
                    Dim oPartDoc As PartDocument = oDoc
                    If CommandName = "SheetMetalStylesCmd" Then
                        Dim partcompdef As PartComponentDefinition = oPartDoc.ComponentDefinition
                        Dim sheetmetalcompdef As SheetMetalComponentDefinition = partcompdef
                        Dim oUnfoldMethod As String = sheetmetalcompdef.UnfoldMethod.Name
                        UpdateCustomiProperty(oDoc, "Sheet Metal Rule", oUnfoldMethod)
                    End If
                End If
            End If

        End Sub

        Private Sub m_ApplicationEvents_OnCloseDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                UpdateDisplayediProperties()
            Else
                If myiPropsForm IsNot Nothing Then
                    myiPropsForm.tbDescription.Text = String.Empty
                    myiPropsForm.tbPartNumber.Text = String.Empty
                    myiPropsForm.tbStockNumber.Text = String.Empty
                    myiPropsForm.tbEngineer.Text = String.Empty
                    myiPropsForm.tbDrawnBy.Text = String.Empty
                    myiPropsForm.tbRevNo.Text = String.Empty
                    myiPropsForm.tbComments.Text = String.Empty
                    myiPropsForm.tbNotes.Text = String.Empty
                    myiPropsForm.Label12.Text = String.Empty
                    myiPropsForm.FileLocation.Text = String.Empty
                    myiPropsForm.ModelFileLocation.Text = String.Empty
                    myiPropsForm.tbService.Text = String.Empty
                End If
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnNewEditObject(EditObject As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
                        UpdateDisplayediProperties()
                    ElseIf TypeOf (EditObject) Is Document Then
                        Dim selecteddoc As Document = AddinGlobal.InventorApp.ActiveEditObject
                        UpdateDisplayediProperties(selecteddoc)
                    Else
                        UpdateDisplayediProperties()
                    End If
                End If
            End If
        End Sub

        Public Shared Sub UpdateStatusBar(ByVal Message As String)
            AddinGlobal.InventorApp.StatusBarText = Message
        End Sub

        Private Sub m_UserInputEvents_OnActivateCommand(CommandName As String, Context As NameValueMap)
            If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
                    If CommandName = "VaultCheckinTop" Or CommandName = "VaultCheckin" Then
                        DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument

                        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                            WhatToDo = MsgBox("Updates are not Deferred, do you want to Defer them?", vbYesNo, "Deferred Checker")
                            If WhatToDo = vbYes Then
                                AddinGlobal.InventorApp.ActiveDocument.DrawingSettings.DeferUpdates = True
                                myiPropsForm.btDefer.BackColor = Drawing.Color.Red
                                myiPropsForm.btDefer.Text = "Drawing Updates Deferred"
                                UpdateStatusBar("Updates are now Deferred")
                                MsgBox("Updates are now Deferred, continue Checkin", vbOKOnly, "Deferred Checker")
                            End If
                        End If
                    End If
                Else
                    If CommandName = "VaultCheckinTop" Or CommandName = "VaultCheckin" Then
                        DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                        Dim PartNo As String = myiPropsForm.tbPartNumber.Text
                        Dim StockNo As String = myiPropsForm.tbStockNumber.Text
                        If Not PartNo = StockNo Then
                            stockNum = MsgBox("Your Stock Number and Part Number are different, is this OK?", vbYesNo, "Stock/Part Number Check")
                            If stockNum = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If CommandName = "GeomToDXFCommand" Then
                        DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument

                        Dim flatName As String = myiPropsForm.tbPartNumber.Text
                        Clipboard.SetText(flatName)
                    End If

                    'Dim oDoc As Document = AddinGlobal.InventorApp.ActiveDocument
                    'If TypeOf oDoc Is PartDocument Then
                    '    If CommandName = "AppFileSaveCmd" Or CommandName = "AppFileSaveAsCmd" Then
                    '        Dim Material As String = iProperties.GetorSetStandardiProperty(oDoc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties)
                    '        Dim Weight As Decimal = iProperties.GetorSetStandardiProperty(oDoc, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties)
                    '        Dim kgWeight As Decimal = Weight / 1000
                    '        Dim Weight2 As Decimal = Math.Round(kgWeight, 1)
                    '        If Material = "Generic" Then
                    '            whatnow = MsgBox("Material is set as " & Material & " are you sure you don't want it to be something more shiny?", vbYesNo, "Material Check")
                    '            If whatnow = vbNo Then
                    '                AllowFileToSave = False
                    '                AllowFileToSaveAs = False
                    '            End If
                    '        End If
                    '        If Weight2 > 10 Then
                    '            whatnow = MsgBox("The weight of this part is quite high, " & Weight2 & "kg. Are you sure you're happy with that?", vbYesNo, "Weight Check")
                    '            If whatnow = vbNo Then
                    '                AllowFileToSave = False
                    '                AllowFileToSaveAs = False
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnActivateView(ViewObject As Inventor.View, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                myiPropsForm.tbComments.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, "", "")
                If TypeOf (DocumentToPulliPropValuesFrom) Is DrawingDocument Then
                    If DocumentToPulliPropValuesFrom IsNot Nothing Then
                        Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                        Dim oSht As Sheet = oDWG.ActiveSheet
                        Dim oView As DrawingView = Nothing
                        Dim drawnDoc As Document = Nothing

                        For Each view As DrawingView In oSht.DrawingViews
                            oView = view
                            Exit For
                        Next
                        If oView IsNot Nothing Then
                            drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                            revno = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                            modrev = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")
                            If revno > modrev Then
                                myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

                                iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, revno, "", True)
                            End If
                        End If

                        SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                        UpdateFormTextBoxColours()
                    End If
                ElseIf TypeOf (DocumentToPulliPropValuesFrom) Is AssemblyDocument Then
                    Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                    If AssyDoc.SelectSet.Count = 1 Then
                        Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                        DocumentToPulliPropValuesFrom = compOcc.Definition.Document
                        If AddinGlobal.InventorApp.ActiveEditObject IsNot DocumentToPulliPropValuesFrom Then
                            DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveEditObject
                        ElseIf DocumentToPulliPropValuesFrom IsNot Nothing Then
                            SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                            UpdateFormTextBoxColours()
                        End If
                    Else
                        If AddinGlobal.InventorApp.ActiveEditObject IsNot DocumentToPulliPropValuesFrom Then
                            DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveEditObject
                        ElseIf DocumentToPulliPropValuesFrom IsNot Nothing Then
                            SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                            UpdateFormTextBoxColours()
                        End If
                    End If
                Else
                    If AddinGlobal.InventorApp.ActiveEditObject IsNot DocumentToPulliPropValuesFrom Then
                        DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveEditObject
                    End If
                    If DocumentToPulliPropValuesFrom IsNot Nothing Then
                        SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                        UpdateFormTextBoxColours()
                    End If
                End If
            End If
        End Sub

        Public Shared Sub ShowOccurrenceProperties(AssyDoc As AssemblyDocument)
            If AssyDoc.SelectSet.Count = 1 Then
                Dim selecteddoc As Document = Nothing
                Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                Dim def As ComponentDefinition
                def = compOcc.Definition
                If TypeOf def Is VirtualComponentDefinition Then
                    Dim virtualDef As VirtualComponentDefinition
                    virtualDef = compOcc.Definition
                    selecteddoc = virtualDef.Document

                    UpdateDisplayediProperties(selecteddoc)
                    'AssyDoc.SelectSet.Select(compOcc)
                    UpdateFormTextBoxColours()
                Else
                    selecteddoc = compOcc.Definition.Document

                    UpdateDisplayediProperties(selecteddoc)
                    If AssyDoc.SelectSet.Count = 0 Then
                        AssyDoc.SelectSet.Select(compOcc) ' _inChangeSelectSetHandler is required because of this
                    End If
                    UpdateFormTextBoxColours()
                End If

                'selecteddoc = compOcc.Definition.Document
                '    'Dim VirtualDef As VirtualComponentDefinition = TryCast(compOcc.Definition, VirtualComponentDefinition)
                '    'Dim selectedVirtdoc As Document = Nothing
                '    'selectedVirtdoc = VirtualDef.Document

                '    'UpdateDisplayediProperties(selectedVirtdoc)
                '    UpdateDisplayediProperties(selecteddoc)
                '    AssyDoc.SelectSet.Select(compOcc)
                '    UpdateFormTextBoxColours()
            End If
        End Sub

        Private Sub m_StyleEvents_OnActivateStyle(DocumentObject As _Document, Material As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject) Then
                    If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                        Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                        Material = iProperties.GetorSetStandardiProperty(
                               DocumentToPulliPropValuesFrom,
                               PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                        'If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
                        AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                        Dim kgMass As Decimal = myMass / 1000
                        Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                        myiPropsForm.tbMass.Text = myMass2 & " kg"
                        Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                        Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                        myiPropsForm.tbDensity.Text = myDensity2 & " g/cm^3"

                        myiPropsForm.Label12.Text = iProperties.GetorSetStandardiProperty(
                           DocumentToPulliPropValuesFrom,
                           PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                        'End If
                    End If
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private _inChangeSelectSetHandler As Boolean = False

        ''' <summary>
        ''' This method is what helps us capture properties from selected items.
        ''' It works fine in Parts and Drawings, but not currently Assemblies.
        ''' </summary>
        ''' <param name="BeforeOrAfter"></param>
        ''' <param name="Context"></param>
        ''' <param name="HandlingCode"></param>
        Private Sub m_DocumentEvents_OnChangeSelectSet(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            HandlingCode = HandlingCodeEnum.kEventNotHandled
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If _inChangeSelectSetHandler Then
                    ' We have probably caused this by changing the select set within the handler.
                    ' Avoid recursion by returning early.
                    _inChangeSelectSetHandler = False
                    Return
                End If

                _inChangeSelectSetHandler = True
                Try
                    If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then

                        If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                            Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                            If AssyDoc.SelectSet.Count = 1 Then
                                UpdateFormTextBoxColours()

                                If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                                    ShowOccurrenceProperties(AssyDoc)
                                ElseIf TypeOf AssyDoc.SelectSet(1) Is HoleFeatureProxy Then
                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                    myiPropsForm.tbRevNo.ReadOnly = True
                                    myiPropsForm.tbComments.ReadOnly = True
                                    myiPropsForm.tbNotes.ReadOnly = True
                                    Dim FeatOcc As PartFeature = AssyDoc.SelectSet(1)
                                    Dim holeOcc As HoleFeature = AssyDoc.SelectSet(1)

                                    myiPropsForm.tbEngineer.Text = "Reading Part Hole Properties"
                                    myiPropsForm.tbPartNumber.Text = FeatOcc.Name
                                    myiPropsForm.tbStockNumber.Text = FeatOcc.Name
                                    myiPropsForm.tbDescription.Text = FeatOcc.ThreadDesignation
                                ElseIf TypeOf AssyDoc.SelectSet(1) Is HoleFeature Then
                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                    myiPropsForm.tbRevNo.ReadOnly = True
                                    myiPropsForm.tbComments.ReadOnly = True
                                    myiPropsForm.tbNotes.ReadOnly = True
                                    Dim holeOcc As PartFeature = AssyDoc.SelectSet(1)

                                    myiPropsForm.tbEngineer.Text = "Reading Assembly Hole Properties"
                                    myiPropsForm.tbPartNumber.Text = holeOcc.Name
                                    myiPropsForm.tbStockNumber.Text = holeOcc.Name
                                    myiPropsForm.tbDescription.Text = "Hole at assy level, cannot show details :("


                                ElseIf TypeOf AssyDoc.SelectSet(1) Is ExtrudeFeature Or TypeOf AssyDoc.SelectSet(1) Is CutFeature Or TypeOf AssyDoc.SelectSet(1) Is HoleFeature Or TypeOf AssyDoc.SelectSet(1) Is BossFeature Or TypeOf AssyDoc.SelectSet(1) Is SweepFeature Or TypeOf AssyDoc.SelectSet(1) Is LoftFeature Or TypeOf AssyDoc.SelectSet(1) Is PartFeature Then
                                    AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AssemblyShowAssemblyFeatureDimsCtxCmd").Execute()

                                Else
                                    myiPropsForm.tbPartNumber.ReadOnly = False
                                    myiPropsForm.tbDescription.ReadOnly = False
                                    myiPropsForm.tbStockNumber.ReadOnly = False
                                    myiPropsForm.tbEngineer.ReadOnly = False
                                    myiPropsForm.tbRevNo.ReadOnly = False
                                    myiPropsForm.tbComments.ReadOnly = False
                                    myiPropsForm.tbNotes.ReadOnly = False
                                    UpdateDisplayediProperties(AssyDoc)
                                End If
                            ElseIf AssyDoc.SelectSet.Count = 0 Then
                                If AddinGlobal.InventorApp.ActiveEditDocument Is Nothing Then
                                    AssyDoc = AddinGlobal.InventorApp.ActiveDocument
                                Else
                                    AssyDoc = AddinGlobal.InventorApp.ActiveEditDocument
                                End If
                                ShowOccurrenceProperties(AssyDoc)
                                If CheckReadOnly(AddinGlobal.InventorApp.ActiveDocument) Then
                                    'myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                                    'myiPropsForm.Label10.Text = "Checked In"
                                    myiPropsForm.PictureBox1.Show()
                                    myiPropsForm.PictureBox2.Hide()
                                    myiPropsForm.btCheckIn.Hide()
                                    myiPropsForm.btCheckOut.Show()

                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                    myiPropsForm.tbRevNo.ReadOnly = True
                                    myiPropsForm.tbComments.ReadOnly = True
                                    myiPropsForm.tbNotes.ReadOnly = True
                                Else
                                    'myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                                    'myiPropsForm.Label10.Text = "Checked Out"
                                    myiPropsForm.PictureBox1.Hide()
                                    myiPropsForm.PictureBox2.Show()
                                    myiPropsForm.btCheckIn.Show()
                                    myiPropsForm.btCheckOut.Hide()

                                    myiPropsForm.tbPartNumber.ReadOnly = False
                                    myiPropsForm.tbDescription.ReadOnly = False
                                    myiPropsForm.tbStockNumber.ReadOnly = False
                                    myiPropsForm.tbEngineer.ReadOnly = False
                                    myiPropsForm.tbRevNo.ReadOnly = False
                                    myiPropsForm.tbComments.ReadOnly = False
                                    myiPropsForm.tbNotes.ReadOnly = False
                                End If
                                UpdateDisplayediProperties(AssyDoc)
                                UpdateFormTextBoxColours()
                            End If
                        ElseIf (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject) Then
                            UpdateFormTextBoxColours()
                            Dim PartDoc As PartDocument = AddinGlobal.InventorApp.ActiveEditDocument
                            If PartDoc.SelectSet.Count = 1 Then
                                If TypeOf PartDoc.SelectSet(1) Is PartFeature Then
                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                    myiPropsForm.tbRevNo.ReadOnly = True
                                    myiPropsForm.tbComments.ReadOnly = True
                                    myiPropsForm.tbNotes.ReadOnly = True
                                    Dim FeatOcc As PartFeature = PartDoc.SelectSet(1)

                                    myiPropsForm.tbEngineer.Text = "Reading Feature Properties"
                                    myiPropsForm.tbPartNumber.Text = FeatOcc.Name
                                    myiPropsForm.tbStockNumber.Text = FeatOcc.Name
                                    If FeatOcc.ExtendedName IsNot Nothing Then
                                        myiPropsForm.tbDescription.Text = FeatOcc.ExtendedName
                                    End If
                                    AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("PartShowDimensionsCtxCmd").Execute()
                                Else
                                    myiPropsForm.tbPartNumber.ReadOnly = False
                                    myiPropsForm.tbDescription.ReadOnly = False
                                    myiPropsForm.tbStockNumber.ReadOnly = False
                                    myiPropsForm.tbEngineer.ReadOnly = False
                                    myiPropsForm.tbRevNo.ReadOnly = False
                                    myiPropsForm.tbComments.ReadOnly = False
                                    myiPropsForm.tbNotes.ReadOnly = False
                                    UpdateDisplayediProperties(PartDoc)
                                End If
                            Else
                                myiPropsForm.tbPartNumber.ReadOnly = False
                                myiPropsForm.tbDescription.ReadOnly = False
                                myiPropsForm.tbStockNumber.ReadOnly = False
                                myiPropsForm.tbEngineer.ReadOnly = False
                                myiPropsForm.tbRevNo.ReadOnly = False
                                myiPropsForm.tbComments.ReadOnly = False
                                myiPropsForm.tbNotes.ReadOnly = False
                                UpdateDisplayediProperties(PartDoc)
                            End If
                        Else
                            'If (AddinGlobal.InventorApp.ActiveEditDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
                            '    'Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument.ReferencedDocument

                            '    Dim DrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

                            '    Dim AssyDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveEditDocument


                            '    If AssyDoc.SelectSet.Count = 1 Then
                            '        If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                            '            ShowOccurrenceProperties(AssyDoc)
                            '        Else
                            '            UpdateFormTextBoxColours()
                            '            UpdateDisplayediProperties(DrawDoc)
                            '        End If
                            '    Else
                            '        UpdateFormTextBoxColours()
                            '        UpdateDisplayediProperties(DrawDoc)
                            '    End If


                            Dim DrawDoc As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument

                            UpdateFormTextBoxColours()
                            UpdateDisplayediProperties(DrawDoc)

                        End If
                        'End If
                    End If
                Finally
                    _inChangeSelectSetHandler = False
                End Try
            End If
        End Sub

        Private Shared Sub UpdateFormTextBoxColours()
            myiPropsForm.tbDescription.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbPartNumber.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbStockNumber.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbEngineer.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbDrawnBy.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbRevNo.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbComments.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbNotes.ForeColor = AddinGlobal.ForeColour
            myiPropsForm.tbService.ForeColor = AddinGlobal.ForeColour
        End Sub

        Private Sub m_ApplicationEvents_OnQuit(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kBefore Then
                InventorAppQuitting = True
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnSaveDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)

            If BeforeOrAfter = EventTimingEnum.kBefore Then
                'put stuff in here that you want Inventor to do prior to saving the file.
                'things such as checking the material is set from something other than generic
                ' and then the messagebox if whatever you're checking returns as true:
                Dim WhateverWeChecked As Boolean = False

                If WhateverWeChecked Then
                    'this should cancel the save event allowing the user to set a material
                    Exit Sub
                End If

            End If

            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
                myiPropsForm.tbDrawnBy.ForeColor = AddinGlobal.ForeColour
                myiPropsForm.GetNewFilePaths()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled

            'If Not AllowFileToSave Or Not AllowFileToSaveAs Then
            '    HandlingCode = HandlingCodeEnum.kEventCanceled
            'End If

        End Sub

        Private Sub m_ApplicationEvents_OnActivateDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                m_DocEvents = DocumentObject.DocumentEvents
                AddHandler m_DocEvents.OnChangeSelectSet, AddressOf Me.m_DocumentEvents_OnChangeSelectSet
                Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                If DocumentObject Is AddinGlobal.InventorApp.ActiveDocument Then
                    SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                    UpdateDisplayediProperties(DocumentToPulliPropValuesFrom)
                    UpdateFormTextBoxColours()
                    myiPropsForm.GetNewFilePaths()
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnOpenDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                'this change prevents this firing for EVERY opening file.
                If DocumentObject Is AddinGlobal.InventorApp.ActiveDocument Then
                    SetFormDisplayOption(DocumentToPulliPropValuesFrom)
                    UpdateDisplayediProperties(DocumentToPulliPropValuesFrom)
                    UpdateFormTextBoxColours()
                    myiPropsForm.GetNewFilePaths()

                End If
            End If

            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnNewDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                    UpdateDisplayediProperties()
                Else
                    myiPropsForm.tbDescription.Text = String.Empty
                    myiPropsForm.tbPartNumber.Text = String.Empty
                    myiPropsForm.tbStockNumber.Text = String.Empty
                    myiPropsForm.tbEngineer.Text = String.Empty
                    myiPropsForm.tbDrawnBy.Text = String.Empty
                    myiPropsForm.tbRevNo.Text = String.Empty
                    myiPropsForm.tbComments.Text = String.Empty
                    myiPropsForm.tbNotes.Text = String.Empty
                    myiPropsForm.Label12.Text = String.Empty
                    myiPropsForm.FileLocation.Text = String.Empty
                    myiPropsForm.ModelFileLocation.Text = String.Empty
                    myiPropsForm.tbService.Text = String.Empty
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        ''' <summary>
        ''' Need to add more updates here as we add textboxes and therefore properties to this list.
        '''
        ''' </summary>
        ''' <param name="DocumentToPulliPropValuesFrom"></param>
        Public Shared Sub UpdateDisplayediProperties(Optional DocumentToPulliPropValuesFrom As Document = Nothing)

            If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing And DocumentToPulliPropValuesFrom Is Nothing Then
                DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
            End If
            'If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then

            If DocumentToPulliPropValuesFrom IsNot Nothing Then
                myiPropsForm.FileLocation.ForeColor = AddinGlobal.ForeColour
                myiPropsForm.FileLocation.Text = DocumentToPulliPropValuesFrom.FullFileName
                myiPropsForm.tbComments.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, "", "")
            Else ' use the active edit object in cases where we're editing-in-place
                If AddinGlobal.InventorApp.ActiveEditObject IsNot Nothing Then
                    myiPropsForm.FileLocation.ForeColor = AddinGlobal.ForeColour
                    myiPropsForm.FileLocation.Text = AddinGlobal.InventorApp.ActiveEditDocument.FullFileName
                    myiPropsForm.tbComments.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, "", "")
                Else
                    myiPropsForm.FileLocation.ForeColor = AddinGlobal.ForeColour
                    myiPropsForm.FileLocation.Text = AddinGlobal.InventorApp.ActiveDocument.FullFileName
                    myiPropsForm.tbComments.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, "", "")
                End If
            End If



            myiPropsForm.tbComments.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kCommentsSummaryInformation, "", "")


            If TypeOf (DocumentToPulliPropValuesFrom) Is DrawingDocument Then
                myiPropsForm.btDefer.Show()
                myiPropsForm.Label7.Show()
                myiPropsForm.tbDrawnBy.Show()
                myiPropsForm.btShtMaterial.Show()
                myiPropsForm.btShtScale.Show()
                myiPropsForm.ModelFileLocation.Show()
                myiPropsForm.Label5.Hide()
                myiPropsForm.tbMass.Hide()
                myiPropsForm.Label9.Hide()
                myiPropsForm.tbDensity.Hide()
                myiPropsForm.Label3.Hide()
                myiPropsForm.tbStockNumber.Hide()
                myiPropsForm.btCopyPN.Hide()
                myiPropsForm.btViewNames.Show()
                myiPropsForm.lbDesigner.Hide()
                myiPropsForm.tbService.Hide()
                myiPropsForm.lbservice.Hide()
                myiPropsForm.btExpDXF.Hide()
                myiPropsForm.btFrame.Show()

                myiPropsForm.tbDrawnBy.Text = iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForSummaryInformationEnum.kAuthorSummaryInformation, "", "")

                Dim oDWG As DrawingDocument = AddinGlobal.InventorApp.ActiveDocument
                Dim oSht As Sheet = oDWG.ActiveSheet
                Dim oView As DrawingView = Nothing
                Dim drawnDoc As Document = Nothing
                Dim MaterialString As String = String.Empty
                Dim DrawDesc As String = String.Empty
                Dim ModelDesc As String = String.Empty

                If CheckReadOnly(DocumentToPulliPropValuesFrom) Then

                    myiPropsForm.tbDrawnBy.ReadOnly = True

                    'drawnDoc = DocumentToPulliPropValuesFrom
                    'DrawDesc = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    'ModelDesc = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    If iProperties.GetorSetStandardiProperty(
                                          DocumentToPulliPropValuesFrom,
                                          PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                        myiPropsForm.btDefer.BackColor = Drawing.Color.Red
                        myiPropsForm.btDefer.Text = "Drawing Updates Deferred"
                    ElseIf iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                        myiPropsForm.btDefer.BackColor = Drawing.Color.Green
                        myiPropsForm.btDefer.Text = "Drawing Updates Not Deferred"
                    End If

                Else

                    myiPropsForm.tbDrawnBy.ReadOnly = False
                    If DocumentToPulliPropValuesFrom.FullDocumentName IsNot Nothing Then

                        If iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                            myiPropsForm.btDefer.BackColor = Drawing.Color.Red
                            myiPropsForm.btDefer.Text = "Drawing Updates Deferred"
                        ElseIf iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                               PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                            myiPropsForm.btDefer.BackColor = Drawing.Color.Green
                            myiPropsForm.btDefer.Text = "Drawing Updates Not Deferred"

                            For Each view As DrawingView In oSht.DrawingViews
                                oView = view
                                Exit For
                            Next

                            If oView IsNot Nothing Then
                                drawnDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument

                                If TypeOf drawnDoc Is AssemblyDocument Then
                                    myiPropsForm.btITEM.Show()
                                    myiPropsForm.btReNum.Show()
                                    myiPropsForm.Label11.Show()
                                    myiPropsForm.Label12.Show()


                                    MaterialString = "See Above"
                                Else
                                    myiPropsForm.btITEM.Hide()
                                    myiPropsForm.btReNum.Hide()
                                    myiPropsForm.Label11.Show()
                                    myiPropsForm.Label12.Show()

                                    MaterialString = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                                End If

                                MainPath = System.IO.Path.GetDirectoryName(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullFileName)
                                ModelPath = MainPath & "\" & System.IO.Path.GetFileNameWithoutExtension(oView.ReferencedDocumentDescriptor.ReferencedDocument.FullDocumentName)

                                myiPropsForm.ModelFileLocation.ForeColor = AddinGlobal.ForeColour
                                myiPropsForm.ModelFileLocation.Text = ModelPath

                                myiPropsForm.Label12.Text = MaterialString
                                'DrawDesc = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                                'ModelDesc = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")

                                'myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                                'myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")

                            End If

                        End If
                    Else
                        'myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        'myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                    End If
                End If

                myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                myiPropsForm.tbRevNo.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

            Else
                myiPropsForm.Label5.Show()
                myiPropsForm.tbMass.Show()
                myiPropsForm.Label9.Show()
                myiPropsForm.tbDensity.Show()
                myiPropsForm.Label3.Show()
                myiPropsForm.tbStockNumber.Show()
                myiPropsForm.Label7.Hide()
                myiPropsForm.tbDrawnBy.Hide()
                myiPropsForm.btDefer.Hide()
                myiPropsForm.btShtMaterial.Hide()
                myiPropsForm.btShtScale.Hide()
                myiPropsForm.ModelFileLocation.Hide()
                myiPropsForm.btCopyPN.Show()
                myiPropsForm.btViewNames.Hide()
                myiPropsForm.lbDesigner.Show()

                If DocumentToPulliPropValuesFrom.FullFileName IsNot Nothing Then
                    If DocumentToPulliPropValuesFrom.FullDocumentName IsNot Nothing Then
                        If DocumentToPulliPropValuesFrom.FullDocumentName.Contains("pisweep") Then
                            myiPropsForm.lbservice.Show()
                            myiPropsForm.tbService.Show()
                            myiPropsForm.btDiaEng.Hide()
                            myiPropsForm.btDegEng.Hide()
                            myiPropsForm.Label4.Hide()
                        ElseIf iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "").Contains("WALL TUBE") Then
                            myiPropsForm.lbservice.Show()
                            myiPropsForm.tbService.Show()
                            myiPropsForm.btDiaEng.Hide()
                            myiPropsForm.btDegEng.Hide()
                            myiPropsForm.Label4.Hide()
                        Else
                            myiPropsForm.lbservice.Hide()
                            myiPropsForm.tbService.Hide()
                            myiPropsForm.btDiaEng.Show()
                            myiPropsForm.btDegEng.Show()
                            myiPropsForm.Label4.Show()
                        End If
                    End If
                End If

                If TypeOf (DocumentToPulliPropValuesFrom) Is AssemblyDocument Then
                    myiPropsForm.btITEM.Show()
                    myiPropsForm.btReNum.Show()
                    myiPropsForm.Label11.Hide()
                    myiPropsForm.Label12.Hide()
                    myiPropsForm.btExpDXF.Hide()
                    myiPropsForm.btFrame.Show()
                ElseIf TypeOf (DocumentToPulliPropValuesFrom) Is PartDocument Then
                    myiPropsForm.btITEM.Hide()
                    myiPropsForm.btReNum.Hide()
                    myiPropsForm.Label11.Show()
                    myiPropsForm.Label12.Show()
                    myiPropsForm.btExpDXF.Show()
                    myiPropsForm.btFrame.Hide()
                End If
                'get the document sub-type

                myiPropsForm.tbStockNumber.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")

                Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                Dim kgMass As Decimal = myMass / 1000
                Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                myiPropsForm.tbMass.Text = myMass2 & " kg"
                Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                myiPropsForm.tbDensity.Text = myDensity2 & " g/cm^3"
                myiPropsForm.Label12.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                myiPropsForm.lbDesigner.Text = "By: " & iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kAuthorSummaryInformation, "", "")

                myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")

                myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")

                myiPropsForm.tbService.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kProjectDesignTrackingProperties, "", "")

                myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")

                myiPropsForm.tbRevNo.Text = iProperties.GetorSetStandardiProperty(DocumentToPulliPropValuesFrom, PropertiesForSummaryInformationEnum.kRevisionSummaryInformation, "", "")

            End If

            Dim todaysdate As String = String.Format("{0:DD/MM/yyyy}", DateTime.Now)

            If PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties = True Then
                myiPropsForm.DateTimePicker1.Value = iProperties.GetorSetStandardiProperty(
                    DocumentToPulliPropValuesFrom,
                    PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties, "", "")
                'Else
                '    DocumentToPulliPropValuesFrom.PropertySets.Item("Design Tracking Properties").Item("Creation Date").Value = todaysdate
            End If

            myiPropsForm.tbNotes.Text = iProperties.GetorSetStandardiProperty(
                    DocumentToPulliPropValuesFrom,
                    PropertiesForDesignTrackingPropertiesEnum.kCatalogWebLinkDesignTrackingProperties, "", "")


            'simplified to this:
            UpdateFormTextBoxColours()
            SetFormDisplayOption(DocumentToPulliPropValuesFrom)

        End Sub

        Private Sub UpdateCustomiProperty(ByRef Doc As Inventor.Document, ByRef PropertyName As String, ByRef PropertyValue As String)
            ' Get the custom property set.
            Dim customPropSet As Inventor.PropertySet
            customPropSet = Doc.PropertySets.Item("Inventor User Defined Properties")

            ' Get the existing property, if it exists.
            Dim prop As Inventor.Property = Nothing
            Dim propExists As Boolean = True
            Try
                prop = customPropSet.Item(PropertyName)
            Catch ex As Exception
                propExists = False
            End Try

            ' Check to see if the property was successfully obtained.
            If Not propExists Then
                ' Failed to get the existing property so create a new one.
                prop = customPropSet.Add(PropertyValue, PropertyName)
            Else
                ' Change the value of the existing property.
                prop.Value = PropertyValue
            End If
        End Sub

        ''' <summary>
        ''' there appeared to be three locations so far that had the exact same signature so have refactored to make this method.
        ''' </summary>
        ''' <param name="DocumentToPulliPropValuesFrom"></param>
        Private Shared Sub SetFormDisplayOption(DocumentToPulliPropValuesFrom As Document)
            If AddinGlobal.InventorApp.ActiveDocument IsNot Nothing Then
                If CheckReadOnly(DocumentToPulliPropValuesFrom) = True Then
                    'myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                    'myiPropsForm.Label10.Text = "Checked In"
                    myiPropsForm.PictureBox1.Show()
                    myiPropsForm.PictureBox2.Hide()
                    myiPropsForm.btCheckIn.Hide()
                    myiPropsForm.btCheckOut.Show()

                    myiPropsForm.tbPartNumber.ReadOnly = True
                    myiPropsForm.tbDescription.ReadOnly = True
                    myiPropsForm.tbStockNumber.ReadOnly = True
                    myiPropsForm.tbEngineer.ReadOnly = True
                    myiPropsForm.tbRevNo.ReadOnly = True
                    myiPropsForm.tbComments.ReadOnly = True
                    myiPropsForm.tbNotes.ReadOnly = True
                    myiPropsForm.tbDrawnBy.ReadOnly = True
                    myiPropsForm.tbService.ReadOnly = True
                Else
                    'myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                    'myiPropsForm.Label10.Text = "Checked Out"
                    myiPropsForm.PictureBox1.Hide()
                    myiPropsForm.PictureBox2.Show()
                    myiPropsForm.btCheckIn.Show()
                    myiPropsForm.btCheckOut.Hide()

                    myiPropsForm.tbPartNumber.ReadOnly = False
                    myiPropsForm.tbDescription.ReadOnly = False
                    myiPropsForm.tbStockNumber.ReadOnly = False
                    myiPropsForm.tbEngineer.ReadOnly = False
                    myiPropsForm.tbRevNo.ReadOnly = False
                    myiPropsForm.tbComments.ReadOnly = False
                    myiPropsForm.tbNotes.ReadOnly = False
                    myiPropsForm.tbDrawnBy.ReadOnly = False
                    myiPropsForm.tbService.ReadOnly = False
                End If
            End If

        End Sub

        ''' <summary>
        ''' Original copied verbatim from here:
        ''' http://adndevblog.typepad.com/manufacturing/2012/05/checking-whether-a-inventor-document-is-read-only-or-not.html
        ''' Modified as suggested by this page:
        ''' https://msdn.microsoft.com/en-us/library/system.io.fileattributes(v=vs.110).aspx?f=255&MSPPError=-2147217396&cs-save-lang=1&cs-lang=vb#code-snippet-2
        ''' </summary>
        ''' <param name="doc"></param>
        ''' <returns></returns>
        Public Shared Function CheckReadOnly(ByVal doc As Document) As Boolean
            Try
                ' Handle the case with the active document never saved
                If System.IO.File.Exists(doc.FullFileName) = False Then
                    UpdateStatusBar("Save file before executing this method. Exiting ...")
                    Return False
                End If


                Dim atts As FileAttributes = IO.File.GetAttributes(doc.FullFileName)

                If ((atts And FileAttributes.ReadOnly) = System.IO.FileAttributes.ReadOnly) Then
                    Return True
                Else
                    'The file is Read/Write
                    Return False
                End If
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
        End Function

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            Try
                ' TODO:  Add ApplicationAddInServer.Deactivate implementation
                For Each item As InventorButton In AddinGlobal.ButtonList
                    Marshal.FinalReleaseComObject(item.ButtonDef)
                Next

                For Each item As DockableWindow In AddinGlobal.DockableList
                    Marshal.FinalReleaseComObject(item)
                Next

                ' Release objects.

                m_UserInputEvents = Nothing
                m_AppEvents = Nothing
                m_uiEvents = Nothing
                m_StyleEvents = Nothing

                myiPropsForm.CurrentPath = Nothing
                myiPropsForm.NewPath = Nothing
                myiPropsForm.RefNewPath = Nothing
                myiPropsForm.RefDoc = Nothing
                thisAssembly = Nothing
                myiPropsForm = Nothing

                If AddinGlobal.RibbonPanel IsNot Nothing Then
                    Marshal.FinalReleaseComObject(AddinGlobal.RibbonPanel)
                End If

                If Not InventorAppQuitting Then
                    If AddinGlobal.InventorApp IsNot Nothing Then
                        Marshal.FinalReleaseComObject(AddinGlobal.InventorApp)
                    End If
                End If

                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Private m_thisWindow As DockableWindow

        Public Property Window() As DockableWindow
            Get
                Return m_thisWindow
            End Get
            Set(ByVal value As DockableWindow)
                m_thisWindow = value
            End Set
        End Property

        'Public Property Window As DockableWindow
        '    Get
        '        Return
        '    End Get
        '    Set(value As DockableWindow)

        '    End Set
        'End Property

        ' Note:this method is now obsolete, you should use the
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

        End Sub

#End Region

#Region "User interface definition"

        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface(button1 As InventorButton)
            ' This is where you'll add code to add buttons to the ribbon.

            '** Sample to illustrate creating a button on a new panel of the Tools tab of the Part ribbon.

            '' Get the part ribbon.
            'Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")

            '' Get the "Tools" tab.
            'Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")

            '' Create a new panel.
            'Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Sample", "MysSample", AddInClientID)

            '' Add a button.
            'customPanel.CommandControls.AddButton(m_sampleButton)
            Try

                Dim uiMan As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
                If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    'kClassicInterface support can be added if necessary.
                    Dim ribbon As Inventor.Ribbon = uiMan.Ribbons("Part")
                    Dim tab As RibbonTab
                    Try
                        tab = ribbon.RibbonTabs("id_TabSheetMetal") 'Change it if necessary.
                    Catch
                        tab = ribbon.RibbonTabs.Add("id_TabSheetMetal", "id_Tabid_TabSheetMetal", Guid.NewGuid().ToString())
                    End Try
                    AddinGlobal.RibbonPanelId = "{51f8ccf4-5fc6-4592-b68d-e19c993f5faa}"
                    AddinGlobal.RibbonPanel = tab.RibbonPanels.Add("InventorNetAddin", "MyVBInventorAddin.RibbonPanel_" & Guid.NewGuid().ToString(), AddinGlobal.RibbonPanelId, String.Empty, True)

                    Dim cmdCtrls As CommandControls = AddinGlobal.RibbonPanel.CommandControls
                    cmdCtrls.AddButton(button1.ButtonDef, button1.DisplayBigIcon, button1.DisplayText, "", button1.InsertBeforeTarget)
                End If
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

        'no need for this since we can just restart Inventor and have it reload the addin.
        'Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
        '    ' The ribbon was reset, so add back the add-ins user-interface.
        '    AddToUserInterface()
        'End Sub

        ' Sample handler for the button.
        'Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
        '    MsgBox("Button was clicked.")
        'End Sub

#End Region

    End Class

End Namespace

Public Module Globals

    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."

    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(iPropertiesController.iPropertiesAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function

#End Region

#Region "hWnd Wrapper Class"

    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class

#End Region

#Region "Image Converter"

    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter

        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
        Private Shared Function OleCreatePictureIndirect(
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object,
            ByRef iid As Guid,
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC

            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1

            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)>
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub

            End Class

            <StructLayout(LayoutKind.Sequential)>
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub

            End Class

        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function

    End Class

#End Region

End Module