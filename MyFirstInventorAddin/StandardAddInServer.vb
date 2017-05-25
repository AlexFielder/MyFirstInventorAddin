
Imports Inventor
Imports System.Runtime.InteropServices
Imports log4net
Imports System.Drawing
Imports System.Reflection
Imports iPropertiesController.iPropertiesController
Imports System.IO

Namespace iPropertiesController
    <ProgIdAttribute("iPropertiesController.StandardAddInServer"),
    GuidAttribute("e691af34-cb32-4296-8cca-5d1027a27c72")>
    Public Class StandardAddInServer
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


        Private thisAssembly As Assembly = Assembly.GetExecutingAssembly()
        Private thisAssemblyPath As String = String.Empty
        Public Shared attribute As GuidAttribute = Nothing
        Public myiPropsForm As iPropertiesForm = Nothing
        Public Property InventorAppQuitting As Boolean = False

        Private logHelper As Log4NetFileHelper.Log4NetFileHelper = New Log4NetFileHelper.Log4NetFileHelper()
        Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(StandardAddInServer))

        'Private WithEvents m_sampleButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) Implements ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            AddinGlobal.InventorApp = addInSiteObject.Application
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
                'AddHandler m_UserInputEvents.OnTerminateCommand, AddressOf Me.m_UserInputEvents_OnTerminateCommand
                'you can add extra handlers like this - if you uncomment the next line Visual Studio will prompt you to create the method:
                'AddHandler m_AssemblyEvents.OnNewOccurrence, AddressOf Me.m_AssemblyEvents_NewOcccurrence
                If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                    m_DocEvents = AddinGlobal.InventorApp.ActiveDocument.DocumentEvents
                    AddHandler m_DocEvents.OnChangeSelectSet, AddressOf Me.m_DocumentEvents_OnChangeSelectSet
                End If

                AddHandler m_StyleEvents.OnActivateStyle, AddressOf Me.m_StyleEvents_OnActivateStyle

                AddHandler m_AppEvents.OnNewEditObject, AddressOf Me.m_ApplicationEvents_OnNewEditObject

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
                    myiPropsForm = New iPropertiesForm(AddinGlobal.InventorApp, attribute.Value, localWindow)
                    Window = localWindow

                End If
                log.Info("Loaded My First Inventor Add-in")
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
        End Sub

        Private Sub m_ApplicationEvents_OnNewEditObject(EditObject As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
                        'Do nothing
                    Else
                        Dim selecteddoc As Document = AddinGlobal.InventorApp.ActiveEditObject
                        UpdateDisplayediProperties(selecteddoc)
                    End If
                End If
            End If
        End Sub

        'Private Sub m_UserInputEvents_OnTerminateCommand(CommandName As String, Context As NameValueMap)
        '    If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
        '        CommandName = "VaultCheckinTop"
        '    End If
        'End Sub

        Public Shared Sub UpdateStatusBar(ByVal Message As String)
            AddinGlobal.InventorApp.StatusBarText = Message
        End Sub

        Private Sub m_UserInputEvents_OnActivateCommand(CommandName As String, Context As NameValueMap)
            If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject) Then
                    If CommandName = "VaultCheckinTop" Then
                        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                            WhatToDo = MsgBox("Updates are not Deferred, do you want to Defer them?", vbYesNo, "Deferred Checker")
                            If WhatToDo = vbYes Then
                                AddinGlobal.InventorApp.ActiveDocument.DrawingSettings.DeferUpdates = True
                                myiPropsForm.Label8.ForeColor = Drawing.Color.Red
                                myiPropsForm.Label8.Text = "Drawing Updates Deferred"
                                UpdateStatusBar("Updates are now Deferred")
                                MsgBox("Updates are now Deferred, continue Checkin", vbOKOnly, "Deferred Checker")
                            Else
                                'Do Nothing
                            End If
                        End If
                    End If
                End If
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnActivateView(ViewObject As View, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                If BeforeOrAfter = EventTimingEnum.kAfter Then
                    Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument

                    If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
                        If CheckReadOnly(DocumentToPulliPropValuesFrom) Then
                            myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                            myiPropsForm.Label10.Text = "Checked In"
                            myiPropsForm.PictureBox1.Show()
                            myiPropsForm.PictureBox2.Hide()

                            myiPropsForm.tbPartNumber.ReadOnly = True
                            myiPropsForm.tbDescription.ReadOnly = True
                            myiPropsForm.tbStockNumber.ReadOnly = True
                            myiPropsForm.tbEngineer.ReadOnly = True
                        Else
                            myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                            myiPropsForm.Label10.Text = "Checked Out"
                            myiPropsForm.PictureBox1.Hide()
                            myiPropsForm.PictureBox2.Show()

                            myiPropsForm.tbPartNumber.ReadOnly = False
                            myiPropsForm.tbDescription.ReadOnly = False
                            myiPropsForm.tbStockNumber.ReadOnly = False
                            myiPropsForm.tbEngineer.ReadOnly = False
                        End If
                    End If
                End If
            End If
        End Sub

        Private Sub m_StyleEvents_OnActivateStyle(DocumentObject As _Document, Material As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject) Then
                    If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                        Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                        Material = iProperties.GetorSetStandardiProperty(
                               DocumentToPulliPropValuesFrom,
                               PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")
                        If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
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
                        End If
                    End If
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_DocumentEvents_OnChangeSelectSet(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                    If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                        Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                        If AssyDoc.SelectSet.Count = 1 Then
                            myiPropsForm.tbDescription.ForeColor = Drawing.Color.Black
                            myiPropsForm.tbPartNumber.ForeColor = Drawing.Color.Black
                            myiPropsForm.tbStockNumber.ForeColor = Drawing.Color.Black
                            myiPropsForm.tbEngineer.ForeColor = Drawing.Color.Black
                            myiPropsForm.tbDrawnBy.ForeColor = Drawing.Color.Black
                            If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                                myiPropsForm.tbPartNumber.ReadOnly = True
                                myiPropsForm.tbDescription.ReadOnly = True
                                myiPropsForm.tbStockNumber.ReadOnly = True
                                myiPropsForm.tbEngineer.ReadOnly = True
                                Dim selecteddoc As Document = Nothing
                                Dim compOcc As ComponentOccurrence = AssyDoc.SelectSet(1)
                                selecteddoc = compOcc.Definition.Document
                                UpdateDisplayediProperties(selecteddoc)
                                AssyDoc.SelectSet.Select(compOcc)

                            ElseIf TypeOf AssyDoc.SelectSet(1) Is PartFeature Then
                                myiPropsForm.tbPartNumber.ReadOnly = True
                                myiPropsForm.tbDescription.ReadOnly = True
                                myiPropsForm.tbStockNumber.ReadOnly = True
                                myiPropsForm.tbEngineer.ReadOnly = True
                                Dim FeatOcc As PartFeature = AssyDoc.SelectSet(1)
                                Dim spotsize = Nothing
                                Dim spotdepth = Nothing
                                Dim cbore = Nothing
                                Dim cboredepth = Nothing
                                Dim chamdia = Nothing
                                Dim holedia = Nothing

                                'If holeOcc.Tapped Then
                                '    HoleTap = holeOcc.TapInfo.ThreadDesignation
                                '    'HoleDia = holeOcc.TapInfo.ThreadType
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    myiPropsForm.tbDescription.Text = HoleTap & " " & holedia & " " & spotsize & " " &
                                '     cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbStockNumber.Text = "Tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'Else
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    holedia = "Ø" & holeOcc.HoleDiameter.Value * 10
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    myiPropsForm.tbDescription.Text = holedia & " " & spotsize & " " &
                                '       cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbStockNumber.Text = "Not a tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'End If
                                myiPropsForm.tbEngineer.Text = "Reading Feature Properties"
                                myiPropsForm.tbPartNumber.Text = FeatOcc.Name
                                myiPropsForm.tbStockNumber.Text = FeatOcc.Name
                                myiPropsForm.tbDescription.Text = FeatOcc.ExtendedName
                            ElseIf TypeOf AssyDoc.SelectSet(1) Is HoleFeature Then
                                myiPropsForm.tbPartNumber.ReadOnly = True
                                myiPropsForm.tbDescription.ReadOnly = True
                                myiPropsForm.tbStockNumber.ReadOnly = True
                                myiPropsForm.tbEngineer.ReadOnly = True
                                Dim holeOcc As HoleFeature = AssyDoc.SelectSet(1)
                                Dim spotsize = Nothing
                                Dim spotdepth = Nothing
                                Dim cbore = Nothing
                                Dim cboredepth = Nothing
                                Dim chamdia = Nothing
                                Dim holedia = Nothing

                                'If holeOcc.Tapped Then
                                '    HoleTap = holeOcc.TapInfo.ThreadDesignation
                                '    'HoleDia = holeOcc.TapInfo.ThreadType
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    myiPropsForm.tbDescription.Text = HoleTap & " " & holedia & " " & spotsize & " " &
                                '     cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbStockNumber.Text = "Tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'Else
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    holedia = "Ø" & holeOcc.HoleDiameter.Value * 10
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    myiPropsForm.tbDescription.Text = holedia & " " & spotsize & " " &
                                '       cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbStockNumber.Text = "Not a tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'End If
                                myiPropsForm.tbEngineer.Text = "Reading Feature Properties"
                                myiPropsForm.tbPartNumber.Text = holeOcc.Name
                                myiPropsForm.tbStockNumber.Text = holeOcc.Name
                                myiPropsForm.tbDescription.Text = holeOcc.ExtendedName
                            Else
                                myiPropsForm.tbPartNumber.ReadOnly = False
                                myiPropsForm.tbDescription.ReadOnly = False
                                myiPropsForm.tbStockNumber.ReadOnly = False
                                myiPropsForm.tbEngineer.ReadOnly = False
                                UpdateDisplayediProperties(AssyDoc)
                            End If
                        ElseIf AssyDoc.SelectSet.Count = 0 Then
                            UpdateDisplayediProperties(AssyDoc)
                            myiPropsForm.tbPartNumber.ReadOnly = False
                            myiPropsForm.tbDescription.ReadOnly = False
                            myiPropsForm.tbStockNumber.ReadOnly = False
                            myiPropsForm.tbEngineer.ReadOnly = False
                        End If
                    ElseIf (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject) Then
                        myiPropsForm.tbDescription.ForeColor = Drawing.Color.Black
                        myiPropsForm.tbPartNumber.ForeColor = Drawing.Color.Black
                        myiPropsForm.tbStockNumber.ForeColor = Drawing.Color.Black
                        myiPropsForm.tbEngineer.ForeColor = Drawing.Color.Black
                        myiPropsForm.tbDrawnBy.ForeColor = Drawing.Color.Black
                        Dim PartDoc As PartDocument = AddinGlobal.InventorApp.ActiveDocument
                        If PartDoc.SelectSet.Count = 1 Then
                            If TypeOf PartDoc.SelectSet(1) Is PartFeature Then
                                myiPropsForm.tbPartNumber.ReadOnly = True
                                myiPropsForm.tbDescription.ReadOnly = True
                                myiPropsForm.tbStockNumber.ReadOnly = True
                                myiPropsForm.tbEngineer.ReadOnly = True
                                Dim FeatOcc As PartFeature = PartDoc.SelectSet(1)
                                Dim spotsize = Nothing
                                Dim spotdepth = Nothing
                                Dim cbore = Nothing
                                Dim cboredepth = Nothing
                                Dim chamdia = Nothing
                                Dim holedia = Nothing

                                'If holeOcc.Tapped Then
                                '    HoleTap = holeOcc.TapInfo.ThreadDesignation
                                '    'HoleDia = holeOcc.TapInfo.ThreadType
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    'myiPropsForm.tbDescription.Text = HoleTap & " " & holedia & " " & spotsize & " " &
                                '    'cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbDescription.Text = holeOcc.ExtendedName
                                '    myiPropsForm.tbStockNumber.Text = "Tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'Else
                                '    If holeOcc.HoleType = HoleTypeEnum.kSpotFaceHole Then
                                '        spotsize = "S'FACE Ø" & holeOcc.SpotFaceDiameter.Value
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterBoreHole Then
                                '        cbore = "C'BORE Ø" & holeOcc.CBoreDiameter.Value
                                '        cboredepth = holeOcc.CBoreDepth.Value & "DEEP"
                                '    ElseIf holeOcc.HoleType = HoleTypeEnum.kCounterSinkHole Then
                                '        chamdia = "C'SINK Ø" & holeOcc.CSinkDiameter.Value
                                '    End If
                                '    holedia = "Ø" & holeOcc.HoleDiameter.Value * 10
                                '    myiPropsForm.tbPartNumber.Text = "This is a HOLE!"
                                '    'myiPropsForm.tbDescription.Text = holedia & " " & spotsize & " " &
                                '    '   cbore & " " & cboredepth & chamdia
                                '    myiPropsForm.tbDescription.Text = holeOcc.ExtendedName
                                '    myiPropsForm.tbStockNumber.Text = "Not a tapped hole"
                                '    myiPropsForm.tbEngineer.Text = "Hole in the head!"
                                'End If
                                myiPropsForm.tbEngineer.Text = "Reading Feature Properties"
                                myiPropsForm.tbPartNumber.Text = FeatOcc.Name
                                myiPropsForm.tbStockNumber.Text = FeatOcc.Name
                                myiPropsForm.tbDescription.Text = FeatOcc.ExtendedName
                            Else
                                myiPropsForm.tbPartNumber.ReadOnly = False
                                myiPropsForm.tbDescription.ReadOnly = False
                                myiPropsForm.tbStockNumber.ReadOnly = False
                                myiPropsForm.tbEngineer.ReadOnly = False
                                UpdateDisplayediProperties()
                            End If
                        Else
                            myiPropsForm.tbPartNumber.ReadOnly = False
                            myiPropsForm.tbDescription.ReadOnly = False
                            myiPropsForm.tbStockNumber.ReadOnly = False
                            myiPropsForm.tbEngineer.ReadOnly = False
                            UpdateDisplayediProperties()
                        End If
                    End If
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnQuit(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kBefore Then
                InventorAppQuitting = True
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnSaveDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
                myiPropsForm.tbDrawnBy.ForeColor = Drawing.Color.Black
                myiPropsForm.GetNewFilePaths()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnActivateDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                m_DocEvents = DocumentObject.DocumentEvents
                AddHandler m_DocEvents.OnChangeSelectSet, AddressOf Me.m_DocumentEvents_OnChangeSelectSet
                Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument

                If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
                    If DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        myiPropsForm.tbPartNumber.ReadOnly = False
                        myiPropsForm.tbDescription.ReadOnly = False
                        myiPropsForm.tbStockNumber.ReadOnly = False
                        myiPropsForm.tbEngineer.ReadOnly = False
                        'AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        'log.Info("Mass Updated correctly")

                    ElseIf DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        myiPropsForm.tbPartNumber.ReadOnly = False
                        myiPropsForm.tbDescription.ReadOnly = False
                        myiPropsForm.tbStockNumber.ReadOnly = False
                        myiPropsForm.tbEngineer.ReadOnly = False
                        'AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        'log.Info("Mass Updated correctly")
                    End If
                End If
                UpdateDisplayediProperties()
                myiPropsForm.tbDescription.ForeColor = Drawing.Color.Black
                myiPropsForm.tbPartNumber.ForeColor = Drawing.Color.Black
                myiPropsForm.tbStockNumber.ForeColor = Drawing.Color.Black
                myiPropsForm.tbEngineer.ForeColor = Drawing.Color.Black
                myiPropsForm.tbDrawnBy.ForeColor = Drawing.Color.Black
                myiPropsForm.GetNewFilePaths()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnOpenDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then

                Dim DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument

                If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
                    If DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        myiPropsForm.tbPartNumber.ReadOnly = False
                        myiPropsForm.tbDescription.ReadOnly = False
                        myiPropsForm.tbStockNumber.ReadOnly = False
                        myiPropsForm.tbEngineer.ReadOnly = False
                        'AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        'log.Info("Mass Updated correctly")

                    ElseIf DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        myiPropsForm.tbPartNumber.ReadOnly = False
                        myiPropsForm.tbDescription.ReadOnly = False
                        myiPropsForm.tbStockNumber.ReadOnly = False
                        myiPropsForm.tbEngineer.ReadOnly = False
                        'AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        'log.Info("Mass Updated correctly")
                    End If
                End If
                UpdateDisplayediProperties()
                myiPropsForm.tbDescription.ForeColor = Drawing.Color.Black
                myiPropsForm.tbPartNumber.ForeColor = Drawing.Color.Black
                myiPropsForm.tbStockNumber.ForeColor = Drawing.Color.Black
                myiPropsForm.tbEngineer.ForeColor = Drawing.Color.Black
                myiPropsForm.tbDrawnBy.ForeColor = Drawing.Color.Black
                myiPropsForm.GetNewFilePaths()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub
        Private Sub m_ApplicationEvents_OnNewDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub



        ''' <summary>
        ''' Need to add more updates here as we add textboxes and therefore properties to this list.
        ''' 
        ''' </summary>
        ''' <param name="DocumentToPulliPropValuesFrom"></param>
        Private Sub UpdateDisplayediProperties(Optional DocumentToPulliPropValuesFrom As Document = Nothing)
            Try
                If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing And DocumentToPulliPropValuesFrom Is Nothing Then
                    DocumentToPulliPropValuesFrom = AddinGlobal.InventorApp.ActiveDocument
                End If
                If DocumentToPulliPropValuesFrom.FullFileName?.Length > 0 Then
                    If AddinGlobal.InventorApp.ActiveEditObject IsNot Nothing Then
                        myiPropsForm.FileLocation.ForeColor = Drawing.Color.Black
                        myiPropsForm.FileLocation.Text = AddinGlobal.InventorApp.ActiveEditDocument.FullFileName
                    Else
                        myiPropsForm.FileLocation.ForeColor = Drawing.Color.Black
                        myiPropsForm.FileLocation.Text = AddinGlobal.InventorApp.ActiveDocument.FullFileName
                    End If

                    If DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                        myiPropsForm.btDefer.Show()
                        myiPropsForm.Label8.Show()
                        myiPropsForm.Label7.Show()
                        myiPropsForm.tbDrawnBy.Show()
                        myiPropsForm.btShtMaterial.Show()
                        myiPropsForm.btShtScale.Show()
                        'myiPropsForm.btUpdateAssy.Show()
                        'myiPropsForm.tbEngineer.Show()
                        myiPropsForm.Label5.Hide()
                        myiPropsForm.tbMass.Hide()
                        myiPropsForm.Label9.Hide()
                        myiPropsForm.tbDensity.Hide()
                        myiPropsForm.Label3.Hide()
                        myiPropsForm.tbStockNumber.Hide()
                        'myiPropsForm.Label4.Show()
                        'myiPropsForm.Button6.Show()
                        'myiPropsForm.Button7.Show()
                        'myiPropsForm.Button8.Show()
                        'myiPropsForm.Button9.Show()

                        myiPropsForm.tbDrawnBy.Text = iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForSummaryInformationEnum.kAuthorSummaryInformation, "", "")

                        If iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                            myiPropsForm.Label8.ForeColor = Drawing.Color.Red
                            myiPropsForm.Label8.Text = "Drawing Updates Deferred"
                        ElseIf iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                            myiPropsForm.Label8.ForeColor = Drawing.Color.Green
                            myiPropsForm.Label8.Text = "Drawing Updates Not Deferred"
                        End If

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

                        Dim MaterialString As String = String.Empty

                        MaterialString = iProperties.GetorSetStandardiProperty(
                        drawnDoc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")

                        myiPropsForm.Label12.Text = MaterialString

                        If iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbPartNumber.Text = "Part Number" Then
                            iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        End If

                        If iProperties.GetorSetStandardiProperty(
                            drawnDoc,
                            PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbDescription.Text = "Description" Then
                            iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        End If

                        If iProperties.GetorSetStandardiProperty(
                            drawnDoc,
                            PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbEngineer.Text = "Engineer" Then
                            iProperties.GetorSetStandardiProperty(
                            drawnDoc,
                            PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(
                                drawnDoc,
                                PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        End If

                    Else
                        'MassProps()
                        'AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        'log.Info("Mass Updated correctly")
                        myiPropsForm.Label5.Show()
                        myiPropsForm.tbMass.Show()
                        myiPropsForm.Label9.Show()
                        myiPropsForm.tbDensity.Show()
                        myiPropsForm.Label3.Show()
                        myiPropsForm.tbStockNumber.Show()
                        'myiPropsForm.Label4.Show()
                        'myiPropsForm.tbEngineer.Show()
                        'myiPropsForm.Button6.Show()
                        'myiPropsForm.Button7.Show()
                        'myiPropsForm.Button8.Show()
                        'myiPropsForm.Button9.Show()
                        myiPropsForm.Label7.Hide()
                        myiPropsForm.tbDrawnBy.Hide()
                        myiPropsForm.btDefer.Hide()
                        myiPropsForm.Label8.Hide()
                        myiPropsForm.btShtMaterial.Hide()
                        myiPropsForm.btShtScale.Hide()
                        'myiPropsForm.btUpdateAssy.Hide()

                        If iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbStockNumber.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbStockNumber.Text = "Stock Number" Then
                            iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbStockNumber.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        End If

                        If iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbEngineer.Text = "Engineer" Then
                            iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbEngineer.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        End If

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

                        If iProperties.GetorSetStandardiProperty(
                               DocumentToPulliPropValuesFrom,
                               PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbPartNumber.Text = "Part Number" Then
                            iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbPartNumber.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                        End If

                        If iProperties.GetorSetStandardiProperty(
                            DocumentToPulliPropValuesFrom,
                            PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.tbDescription.Text = "Description" Then
                            iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.tbDescription.Text = iProperties.GetorSetStandardiProperty(
                                DocumentToPulliPropValuesFrom,
                                PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                        End If

                    End If

                    If DocumentToPulliPropValuesFrom.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        myiPropsForm.btITEM.Show()
                        myiPropsForm.Label11.Hide()
                        myiPropsForm.Label12.Hide()
                    Else
                        myiPropsForm.btITEM.Hide()
                        myiPropsForm.Label11.Show()
                        myiPropsForm.Label12.Show()
                    End If

                    myiPropsForm.DateTimePicker1.Value = iProperties.GetorSetStandardiProperty(
                        DocumentToPulliPropValuesFrom,
                        PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties, "", "")

                    If (AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                        Dim AssyDoc As AssemblyDocument = AddinGlobal.InventorApp.ActiveDocument
                        If AssyDoc.SelectSet.Count = 1 Then
                            If TypeOf AssyDoc.SelectSet(1) Is ComponentOccurrence Then
                                If CheckReadOnly(DocumentToPulliPropValuesFrom) Then

                                    myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                                    myiPropsForm.Label10.Text = "Checked In"
                                    myiPropsForm.PictureBox1.Show()
                                    myiPropsForm.PictureBox2.Hide()

                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                Else
                                    myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                                    myiPropsForm.Label10.Text = "Checked Out"
                                    myiPropsForm.PictureBox1.Hide()
                                    myiPropsForm.PictureBox2.Show()

                                    myiPropsForm.tbPartNumber.ReadOnly = True
                                    myiPropsForm.tbDescription.ReadOnly = True
                                    myiPropsForm.tbStockNumber.ReadOnly = True
                                    myiPropsForm.tbEngineer.ReadOnly = True
                                End If
                            End If
                        Else
                            If CheckReadOnly(DocumentToPulliPropValuesFrom) Then
                                myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                                myiPropsForm.Label10.Text = "Checked In"
                                myiPropsForm.PictureBox1.Show()
                                myiPropsForm.PictureBox2.Hide()

                                myiPropsForm.tbPartNumber.ReadOnly = True
                                myiPropsForm.tbDescription.ReadOnly = True
                                myiPropsForm.tbStockNumber.ReadOnly = True
                                myiPropsForm.tbEngineer.ReadOnly = True
                            Else
                                myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                                myiPropsForm.Label10.Text = "Checked Out"
                                myiPropsForm.PictureBox1.Hide()
                                myiPropsForm.PictureBox2.Show()

                                myiPropsForm.tbPartNumber.ReadOnly = False
                                myiPropsForm.tbDescription.ReadOnly = False
                                myiPropsForm.tbStockNumber.ReadOnly = False
                                myiPropsForm.tbEngineer.ReadOnly = False
                            End If
                        End If
                    Else
                        If CheckReadOnly(DocumentToPulliPropValuesFrom) Then
                            myiPropsForm.Label10.ForeColor = Drawing.Color.Red
                            myiPropsForm.Label10.Text = "Checked In"
                            myiPropsForm.PictureBox1.Show()
                            myiPropsForm.PictureBox2.Hide()

                            myiPropsForm.tbPartNumber.ReadOnly = True
                            myiPropsForm.tbDescription.ReadOnly = True
                            myiPropsForm.tbStockNumber.ReadOnly = True
                            myiPropsForm.tbEngineer.ReadOnly = True
                        Else
                            myiPropsForm.Label10.ForeColor = Drawing.Color.Green
                            myiPropsForm.Label10.Text = "Checked Out"
                            myiPropsForm.PictureBox1.Hide()
                            myiPropsForm.PictureBox2.Show()

                            myiPropsForm.tbPartNumber.ReadOnly = False
                            myiPropsForm.tbDescription.ReadOnly = False
                            myiPropsForm.tbStockNumber.ReadOnly = False
                            myiPropsForm.tbEngineer.ReadOnly = False
                        End If
                    End If
                End If
            Catch ex As Exception
                log.Error(ex.Message)
            End Try
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
                thisAssembly = Nothing
                myiPropsForm = Nothing
                m_UserInputEvents = Nothing
                m_AppEvents = Nothing
                m_uiEvents = Nothing
                m_StyleEvents = Nothing

                myiPropsForm.CurrentPath = Nothing
                myiPropsForm.NewPath = Nothing
                myiPropsForm.RefNewPath = Nothing
                myiPropsForm.RefDoc = Nothing

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
            Dim t As Type = GetType(iPropertiesController.StandardAddInServer)
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