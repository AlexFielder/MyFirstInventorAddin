Imports Inventor
Imports System.Runtime.InteropServices
Imports log4net
Imports System.Drawing
Imports System.Reflection
Imports MyFirstInventorAddin.yFirstInventorAddin

Namespace MyFirstInventorAddin
    <ProgIdAttribute("MyFirstInventorAddin.StandardAddInServer"), _
    GuidAttribute("e691af34-cb32-4296-8cca-5d1027a27c72")> _
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

                AddHandler m_AppEvents.OnOpenDocument, AddressOf Me.m_ApplicationEvents_OnOpenDocument
                AddHandler m_AppEvents.OnActivateDocument, AddressOf Me.m_ApplicationEvents_OnActivateDocument
                AddHandler m_AppEvents.OnSaveDocument, AddressOf Me.m_ApplicationEvents_OnSaveDocument
                AddHandler m_AppEvents.OnQuit, AddressOf Me.m_ApplicationEvents_OnQuit
                'you can add extra handlers like this - if you uncomment the next line Visual Studio will prompt you to create the method:
                'AddHandler m_AssemblyEvents.OnNewOccurrence, AddressOf Me.m_AssemblyEvents_NewOcccurrence
                'AddHandler m_DocEvents.OnChangeSelectSet, AddressOf Me.m_DocumentEvents_OnChangeSelectSet
                'AddHandler m_AppEvents.OnNewDocument, AddressOf Me.m_ApplicationEvents_OnNewDocument
                'AddHandler m_StyleEvents.OnActivateStyle, AddressOf Me.m_StyleEvents_OnActivateStyle

                'start our logger.
                logHelper.Init()
                logHelper.AddFileLogging(IO.Path.Combine(thisAssemblyPath, "MyFirstInventorAddin.log"))
                logHelper.AddFileLogging("C:\Logs\MyLogFile.txt", Core.Level.All, True)
                logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", Core.Level.All, True)
                log.Debug("Loading My First Inventor Addin")
                ' TODO: Add button definitions.

                ' Sample to illustrate creating a button definition.
                'Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourBigImage)
                'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
                'Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
                'm_sampleButton = controlDefs.AddButtonDefinition("Command Name", "Internal Name", CommandTypesEnum.kShapeEditCmdType, AddInClientID)

                Dim icon1 As New Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("MyFirstInventorAddin.addin.ico"))
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

        Private Sub m_StyleEvents_OnActivateStyle(DocumentObject As _Document, Style As Object, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            Throw New NotImplementedException()
        End Sub

        Private Sub m_ApplicationEvents_OnQuit(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kBefore Then
                InventorAppQuitting = True
            End If
        End Sub

        Private Sub m_ApplicationEvents_OnSaveDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnActivateDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub m_ApplicationEvents_OnOpenDocument(DocumentObject As _Document, FullDocumentName As String, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub
        Private Sub m_ApplicationEvents_OnNewDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                UpdateDisplayediProperties()
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub
        Private Sub m_DocumentEvents_OnChangeSelectSet(BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    UpdateDisplayediProperties()
                End If
            End If
            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub
        ''' <summary>
        ''' Need to add more updates here as we add textboxes and therefore properties to this list.
        ''' </summary>
        Private Sub UpdateDisplayediProperties()
            If Not AddinGlobal.InventorApp.ActiveDocument Is Nothing Then
                If AddinGlobal.InventorApp.ActiveDocument.FullFileName?.Length > 0 Then
                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "").Length > 0 Then
                        myiPropsForm.TextBox1.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                    ElseIf myiPropsForm.TextBox1.Text = "Part Number" Then
                        iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                    Else
                        myiPropsForm.TextBox1.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties, "", "")
                    End If

                    If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "").Length > 0 Then
                        myiPropsForm.TextBox2.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    ElseIf myiPropsForm.TextBox2.Text = "Description" Then
                        iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    Else
                        myiPropsForm.TextBox2.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties, "", "")
                    End If

                    If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        myiPropsForm.Label3.Show()
                        myiPropsForm.TextBox3.Show()
                        myiPropsForm.Label4.Show()
                        myiPropsForm.TextBox4.Show()
                        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.TextBox3.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.TextBox3.Text = "Stock Number" Then
                            iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.TextBox3.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kStockNumberDesignTrackingProperties, "", "")
                        End If

                        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "").Length > 0 Then
                            myiPropsForm.TextBox4.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        ElseIf myiPropsForm.TextBox4.Text = "Engineer" Then
                            iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        Else
                            myiPropsForm.TextBox4.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kEngineerDesignTrackingProperties, "", "")
                        End If
                    Else
                        myiPropsForm.Label3.Hide()
                        myiPropsForm.TextBox3.Hide()
                        myiPropsForm.Label4.Hide()
                        myiPropsForm.TextBox4.Hide()
                    End If



                    If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                        myiPropsForm.Label5.Hide()
                        myiPropsForm.TextBox5.Hide()
                        myiPropsForm.Button2.Show()
                        myiPropsForm.Label8.Show()
                        myiPropsForm.Label9.Hide()
                        myiPropsForm.TextBox6.Hide()
                        myiPropsForm.Label7.Show()
                        myiPropsForm.TextBox7.Show()

                        myiPropsForm.TextBox7.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForSummaryInformationEnum.kAuthorSummaryInformation, "", "")

                        If iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = True Then
                            myiPropsForm.Label8.Text = "Drawing Updates Deferred"
                        ElseIf iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDrawingDeferUpdateDesignTrackingProperties, "", "") = False Then
                            myiPropsForm.Label8.Text = "Drawing Updates Not Deferred"
                        End If

                    Else
                        AddinGlobal.InventorApp.CommandManager.ControlDefinitions.Item("AppUpdateMassPropertiesCmd").Execute()
                        myiPropsForm.Label5.Show()
                        myiPropsForm.TextBox5.Show()
                        myiPropsForm.Button2.Hide()
                        myiPropsForm.Label8.Hide()
                        myiPropsForm.Label9.Show()
                        myiPropsForm.TextBox6.Show()
                        myiPropsForm.Label7.Hide()
                        myiPropsForm.TextBox7.Hide()

                        Dim myMass As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMassDesignTrackingProperties, "", "")
                        Dim kgMass As Decimal = myMass / 1000
                        Dim myMass2 As Decimal = Math.Round(kgMass, 3)
                        myiPropsForm.TextBox5.Text = myMass2 & " kg"
                        Dim myDensity As Decimal = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kDensityDesignTrackingProperties, "", "")
                        Dim myDensity2 As Decimal = Math.Round(myDensity, 3)
                        myiPropsForm.TextBox6.Text = myDensity2 & " g/cm^3"
                    End If

                    If AddinGlobal.InventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        myiPropsForm.Button3.Show()
                    Else
                        myiPropsForm.Button3.Hide()
                    End If

                    myiPropsForm.DateTimePicker1.Value = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kCreationDateDesignTrackingProperties, "", "")
                    myiPropsForm.Label12.Text = iProperties.GetorSetStandardiProperty(AddinGlobal.InventorApp.ActiveDocument, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", "")

                End If
            End If
        End Sub

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

                m_DocEvents = Nothing
                m_AssemblyEvents = Nothing
                m_PartEvents = Nothing
                m_ModelingEvents = Nothing
                m_StyleEvents = Nothing

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
            Dim t As Type = GetType(MyFirstInventorAddin.StandardAddInServer)
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
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
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

            <StructLayout(LayoutKind.Sequential)> _
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
