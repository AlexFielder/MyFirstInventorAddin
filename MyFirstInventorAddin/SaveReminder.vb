Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Inventor
Imports System.ComponentModel.Composition
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports log4net

Namespace TimeSinceLastSave
    Public Class TimeSinceLastSave

        Private thisAssembly As Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Private Shared logHelper As Log4NetFileHelper.Log4NetFileHelper = New Log4NetFileHelper.Log4NetFileHelper()
        Public Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(TimeSinceLastSave))
        Public Shared m_InventorApp As Application = Nothing
        Private m_uiEvents As UserInterfaceEvents = Nothing
        Private m_UserInputEvents As UserInputEvents = Nothing
        Private m_TimeSinceLastSaveAppEvents As ApplicationEvents = Nothing
        Private m_FileAccesEvents As FileAccessEvents = Nothing
        Private DirtiedTime As DateTime
        Private m_ActiveDoc As Document

        Public ReadOnly Property CommandDisplayName As String
            Get
                Return "Time Since Last Save"
            End Get
        End Property

        Public ReadOnly Property CommandInternalName As String
            Get
                Dim attribute = CType(thisAssembly.GetCustomAttributes(GetType(GuidAttribute), True)(0), GuidAttribute)
                Dim AssemblyGuid As String = attribute.Value
                Return AssemblyGuid
            End Get
        End Property

        Public ReadOnly Property CommandName As String
            Get
                Return "TimeSinceLastSave"
            End Get
        End Property

        Public ReadOnly Property CommandType As CommandTypesEnum
            Get
                Return CommandTypesEnum.kQueryOnlyCmdType
            End Get
        End Property

        Public ReadOnly Property DefaultResourceName As String
            Get
                Return "TimeSinceLastSave.Resources.TurdPolish.ico"
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return "This is the Description"
            End Get
        End Property

        Public ReadOnly Property DisplayBigIcon As Boolean
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property DisplayText As Boolean
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property InsertBeforeTarget As Boolean
            Get
                Return True
            End Get
        End Property

        Public Property InventorApp As Application
            Get
                Return m_InventorApp
            End Get
            Set(ByVal value As Application)
                m_InventorApp = value
            End Set
        End Property

        Public Property LogFileHelper As Log4NetFileHelper.Log4NetFileHelper
            Get
                Return logHelper
            End Get
            Set(ByVal value As Log4NetFileHelper.Log4NetFileHelper)
                logHelper = value
            End Set
        End Property

        Public ReadOnly Property Path As String
            Get
                Dim thisAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Return thisAssembly.Location
            End Get
        End Property

        Public ReadOnly Property TargetControlName As String
            Get
                Return "PartMoveEOPMarkerCmd"
            End Get
        End Property

        Public ReadOnly Property ToolTip As String
            Get
                Return "This is the Tooltip"
            End Get
        End Property

        Public ReadOnly Property color As Color
            Get
                Return InventorApp.TransientObjects.CreateColor(117, 45, 86)
            End Get
        End Property

        Public ReadOnly Property ClientGraphicsCollectionName As String
            Get
                Return "TimeSinceLastSave"
            End Get
        End Property

        Public Property GenericAppEventHandler As ApplicationEvents
            Get
                Return m_TimeSinceLastSaveAppEvents
            End Get
            Set(ByVal value As ApplicationEvents)
                m_TimeSinceLastSaveAppEvents = value
                AddHandler m_TimeSinceLastSaveAppEvents.OnSaveDocument, AddressOf M_TimeSinceLastSaveAppEventsDocumentSaved
                AddHandler m_TimeSinceLastSaveAppEvents.OnActivateDocument, AddressOf M_TimeSinceLastSaveDocActivated
            End Set
        End Property

        Public Property GenericUserInteractionEventHandler As UserInterfaceEvents
            Get
                Return m_uiEvents
            End Get
            Set(ByVal value As UserInterfaceEvents)
                m_uiEvents = value
            End Set
        End Property

        Public Property GenericUserInputEventHandler As UserInputEvents
            Get
                Return m_UserInputEvents
            End Get
            Set(ByVal value As UserInputEvents)
                m_UserInputEvents = value
                AddHandler m_UserInputEvents.OnTerminateCommand, AddressOf M_UserInputEvents_OnTerminateCommand
            End Set
        End Property

        Public Property GenericFileAccessEventHandler As FileAccessEvents
            Get
                Return m_FileAccesEvents
            End Get
            Set(ByVal value As FileAccessEvents)
                m_FileAccesEvents = value
                AddHandler m_FileAccesEvents.OnFileDirty, AddressOf M_FileAccesEvents_OnFileDirty
            End Set
        End Property

        Public Property GenericActiveDocument As Document
            Get
                Return InventorApp.ActiveDocument
            End Get
            Set(ByVal value As Document)
                GenericActiveDocument = value
            End Set
        End Property

        Public ReadOnly Property PluginSettingsPrefix As String
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
        End Property

        Public Property PluginSettingsPrefixVar As String
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
            Set(ByVal value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property ParentSettingsFilePath As String
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
            Set(ByVal value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property m_SelectEvents As SelectEvents
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
            Set(ByVal value As SelectEvents)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property m_MouseEvents As MouseEvents
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
            Set(ByVal value As MouseEvents)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property m_TriadEvents As TriadEvents
            Get
                Return CSharpImpl.__Throw(Of Object)(New NotImplementedException())
            End Get
            Set(ByVal value As TriadEvents)
                Throw New NotImplementedException()
            End Set
        End Property

        Private Sub M_FileAccesEvents_OnFileDirty(ByVal RelativeFileName As String, ByVal LibraryName As String, ByRef CustomLogicalName As Byte(), ByVal FullFileName As String, ByVal DocumentObject As _Document, ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, <Out> ByRef HandlingCode As HandlingCodeEnum)
            Try
                Dim dirtiedAttSet As AttributeSet = Nothing

                If BeforeOrAfter = EventTimingEnum.kBefore Then
                    DirtiedTime = DateTime.Now
                    dirtiedAttSet = DocumentObject.AttributeSets.Add("DirtiedTimeSet")
                    Dim oAtt As Inventor.Attribute = Nothing
                    Dim dirtiedTimeStr As String = String.Format("{0}", DirtiedTime)
                    oAtt = dirtiedAttSet.Add("DirtiedTime", ValueTypeEnum.kStringType, dirtiedTimeStr)
                    oAtt = dirtiedAttSet.Add("User", ValueTypeEnum.kStringType, InventorApp.UserName)
                End If

                Dim timeSincelastsave As TimeSpan = GetTimeSinceFileDirtied()
                If timeSincelastsave.Seconds > 5 Then
                    SaveChecker = MsgBox("You haven't saved in a while, would you like to now?", vbYesNo, "Save Checker")
                    If SaveChecker = vbYes Then
                        oDocu = InventorApp.ActiveDocument
                        oDocu.Save2(True)
                        InventorApp.ActiveView.Update()
                    Else
                        InventorApp.ActiveView.Update()
                    End If
                End If

            Catch ex As Exception
                log.[Error](ex.Message, ex)
            End Try

            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub M_UserInputEvents_OnTerminateCommand(ByVal CommandName As String, ByVal Context As NameValueMap)
            If InventorApp.ActiveDocument IsNot Nothing AndAlso InventorApp.ActiveDocument.Dirty Then
                DrawTimerGraphics(InventorApp.ActiveView)
            End If
        End Sub

        Private Sub M_UserInputEvents_OnActivateCommand(ByVal CommandName As String, ByVal Context As NameValueMap)
            Throw New NotImplementedException()
        End Sub

        Public Sub Execute()
            m_TimeSinceLastSaveAppEvents = m_InventorApp.ApplicationEvents
            AddHandler m_TimeSinceLastSaveAppEvents.OnActivateView, AddressOf M_TimeSinceLastSaveAppEvents_OnActivateView
        End Sub

        Private Sub M_TimeSinceLastSaveAppEvents_OnActivateView(ByVal ViewObject As View, ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, <Out> ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                DrawTimerGraphics(ViewObject)
            End If

            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub M_TimeSinceLastSaveDocActivated(ByVal DocumentObject As _Document, ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, <Out> ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                DrawTimerGraphics(InventorApp.ActiveView)
            End If

            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub M_TimeSinceLastSaveAppEventsDocumentSaved(ByVal DocumentObject As _Document, ByVal BeforeOrAfter As EventTimingEnum, ByVal Context As NameValueMap, <Out> ByRef HandlingCode As HandlingCodeEnum)
            If BeforeOrAfter = EventTimingEnum.kAfter Then
                DrawTimerGraphics()
            Else
                CheckForAndClearExistingAttributes()
            End If

            HandlingCode = HandlingCodeEnum.kEventNotHandled
        End Sub

        Private Sub DrawTimerGraphics(ByVal Optional activeView As View = Nothing)
            Dim trans As Transaction = InventorApp.TransactionManager.StartTransaction(InventorApp.ActiveDocument, "Timergraphics")

            Try

                If activeView IsNot Nothing Then
                    Dim compDef As ComponentDefinition = InventorApp.ActiveEditDocument
                    Dim clientGfx As ClientGraphics = Nothing

                    Try
                        clientGfx = compDef.ClientGraphicsCollection("TimeSinceLastSave")
                        clientGfx.Delete()
                        InventorApp.ActiveView.Update()
                        clientGfx = compDef.ClientGraphicsCollection.Add("TimeSinceLastSave")
                    Catch __unusedException1__ As Exception
                        clientGfx = compDef.ClientGraphicsCollection.Add("TimeSinceLastSave")
                    End Try

                    Dim gfxNode As GraphicsNode = clientGfx.AddNode(1)
                    Dim textGfx As TextGraphics = gfxNode.AddTextGraphics()
                    Dim currentDoc As Document = InventorApp.ActiveDocument
                    Dim formattedTimespan As String = String.Empty

                    If currentDoc.FullFileName <> String.Empty Then

                        If currentDoc.Dirty Then
                            formattedTimespan = "Time Since File Dirtied: " & GetFormattedTimespan()
                            textGfx.PutTextColor(117, 45, 50)
                        Else
                            formattedTimespan = "Date Last Saved: " & GetLastSavedDate(currentDoc)
                            textGfx.PutTextColor(0, 122, 0)
                        End If
                    Else
                        formattedTimespan = "Not Saved!"
                    End If

                    textGfx.Text = formattedTimespan
                    textGfx.Anchor = InventorApp.TransientGeometry.CreatePoint(0, 0, 0)
                    Dim textXoffset As Double = formattedTimespan.Length
                    textGfx.SetViewSpaceAnchor(textGfx.Anchor, InventorApp.TransientGeometry.CreatePoint2d(250 + textXoffset, 30), ViewLayoutEnum.kBottomRightViewCorner)
                    InventorApp.ActiveView.Update()
                End If

            Catch __unusedException1__ As Exception
                trans.Abort()
                Throw
            Finally
                trans.[End]()
            End Try
        End Sub

        Private Function GetLastSavedDate(ByVal activeDocument As Document) As String
            Dim dateLastSaved As DateTime = System.IO.File.GetLastWriteTime(activeDocument.FullFileName)
            Return String.Format("{0}", dateLastSaved)
        End Function

        Private Sub CheckForAndClearExistingAttributes()
            Dim objCollection As ObjectCollection = GenericActiveDocument.AttributeManager.FindObjects("DirtiedTimeSet", "*")

            For i As Integer = 1 To objCollection.Count + 1 - 1

                If TypeOf objCollection(i) Is Document Then
                    Dim thisDoc As Document = CType(objCollection(i), Document)
                    thisDoc.AttributeSets(1).Delete()
                End If
            Next

            If objCollection.Count = 0 Then
                Dim attribEnum As AttributesEnumerator = GenericActiveDocument.AttributeManager.FindAttributes("DirtiedTimeSet", "*")

                For Each item As Inventor.Attribute In attribEnum
                    item.Delete()
                Next

                Dim attribSetsEnum As AttributeSetsEnumerator = GenericActiveDocument.AttributeManager.FindAttributeSets("DirtiedTimeSet")

                For Each attSet As AttributeSet In attribSetsEnum
                    attSet.Delete()
                Next
            End If
        End Sub

        Private Function GetFormattedTimespan() As String
            Dim formattedTimespan As String
            Dim timeSincelastsave As TimeSpan = GetTimeSinceFileDirtied()
            formattedTimespan = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", timeSincelastsave.Hours, timeSincelastsave.Minutes, timeSincelastsave.Seconds, timeSincelastsave.Milliseconds / 10)
            Return formattedTimespan
        End Function

        Private Function GetTimeSinceLastSave(ByVal activeDocument As _Document) As TimeSpan
            Dim dateLastSaved As DateTime = System.IO.File.GetLastWriteTime(activeDocument.FullFileName)
            Return DateTime.Now - dateLastSaved
        End Function

        Private Function GetTimeSinceFileDirtied() As TimeSpan
            Dim objCollection As ObjectCollection = GenericActiveDocument.AttributeManager.FindObjects("DirtiedTime", "DirtiedTime")
            Dim storedDirtiedTime As DateTime

            For i As Integer = 1 To objCollection.Count + 1 - 1

                If TypeOf objCollection(i) Is Document Then
                    Dim thisDoc As Document = CType(objCollection(i), Document)
                    Dim AttSet As AttributeSet = thisDoc.AttributeSets(1)
                    Dim dirtiedTimeAtt As Inventor.Attribute = AttSet("DirtiedTime")
                    storedDirtiedTime = dirtiedTimeAtt.Value
                    Return DateTime.Now - storedDirtiedTime
                Else
                    Return DateTime.Now - DirtiedTime
                End If
            Next

            Return DateTime.Now - DirtiedTime
        End Function

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal throw statements")>
            Shared Function __Throw(Of T)(ByVal e As Exception) As T
                Throw e
            End Function
        End Class
    End Class
End Namespace