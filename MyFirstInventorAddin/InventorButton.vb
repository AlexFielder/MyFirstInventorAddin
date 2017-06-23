Imports System.Drawing
Imports System.Windows.Forms
Imports Inventor

Namespace iPropertiesController

    Public Class InventorButton

        Private mButtonDef As ButtonDefinition

        Public Property ButtonDef As ButtonDefinition
            Get
                Return mButtonDef
            End Get

            Set(ByVal value As ButtonDefinition)
                mButtonDef = value
            End Set
        End Property

        Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon, ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
            internalName = internalName
            Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, standardIcon, largeIcon, commandType, buttonDisplayType)
        End Sub

        Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon)
            Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing, Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
        End Sub

        Public Sub New(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
            Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing, Nothing, commandType, buttonDisplayType)
        End Sub

        Public Sub New(ByVal displayName As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon)
            Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, standardIcon, largeIcon, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
        End Sub

        Public Sub New(ByVal displayName As String)
            Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, Nothing, Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
        End Sub

        Public Sub Create(ByVal displayName As String, ByVal internalName As String, ByVal description As String, ByVal tooltip As String, ByVal clientId As String, ByVal standardIcon As Icon, ByVal largeIcon As Icon, ByVal commandType As CommandTypesEnum, ByVal buttonDisplayType As ButtonDisplayEnum)
            If String.IsNullOrEmpty(clientId) Then clientId = AddinGlobal.ClassId
            Dim standardIconIPictureDisp As stdole.IPictureDisp = Nothing
            Dim largeIconIPictureDisp As stdole.IPictureDisp = Nothing
            If standardIcon IsNot Nothing Then
                standardIconIPictureDisp = IconToPicture(standardIcon)
                largeIconIPictureDisp = IconToPicture(largeIcon)
            End If

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, internalName, commandType, clientId, description, tooltip, standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType)
            mButtonDef.Enabled = True
            AddHandler mButtonDef.OnExecute, AddressOf ButtonDefinition_OnExecute
            DisplayText = True
            AddinGlobal.ButtonList.Add(Me)
        End Sub

        Public Property DisplayBigIcon As Boolean

        Public Property DisplayText As Boolean

        Public Property InsertBeforeTarget As Boolean

        Public Property TargetControlName As String

        Public Property InternalName As String

        Public Sub SetBehavior(ByVal displayBigIcon As Boolean, ByVal displayText As Boolean, ByVal insertBeforeTarget As Boolean)
            displayBigIcon = displayBigIcon
            displayText = displayText
            insertBeforeTarget = insertBeforeTarget
        End Sub

        Public Sub CopyBehaviorFrom(ByVal button As InventorButton)
            DisplayBigIcon = button.DisplayBigIcon
            DisplayText = button.DisplayText
            InsertBeforeTarget = Me.InsertBeforeTarget
        End Sub

        Private Sub ButtonDefinition_OnExecute(ByVal context As NameValueMap)
            If Execute IsNot Nothing Then Execute() Else MessageBox.Show("Nothing to execute.")
        End Sub

        Public Execute As Action

        Public Shared Function ImageToPicture(ByVal image As Image) As stdole.IPictureDisp
            Return ImageConverter.ImageToPicture(image)
        End Function

        Public Shared Function IconToPicture(ByVal icon As Icon) As stdole.IPictureDisp
            Return ImageConverter.ImageToPicture(icon.ToBitmap())
        End Function

        Public Shared Function PictureToImage(ByVal picture As stdole.IPictureDisp) As Image
            Return ImageConverter.PictureToImage(picture)
        End Function

        Public Shared Function PictureToIcon(ByVal picture As stdole.IPictureDisp) As Icon
            Return ImageConverter.PictureToIcon(picture)
        End Function

        Private Class ImageConverter
            Inherits AxHost

            Public Sub New()
                MyBase.New(String.Empty)
            End Sub

            Public Shared Function ImageToPicture(ByVal image As Image) As stdole.IPictureDisp
                Return CType(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
            End Function

            Public Shared Function IconToPicture(ByVal icon As Icon) As stdole.IPictureDisp
                Return ImageToPicture(icon.ToBitmap())
            End Function

            Public Shared Function PictureToImage(ByVal picture As stdole.IPictureDisp) As Image
                Return GetPictureFromIPicture(picture)
            End Function

            Public Shared Function PictureToIcon(ByVal picture As stdole.IPictureDisp) As Icon
                Dim bitmap As Bitmap = New Bitmap(PictureToImage(picture))
                Return System.Drawing.Icon.FromHandle(bitmap.GetHicon())
            End Function

        End Class

    End Class

End Namespace