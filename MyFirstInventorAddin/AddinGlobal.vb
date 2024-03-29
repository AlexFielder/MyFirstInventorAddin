﻿Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor
Imports iPropertiesController.iPropertiesController

Public Class AddinGlobal

    Public Shared InventorApp As Application

    Public Shared RibbonPanelId As String

    Public Shared RibbonPanel As RibbonPanel

    Public Shared ButtonList As List(Of InventorButton) = New List(Of InventorButton)()

    Public Shared DockableList As List(Of DockableWindow) = New List(Of DockableWindow)()

    Public Shared DisplayableVersion As String

    'color objects because of Inventor's Themes
    Public Shared BackColour As Drawing.Color
    Public Shared ForeColour As Drawing.Color
    Public Shared ControlBackColour As Drawing.Color
    Public Shared ControlHighlightedColour As Drawing.Color

    'Not implemented here
    'Public Shared PluginList As List(Of IVPlugin) = New List(Of IVPlugin)()

    Private Shared mClassId As String

    Public Shared Property ClassId As String
        Get
            If Not String.IsNullOrEmpty(mClassId) Then Return AddinGlobal.mClassId Else Throw New System.Exception("The addin server class id hasn't been gotten yet!")
        End Get

        Set(ByVal value As String)
            AddinGlobal.mClassId = value
        End Set
    End Property

    Public Shared Sub GetAddinClassId(ByVal t As Type)
        Dim guidAtt As GuidAttribute = CType(GuidAttribute.GetCustomAttribute(t, GetType(GuidAttribute)), GuidAttribute)
        mClassId = "{" & guidAtt.Value & "}"
    End Sub

End Class