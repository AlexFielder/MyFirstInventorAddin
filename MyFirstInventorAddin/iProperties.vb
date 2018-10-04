Imports Inventor

Namespace iPropertiesController

    Public Class iProperties

        Public Shared Function GetiPropertyDisplayName(ByVal iProp As Inventor.Property) As String
            Return iProp.DisplayName
        End Function

        Public Shared Function GetiPropertyType(ByVal iProp As Inventor.Property) As ObjectTypeEnum
            Return iProp.Type
        End Function

        Public Shared Function GetiPropertyTypeString(ByVal iProp As Inventor.Property) As String
            Dim valToTest As Object = iProp.Value
            Dim intResult As Integer = Nothing
            If Integer.TryParse(iProp.Value, intResult) Then
                Return "Number"
            End If

            Dim doubleResult As Double = Nothing
            If Double.TryParse(iProp.Value, doubleResult) Then
                Return "Number"
            End If

            Dim dateResult As Date = Nothing
            If Date.TryParse(iProp.Value, dateResult) Then
                Return "Date"
            End If

            Dim booleanResult As Boolean = Nothing
            If Boolean.TryParse(iProp.Value, booleanResult) Then
                Return "Boolean"
            End If

            'Dim currencyResult As Currency = Nothing

            'should probabyl do this last as most property values will equate to string!
            Dim strResult As String = String.Empty
            If Not iProp.Value.ToString() = String.Empty Then
                Return "String"
            End If
            Return Nothing
        End Function

#Region "Set iProperty Values"

#Region "Get or Set Standard iProperty Values"

        ''' <summary>
        ''' Design Tracking Properties
        ''' </summary>
        ''' <param name="DocToUpdate"></param>
        ''' <param name="iPropertyTypeEnum"></param>
        ''' <param name="newpropertyvalue"></param>
        ''' <returns></returns>
        Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                         ByVal iPropertyTypeEnum As PropertiesForDesignTrackingPropertiesEnum,
                                                         Optional ByRef newpropertyvalue As String = "",
                                                         Optional ByRef propertyTypeStr As String = "",
                                                         Optional ByVal IsUpdating As Boolean = False) As String
            Dim invProjProperties As PropertySet = DocToUpdate.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
            Dim currentvalue As String = invProjProperties.ItemByPropId(iPropertyTypeEnum).Value
            If IsUpdating Then
                invProjProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
                Return newpropertyvalue
            Else
                Return currentvalue
            End If

        End Function

        ''' <summary>
        ''' Document Summary Properties
        ''' </summary>
        ''' <param name="DocToUpdate"></param>
        ''' <param name="iPropertyTypeEnum"></param>
        ''' <param name="newpropertyvalue"></param>
        ''' <returns></returns>
        Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                         ByVal iPropertyTypeEnum As PropertiesForDocSummaryInformationEnum,
                                                         Optional ByRef newpropertyvalue As String = "",
                                                         Optional ByRef propertyTypeStr As String = "",
                                                         Optional ByVal IsUpdating As Boolean = False) As String
            Dim invDocSummaryProperties As PropertySet = DocToUpdate.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
            Dim currentvalue As String = invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum).Value
            If IsUpdating Then
                invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
                Return newpropertyvalue
            Else
                Return currentvalue
            End If
        End Function

        ''' <summary>
        ''' Summary Properties
        ''' </summary>
        ''' <param name="DocToUpdate"></param>
        ''' <param name="iPropertyTypeEnum"></param>
        ''' <param name="newpropertyvalue"></param>
        ''' <returns></returns>
        Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                         ByVal iPropertyTypeEnum As PropertiesForSummaryInformationEnum,
                                                         Optional ByRef newpropertyvalue As String = "",
                                                         Optional ByRef propertyTypeStr As String = "",
                                                         Optional ByVal IsUpdating As Boolean = False) As String
            Dim invSummaryiProperties As PropertySet = DocToUpdate.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
            Dim currentvalue As String = invSummaryiProperties.ItemByPropId(iPropertyTypeEnum).Value
            If IsUpdating Then
                invSummaryiProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
                Return newpropertyvalue
            Else
                Return currentvalue
            End If
        End Function

#End Region

#Region "Get or Set Custom iProperty Values"

        ''' <summary>
        ''' This method should set or get any custom iProperty value
        ''' </summary>
        ''' <param name="Doc">the document to edit</param>
        ''' <param name="PropertyName">the iProperty name to retrieve or update</param>
        ''' <param name="PropertyValue">the optional value to assign - if empty we are retrieving a value</param>
        ''' <returns></returns>
        Friend Shared Function SetorCreateCustomiProperty(ByVal Doc As Inventor.Document, ByVal PropertyName As String, Optional ByVal PropertyValue As Object = Nothing) As Object
            ' Get the custom property set.
            Dim customPropSet As Inventor.PropertySet
            Dim customproperty As Object = Nothing

            customPropSet = Doc.PropertySets.Item("Inventor User Defined Properties")

            ' Get the existing property, if it exists.
            Dim prop As Inventor.Property = Nothing
            Dim propExists As Boolean = True
            Try
                prop = customPropSet.Item(PropertyName)
            Catch ex As Exception
                propExists = False
            End Try
            If Not PropertyValue Is Nothing Then
                ' Check to see if the property was successfully obtained.
                If Not propExists Then
                    ' Failed to get the existing property so create a new one.
                    prop = customPropSet.Add(PropertyValue, PropertyName)
                Else
                    ' Change the value of the existing property.
                    prop.Value = PropertyValue
                End If
            Else
                customproperty = prop.Value
            End If
            Return customproperty
        End Function

#End Region

#End Region

    End Class

End Namespace