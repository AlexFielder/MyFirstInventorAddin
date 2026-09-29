Imports Serilog

Namespace iPropertiesController

    ' Feature-usage events for Seq. Each call writes one Information event tagged
    ' EventKind = 'Usage', so usage can be counted in Seq with, for example:
    '     select count(*) from stream where EventKind = 'Usage' group by Feature
    ' Events sent to Seq also carry the session properties added in ConfigureLogging: add-in and
    ' Inventor versions, a per-session id, and the InstallId and LicenseId from the licence.
    Friend Module UsageTracking

        Friend Sub TrackUsage(feature As String, Optional detail As String = Nothing)
            Try
                Dim usageLog = Log.ForContext("EventKind", "Usage")
                If String.IsNullOrEmpty(detail) Then
                    usageLog.Information("Used {Feature} on {DocumentType}", feature, ActiveDocumentType())
                Else
                    usageLog.Information("Used {Feature} ({Detail}) on {DocumentType}", feature, detail, ActiveDocumentType())
                End If
            Catch ex As Exception
                ' Usage tracking must never get in the way of the feature being used.
                Log.Debug(ex, "Usage tracking failed for {Feature}", feature)
            End Try
        End Sub

        Private Function ActiveDocumentType() As String
            Dim doc = AddinGlobal.InventorApp?.ActiveDocument
            If doc Is Nothing Then Return "no document"
            Select Case doc.DocumentType
                Case Inventor.DocumentTypeEnum.kPartDocumentObject
                    Return "part"
                Case Inventor.DocumentTypeEnum.kAssemblyDocumentObject
                    Return "assembly"
                Case Inventor.DocumentTypeEnum.kDrawingDocumentObject
                    Return "drawing"
                Case Inventor.DocumentTypeEnum.kPresentationDocumentObject
                    Return "presentation"
                Case Else
                    Return doc.DocumentType.ToString()
            End Select
        End Function

    End Module

End Namespace
