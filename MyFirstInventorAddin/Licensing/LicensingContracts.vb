Imports System.Threading.Tasks

Namespace iPropertiesController.Licensing

    ' The add-in's view of licensing (lc.afautomations.co.uk, contract v1 rev 2 in
    ' AF-Automations/AFAutomations.LicenseServer docs/design.md). Everything behind this interface -
    ' licence verification, the InstallId fingerprint, the protected state files and the HTTP calls -
    ' comes from the shared AFAutomations.Licensing package; the add-in only orchestrates it.
    Public Interface ILicensingService

        ' "IPC-XXXXX-XXXXX-XXXXX-XXXXX", available before activation (shown for support).
        ReadOnly Property InstallId As String

        ' Reads and verifies licence.json and install.dat. No network access, so it is safe to call
        ' at the very start of Activate.
        Function LoadLocalState() As LicenceStatus

        ' Activation by emailed code: step 1 emails a code (same reply whether or not the address is
        ' known), step 2 exchanges it for a licence.
        Function RequestCodeAsync(email As String) As Task(Of ActivationChallenge)
        Function ActivateWithCodeAsync(challengeId As String, code As String) As Task(Of LicenceStatus)

        ' Activation by an admin-issued AFA-XXXX-XXXX-XXXX code.
        Function ActivateWithActivationCodeAsync(email As String, activationCode As String) As Task(Of LicenceStatus)

        ' New lease and current Seq settings. A "revoked" reply deletes the local state and returns
        ' an unlicensed status.
        Function RenewAsync() As Task(Of LicenceStatus)

        ' Frees this PC's seat and deletes the local state.
        Function DeactivateAsync() As Task

        ' Nothing when the installed version is current.
        Function CheckForUpdateAsync() As Task(Of UpdateOffer)

        ' Downloads the installer to a temporary file, verifies its SHA-256 and returns the path.
        Function DownloadUpdateAsync(offer As UpdateOffer) As Task(Of String)

    End Interface

    Public NotInheritable Class LicenceStatus
        Public Property IsLicensed As Boolean
        ' Why the add-in is unlicensed, in words a user can act on.
        Public Property Reason As String
        Public Property LicenseId As String
        Public Property InstallId As String
        Public Property LicenseType As String
        Public Property Customer As String
        Public Property Email As String
        Public Property ExpiresDate As Date?
        ' Nothing for an offline (air-gapped) licence.
        Public Property LeaseExpiresUtc As DateTime?
        Public Property RenewDue As Boolean
        ' Nothing until activation has delivered a Seq key.
        Public Property Seq As SeqSettings

        Public Shared Function Unlicensed(reason As String, Optional installId As String = Nothing) As LicenceStatus
            Return New LicenceStatus With {.IsLicensed = False, .Reason = reason, .InstallId = installId}
        End Function
    End Class

    Public NotInheritable Class SeqSettings
        Public Property ServerUrl As String
        Public Property ApiKey As String
    End Class

    Public NotInheritable Class ActivationChallenge
        Public Property ChallengeId As String
        Public Property ExpiresUtc As DateTime
        Public Property CodeLength As Integer
    End Class

    Public NotInheritable Class UpdateOffer
        Public Property LatestVersion As String
        Public Property PublishedUtc As DateTime
        ' Markdown from the GitHub release body.
        Public Property ReleaseNotes As String
        ' The installed version is below the minimum the server still supports.
        Public Property IsRequired As Boolean
        Public Property FileName As String
        Public Property SizeBytes As Long
        Public Property Sha256 As String
        ' Server-relative download path, for the service implementation.
        Public Property DownloadPath As String
    End Class

    ' A failure the user can act on. Code is the server's problem-details code (email_not_verified,
    ' no_seats, revoked, invalid_code, rate_limited, unknown_product, entitlement_expired) or one of
    ' the client-side codes below.
    Public Class LicensingException
        Inherits Exception

        Public Const ServerUnreachable As String = "server_unreachable"
        Public Const NotAvailable As String = "not_available"
        Public Const DownloadCorrupt As String = "download_corrupt"

        Public ReadOnly Property Code As String

        Public Sub New(code As String, message As String, Optional inner As Exception = Nothing)
            MyBase.New(message, inner)
            Me.Code = code
        End Sub

        ' The message to show for a code, falling back to the server's own detail.
        Public Function UserMessage() As String
            Select Case Code
                Case "invalid_code"
                    Return "That code isn't right, or it has expired. Check it, or ask for a new one."
                Case "email_not_verified"
                    Return "That email address hasn't been confirmed yet. Ask for a code first."
                Case "no_seats"
                    Return "All of your organisation's licences are in use. Deactivate iPropertiesController on a PC you no longer use, or contact AF Automations."
                Case "revoked"
                    Return "This PC's licence has been withdrawn. Contact AF Automations if you think that's wrong."
                Case "rate_limited"
                    Return "Too many attempts. Wait a few minutes and try again."
                Case "entitlement_expired"
                    Return "Your organisation's iPropertiesController licence has expired. Contact AF Automations to renew it."
                Case "unknown_product"
                    Return "The licence server doesn't recognise this product. Contact AF Automations."
                Case ServerUnreachable
                    Return "Couldn't reach the licence server. Check your internet connection and try again."
                Case NotAvailable
                    Return "Licensing isn't available in this build of iPropertiesController."
                Case DownloadCorrupt
                    Return "The update download was damaged (its checksum didn't match), so it wasn't run. Try again later."
                Case Else
                    Return If(String.IsNullOrWhiteSpace(Message), "The licence server returned an error.", Message)
            End Select
        End Function
    End Class

    Public Module ProductInfo

        Public Const ProductId As String = "iPropertiesController"

        ' The licence server compares only the first three version parts: the T4 AssemblyInfo scheme
        ' is 2026.0.<release>.<build>, and the build part increments on every build on any machine.
        Public Function ReleaseVersion() As String
            Dim v = GetType(ProductInfo).Assembly.GetName().Version
            Return $"{v.Major}.{v.Minor}.{v.Build}"
        End Function

    End Module

End Namespace
