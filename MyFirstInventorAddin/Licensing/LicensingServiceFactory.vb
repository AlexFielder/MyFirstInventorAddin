Imports System.Threading.Tasks

Namespace iPropertiesController.Licensing

    Public Module LicensingServiceFactory

        ' Debug builds only: set to "fake" (starts unlicensed) or "fake-licensed" to exercise the
        ' activation and licence UI without the licence server. Ignored by Release builds.
        Public Const FakeLicensingVariable As String = "AFA_IPROPERTIES_LICENSING"

        Public Function Create() As ILicensingService
#If DEBUG Then
            Dim mode = System.Environment.GetEnvironmentVariable(FakeLicensingVariable)
            If String.Equals(mode, "fake", StringComparison.OrdinalIgnoreCase) Then Return New FakeLicensingService(startLicensed:=False)
            If String.Equals(mode, "fake-licensed", StringComparison.OrdinalIgnoreCase) Then Return New FakeLicensingService(startLicensed:=True)
#End If
            ' The package needs at least one verification key; until the production key is added to
            ' LicensingKeys the add-in fails safe (Activate button only).
            If LicensingKeys.PublicKeys.Count = 0 Then Return New UnavailableLicensingService()
            Return PackageLicensingService.CreateDefault()
        End Function

    End Module

    ' Stand-in until the AFAutomations.Licensing package is referenced: always unlicensed, so a
    ' build without the package fails safe (Activate button only) rather than unlocking.
    Friend NotInheritable Class UnavailableLicensingService
        Implements ILicensingService

        Public ReadOnly Property InstallId As String Implements ILicensingService.InstallId
            Get
                Return Nothing
            End Get
        End Property

        Public Function LoadLocalState() As LicenceStatus Implements ILicensingService.LoadLocalState
            Return LicenceStatus.Unlicensed("Licensing isn't available in this build.")
        End Function

        Public Function RequestCodeAsync(email As String) As Task(Of ActivationChallenge) Implements ILicensingService.RequestCodeAsync
            Return Task.FromException(Of ActivationChallenge)(NotAvailable())
        End Function

        Public Function ActivateWithCodeAsync(challengeId As String, code As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithCodeAsync
            Return Task.FromException(Of LicenceStatus)(NotAvailable())
        End Function

        Public Function ActivateWithActivationCodeAsync(email As String, activationCode As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithActivationCodeAsync
            Return Task.FromException(Of LicenceStatus)(NotAvailable())
        End Function

        Public Function RenewAsync() As Task(Of LicenceStatus) Implements ILicensingService.RenewAsync
            Return Task.FromException(Of LicenceStatus)(NotAvailable())
        End Function

        Public Function DeactivateAsync() As Task Implements ILicensingService.DeactivateAsync
            Return Task.FromException(NotAvailable())
        End Function

        Public Function CheckForUpdateAsync() As Task(Of UpdateOffer) Implements ILicensingService.CheckForUpdateAsync
            Return Task.FromException(Of UpdateOffer)(NotAvailable())
        End Function

        Public Function DownloadUpdateAsync(offer As UpdateOffer) As Task(Of String) Implements ILicensingService.DownloadUpdateAsync
            Return Task.FromException(Of String)(NotAvailable())
        End Function

        Private Shared Function NotAvailable() As LicensingException
            Return New LicensingException(LicensingException.NotAvailable, "Licensing isn't available in this build.")
        End Function

    End Class

End Namespace
