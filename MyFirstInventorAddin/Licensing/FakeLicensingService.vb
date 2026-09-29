#If DEBUG Then
Imports System.Threading.Tasks

Namespace iPropertiesController.Licensing

    ' Debug-only stand-in for the licence server, to exercise the activation and licence UI before the
    ' server exists (see LicensingServiceFactory.FakeLicensingVariable). State lives in memory only.
    '   Emailed code: any email; the code is 000000. Any other code fails with invalid_code.
    '   Activation code: AFA-TEST-TEST-TEST. An email ending @noseats.test fails with no_seats.
    '   AFA_IPROPERTIES_FAKE_UPDATE=1 makes the update check offer version 2026.0.999.
    Friend NotInheritable Class FakeLicensingService
        Implements ILicensingService

        Private Const FakeInstallId As String = "IPC-FAKE0-00000-00000-00000"
        Private _status As LicenceStatus

        Public Sub New(startLicensed As Boolean)
            _status = If(startLicensed, Licensed("developer@afautomations.test"), LicenceStatus.Unlicensed("Not activated (fake licensing).", FakeInstallId))
        End Sub

        Public ReadOnly Property InstallId As String Implements ILicensingService.InstallId
            Get
                Return FakeInstallId
            End Get
        End Property

        Public Function LoadLocalState() As LicenceStatus Implements ILicensingService.LoadLocalState
            Return _status
        End Function

        Public Async Function RequestCodeAsync(email As String) As Task(Of ActivationChallenge) Implements ILicensingService.RequestCodeAsync
            Await Task.Delay(300)
            Return New ActivationChallenge With {.ChallengeId = "chl_fake", .ExpiresUtc = DateTime.UtcNow.AddMinutes(10), .CodeLength = 6}
        End Function

        Public Async Function ActivateWithCodeAsync(challengeId As String, code As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithCodeAsync
            Await Task.Delay(300)
            If code <> "000000" Then Throw New LicensingException("invalid_code", "Invalid code.")
            _status = Licensed("developer@afautomations.test")
            Return _status
        End Function

        Public Async Function ActivateWithActivationCodeAsync(email As String, activationCode As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithActivationCodeAsync
            Await Task.Delay(300)
            If email.EndsWith("@noseats.test", StringComparison.OrdinalIgnoreCase) Then Throw New LicensingException("no_seats", "No seats.")
            If Not String.Equals(activationCode, "AFA-TEST-TEST-TEST", StringComparison.OrdinalIgnoreCase) Then Throw New LicensingException("invalid_code", "Invalid code.")
            _status = Licensed(email)
            Return _status
        End Function

        Public Async Function RenewAsync() As Task(Of LicenceStatus) Implements ILicensingService.RenewAsync
            Await Task.Delay(300)
            If _status.IsLicensed Then _status = Licensed(_status.Email)
            Return _status
        End Function

        Public Async Function DeactivateAsync() As Task Implements ILicensingService.DeactivateAsync
            Await Task.Delay(300)
            _status = LicenceStatus.Unlicensed("This PC was deactivated.", FakeInstallId)
        End Function

        Public Async Function CheckForUpdateAsync() As Task(Of UpdateOffer) Implements ILicensingService.CheckForUpdateAsync
            Await Task.Delay(300)
            If System.Environment.GetEnvironmentVariable("AFA_IPROPERTIES_FAKE_UPDATE") <> "1" Then Return Nothing
            Return New UpdateOffer With {
                .LatestVersion = "2026.0.999",
                .PublishedUtc = DateTime.UtcNow.AddDays(-1),
                .ReleaseNotes = "Fake release for testing the update offer.",
                .FileName = "iPropertiesController-Setup-2026.0.999.exe",
                .SizeBytes = 3_300_000
            }
        End Function

        Public Function DownloadUpdateAsync(offer As UpdateOffer) As Task(Of String) Implements ILicensingService.DownloadUpdateAsync
            Return Task.FromException(Of String)(New LicensingException(LicensingException.NotAvailable, "The fake licence server has no installer to download."))
        End Function

        Private Shared Function Licensed(email As String) As LicenceStatus
            Return New LicenceStatus With {
                .IsLicensed = True,
                .LicenseId = "LIC-FAKE0000",
                .InstallId = FakeInstallId,
                .LicenseType = "Internal",
                .Customer = "AF Automations (fake)",
                .Email = email,
                .ExpiresDate = Date.Today.AddYears(1),
                .LeaseExpiresUtc = DateTime.UtcNow.AddDays(30)
            }
        End Function

    End Class

End Namespace
#End If
