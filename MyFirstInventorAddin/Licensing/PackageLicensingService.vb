Imports System.IO
Imports System.Net.Http
Imports System.Threading.Tasks
Imports Serilog
Imports Pkg = AFAutomations.Licensing
Imports PkgContracts = AFAutomations.Licensing.Contracts

Namespace iPropertiesController.Licensing

    ' Public keys that verify licences, keyId -> base64 SubjectPublicKeyInfo (ECDSA P-256). Safe to ship:
    ' they can check a licence but not mint one. To rotate, add the new key alongside the old one so
    ' licences signed by either stay valid; with no key listed the add-in stays unlicensed.
    Public Module LicensingKeys
        Public ReadOnly PublicKeys As IReadOnlyDictionary(Of String, String) = New Dictionary(Of String, String) From {
            {"afa-lc-2026", "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAECay13Fy2oLeFVG05cBsTiyNgwZCSgKRIYOfOVJNjGuDG6KXW+2240PHwsYR1LlGtW211MUKI+d8+X9Cz0YQNbg=="}
        }
    End Module

    ' ILicensingService over the shared AFAutomations.Licensing package (AddinLicense), translating its
    ' results and exceptions into the add-in's own types.
    Public NotInheritable Class PackageLicensingService
        Implements ILicensingService

        Private ReadOnly _licence As Pkg.AddinLicense
        ' The last update check, needed to download the offered installer.
        Private _lastUpdate As PkgContracts.UpdateCheckResponse

        Public Sub New(licence As Pkg.AddinLicense)
            _licence = licence
        End Sub

        ' Production wiring: default state folder, lc.afautomations.co.uk, this machine.
        Public Shared Function CreateDefault() As PackageLicensingService
            Dim version = GetType(PackageLicensingService).Assembly.GetName().Version.ToString()
            Return New PackageLicensingService(Pkg.AddinLicense.CreateDefault(Pkg.Products.IPropertiesController, version, LicensingKeys.PublicKeys))
        End Function

        Public ReadOnly Property InstallId As String Implements ILicensingService.InstallId
            Get
                Return _licence.InstallId
            End Get
        End Property

        Public Function LoadLocalState() As LicenceStatus Implements ILicensingService.LoadLocalState
            Return ToStatus(_licence.CheckLocal())
        End Function

        Public Function RequestCodeAsync(email As String) As Task(Of ActivationChallenge) Implements ILicensingService.RequestCodeAsync
            Return CallServerAsync(Async Function()
                                       Dim challenge = Await _licence.StartActivationAsync(email)
                                       Return New ActivationChallenge With {
                                           .ChallengeId = challenge.ChallengeId,
                                           .ExpiresUtc = challenge.ExpiresUtc.UtcDateTime,
                                           .CodeLength = challenge.CodeLength
                                       }
                                   End Function)
        End Function

        Public Function ActivateWithCodeAsync(challengeId As String, code As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithCodeAsync
            Return CallServerAsync(Async Function() ToStatus(Await _licence.CompleteActivationAsync(challengeId, code)))
        End Function

        Public Function ActivateWithActivationCodeAsync(email As String, activationCode As String) As Task(Of LicenceStatus) Implements ILicensingService.ActivateWithActivationCodeAsync
            Return CallServerAsync(Async Function() ToStatus(Await _licence.ActivateWithCodeAsync(email, activationCode)))
        End Function

        Public Async Function RenewAsync() As Task(Of LicenceStatus) Implements ILicensingService.RenewAsync
            Dim outcome = Await _licence.RenewAsync()
            Select Case outcome
                Case Pkg.RenewOutcome.Renewed, Pkg.RenewOutcome.NotDue
                    Return LoadLocalState()
                Case Pkg.RenewOutcome.Revoked
                    ' The package has already deleted licence.json and install.dat.
                    Return LicenceStatus.Unlicensed("This PC's licence was withdrawn by the licence server.", InstallId)
                Case Pkg.RenewOutcome.NotActivated
                    Return LicenceStatus.Unlicensed("Not activated on this PC.", InstallId)
                Case Pkg.RenewOutcome.Offline
                    Throw New LicensingException(LicensingException.ServerUnreachable, "The licence server couldn't be reached; the current lease stands.")
                Case Else
                    Throw New LicensingException("renew_failed", "The licence server refused the renewal; the current lease stands.")
            End Select
        End Function

        Public Function DeactivateAsync() As Task Implements ILicensingService.DeactivateAsync
            Return CallServerAsync(Async Function()
                                       Await _licence.DeactivateAsync()
                                       Return True
                                   End Function)
        End Function

        Public Function CheckForUpdateAsync() As Task(Of UpdateOffer) Implements ILicensingService.CheckForUpdateAsync
            Return CallServerAsync(Async Function()
                                       Dim update = Await _licence.CheckForUpdateAsync("stable")
                                       If update Is Nothing OrElse Not update.UpdateAvailable OrElse update.Download Is Nothing Then Return CType(Nothing, UpdateOffer)
                                       _lastUpdate = update
                                       Return New UpdateOffer With {
                                           .LatestVersion = update.LatestVersion,
                                           .PublishedUtc = If(update.PublishedUtc.HasValue, update.PublishedUtc.Value.UtcDateTime, DateTime.UtcNow),
                                           .ReleaseNotes = update.ReleaseNotes,
                                           .IsRequired = IsBelow(_licence.ProductVersion, update.MinimumSupportedVersion),
                                           .FileName = update.Download.FileName,
                                           .SizeBytes = update.Download.SizeBytes,
                                           .Sha256 = update.Download.Sha256,
                                           .DownloadPath = update.Download.Url
                                       }
                                   End Function)
        End Function

        Public Function DownloadUpdateAsync(offer As UpdateOffer) As Task(Of String) Implements ILicensingService.DownloadUpdateAsync
            Return CallServerAsync(Async Function()
                                       If _lastUpdate Is Nothing OrElse _lastUpdate.LatestVersion <> offer.LatestVersion Then
                                           Throw New LicensingException("update_stale", "Check for updates again before installing.")
                                       End If
                                       Dim folder = Path.Combine(Path.GetTempPath(), "iPropertiesController-update")
                                       Try
                                           Return Await _licence.DownloadUpdateAsync(_lastUpdate, folder)
                                       Catch ex As InvalidDataException
                                           Throw New LicensingException(LicensingException.DownloadCorrupt, ex.Message, ex)
                                       End Try
                                   End Function)
        End Function

        Private Function ToStatus(result As Pkg.LicenseCheckResult) As LicenceStatus
            If Not result.IsValid OrElse result.Payload Is Nothing Then
                Log.Information("Licence check: {Status} - {Message}", result.Status, result.Message)
                Return LicenceStatus.Unlicensed(Describe(result.Status), InstallId)
            End If

            Dim payload = result.Payload
            Dim seq = _licence.Seq
            Dim expires As Date
            Return New LicenceStatus With {
                .IsLicensed = True,
                .LicenseId = payload.LicenseId,
                .InstallId = InstallId,
                .LicenseType = payload.LicenseType,
                .Customer = payload.Customer,
                .Email = payload.Email,
                .ExpiresDate = If(Date.TryParseExact(payload.ExpiresDate, "yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, expires), expires, CType(Nothing, Date?)),
                .LeaseExpiresUtc = If(payload.LeaseExpiresUtc.HasValue, payload.LeaseExpiresUtc.Value.UtcDateTime, CType(Nothing, DateTime?)),
                .RenewDue = _licence.IsRenewalDue(),
                .Seq = If(seq Is Nothing, Nothing, New SeqSettings With {.ServerUrl = seq.ServerUrl, .ApiKey = seq.ApiKey})
            }
        End Function

        Private Shared Function Describe(status As Pkg.LicenseStatus) As String
            Select Case status
                Case Pkg.LicenseStatus.Missing
                    Return "Not activated on this PC."
                Case Pkg.LicenseStatus.LeaseExpired
                    Return "The licence hasn't been renewed for too long (the licence server couldn't be reached). Activate again to continue."
                Case Pkg.LicenseStatus.EntitlementExpired
                    Return "Your organisation's licence has expired."
                Case Pkg.LicenseStatus.WrongInstall
                    Return "The licence file belongs to a different PC."
                Case Else
                    Return "The licence file isn't valid."
            End Select
        End Function

        ' Server refusals keep their problem-details code; network failures become ServerUnreachable.
        Private Shared Async Function CallServerAsync(Of T)(callAsync As Func(Of Task(Of T))) As Task(Of T)
            Try
                Return Await callAsync()
            Catch ex As Pkg.LicenseServerException
                Throw New LicensingException(If(ex.Code, "server_error"), ex.Message, ex)
            Catch ex As HttpRequestException
                Throw New LicensingException(LicensingException.ServerUnreachable, ex.Message, ex)
            Catch ex As TaskCanceledException
                Throw New LicensingException(LicensingException.ServerUnreachable, "The licence server didn't answer in time.", ex)
            End Try
        End Function

        Private Shared Function IsBelow(current As String, minimum As String) As Boolean
            Dim currentVersion As Version = Nothing
            Dim minimumVersion As Version = Nothing
            Return Not String.IsNullOrEmpty(minimum) AndAlso
                   Version.TryParse(current, currentVersion) AndAlso Version.TryParse(minimum, minimumVersion) AndAlso
                   currentVersion < minimumVersion
        End Function

    End Class

End Namespace
