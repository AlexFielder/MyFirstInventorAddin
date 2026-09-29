Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports Serilog

Namespace iPropertiesController.Licensing

    ' Licence details, opened from the "iPropertiesController licence" ribbon button: check for
    ' updates, install one, or deactivate this PC to free its seat.
    Partial Public Class LicenceWindow

        Private ReadOnly _licensing As ILicensingService
        Private _offer As UpdateOffer

        ' True once the user has deactivated this PC.
        Public Property DeactivatedThisPc As Boolean

        Public Sub New(licensing As ILicensingService, status As LicenceStatus, Optional pendingOffer As UpdateOffer = Nothing)
            InitializeComponent()
            _licensing = licensing
            EmailText.Text = status.Email
            CustomerText.Text = status.Customer
            LicenceTypeText.Text = status.LicenseType
            ExpiresText.Text = If(status.ExpiresDate.HasValue, status.ExpiresDate.Value.ToString("yyyy-MM-dd"), "-")
            LeaseText.Text = If(status.LeaseExpiresUtc.HasValue,
                                status.LeaseExpiresUtc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm") & " (renews automatically)",
                                "Offline licence - no renewal needed")
            VersionText.Text = ProductInfo.ReleaseVersion()
            InstallIdText.Text = If(status.InstallId, licensing.InstallId)
            LicenseIdText.Text = status.LicenseId
            If pendingOffer IsNot Nothing Then ShowOffer(pendingOffer)
        End Sub

        ' Shows the window owned by Inventor's main window. Returns True if the user deactivated this PC.
        Public Shared Function ShowLicence(licensing As ILicensingService, status As LicenceStatus, Optional pendingOffer As UpdateOffer = Nothing) As Boolean
            Dim window As New LicenceWindow(licensing, status, pendingOffer)
            Dim helper As New WindowInteropHelper(window) With {.Owner = New IntPtr(AddinGlobal.InventorApp.MainFrameHWND)}
            window.ShowDialog()
            Return window.DeactivatedThisPc
        End Function

        Private Async Sub CheckUpdatesButton_Click(sender As Object, e As RoutedEventArgs) Handles CheckUpdatesButton.Click
            Await RunStepAsync(Async Function()
                                   Dim offer = Await _licensing.CheckForUpdateAsync()
                                   If offer Is Nothing Then
                                       UpdateText.Text = $"You have the latest version ({ProductInfo.ReleaseVersion()})."
                                       UpdateText.Visibility = Visibility.Visible
                                   Else
                                       ShowOffer(offer)
                                   End If
                               End Function)
        End Sub

        Private Async Sub InstallUpdateButton_Click(sender As Object, e As RoutedEventArgs) Handles InstallUpdateButton.Click
            Await RunStepAsync(Async Function()
                                   If Await UpdateInstaller.DownloadAndRunAsync(_licensing, _offer, Sub(message) UpdateText.Text = message) Then Close()
                               End Function)
        End Sub

        Private Async Sub DeactivateButton_Click(sender As Object, e As RoutedEventArgs) Handles DeactivateButton.Click
            If ShowMessage("Deactivate iPropertiesController on this PC?" & vbLf & vbLf &
                           "This frees the licence so it can be used on another PC. iPropertiesController stops working here the next time Inventor starts.",
                           MsgBoxStyle.YesNo, "Deactivate this PC") <> MsgBoxResult.Yes Then Return
            Await RunStepAsync(Async Function()
                                   Await _licensing.DeactivateAsync()
                                   TrackUsage("Deactivate this PC")
                                   DeactivatedThisPc = True
                                   ShowMessage("This PC has been deactivated. iPropertiesController will stop working the next time Inventor starts.", MsgBoxStyle.OkOnly, "Deactivate this PC")
                                   Close()
                               End Function)
        End Sub

        Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs) Handles CloseButton.Click
            Close()
        End Sub

        Private Sub ShowOffer(offer As UpdateOffer)
            _offer = offer
            UpdateText.Text = If(offer.IsRequired,
                                 $"Version {offer.LatestVersion} is available, and this version is no longer supported - please update.",
                                 $"Version {offer.LatestVersion} is available (released {offer.PublishedUtc.ToLocalTime():yyyy-MM-dd}).")
            UpdateText.Visibility = Visibility.Visible
            If Not String.IsNullOrWhiteSpace(offer.ReleaseNotes) Then
                ReleaseNotesBox.Text = offer.ReleaseNotes
                ReleaseNotesBox.Visibility = Visibility.Visible
            End If
            InstallUpdateButton.Visibility = Visibility.Visible
        End Sub

        Private Async Function RunStepAsync(stepAsync As Func(Of Task)) As Task
            StatusText.Visibility = Visibility.Collapsed
            IsEnabled = False
            Mouse.OverrideCursor = Cursors.Wait
            Try
                Await stepAsync()
            Catch ex As LicensingException
                Log.Warning(ex, "Licence window action failed with {Code}", ex.Code)
                StatusText.Text = ex.UserMessage()
                StatusText.Visibility = Visibility.Visible
            Catch ex As Exception
                Log.Error(ex, "Licence window action failed")
                StatusText.Text = "Something went wrong: " & ex.Message
                StatusText.Visibility = Visibility.Visible
            Finally
                Mouse.OverrideCursor = Nothing
                IsEnabled = True
            End Try
        End Function

    End Class

End Namespace
