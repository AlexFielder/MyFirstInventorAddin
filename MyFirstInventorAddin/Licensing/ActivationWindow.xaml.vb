Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports Serilog

Namespace iPropertiesController.Licensing

    ' Activation dialog, opened from the "Activate iPropertiesController" ribbon button (never during
    ' Activate: Inventor may still be showing its splash screen). Either confirm an email address with
    ' an emailed code, or enter an admin-issued activation code.
    Partial Public Class ActivationWindow

        Private ReadOnly _licensing As ILicensingService
        Private _challenge As ActivationChallenge

        ' The licensed status once activation succeeds.
        Public Property Result As LicenceStatus

        Public Sub New(licensing As ILicensingService)
            InitializeComponent()
            _licensing = licensing
            If Not String.IsNullOrEmpty(licensing.InstallId) Then InstallIdText.Text = "This PC: " & licensing.InstallId
            AddHandler Loaded, Sub() EmailBox.Focus()
        End Sub

        ' Shows the dialog owned by Inventor's main window. Returns the licensed status, or Nothing
        ' if the user cancelled.
        Public Shared Function ShowActivation(licensing As ILicensingService) As LicenceStatus
            Dim window As New ActivationWindow(licensing)
            Dim helper As New WindowInteropHelper(window) With {.Owner = New IntPtr(AddinGlobal.InventorApp.MainFrameHWND)}
            Return If(window.ShowDialog().GetValueOrDefault(), window.Result, Nothing)
        End Function

        Private Async Sub SendCodeButton_Click(sender As Object, e As RoutedEventArgs) Handles SendCodeButton.Click
            Dim email = EmailBox.Text.Trim()
            If Not LooksLikeEmail(email) Then
                ShowError("Enter your email address.")
                Return
            End If
            Await RunStepAsync(Async Function()
                                   _challenge = Await _licensing.RequestCodeAsync(email)
                                   CodeSentText.Text = $"If {email} is registered for iPropertiesController, a {_challenge.CodeLength}-digit code is on its way. It expires at {_challenge.ExpiresUtc.ToLocalTime():HH:mm}."
                                   CodeEntryPanel.Visibility = Visibility.Visible
                                   SendCodeButton.Content = "Send a new code"
                                   SendCodeButton.IsDefault = False
                                   ActivateWithCodeButton.IsDefault = True
                                   CodeBox.Clear()
                                   CodeBox.Focus()
                               End Function)
        End Sub

        Private Async Sub ActivateWithCodeButton_Click(sender As Object, e As RoutedEventArgs) Handles ActivateWithCodeButton.Click
            Dim code = CodeBox.Text.Replace(" ", "").Trim()
            If _challenge Is Nothing OrElse code.Length = 0 Then
                ShowError("Enter the code from the email.")
                Return
            End If
            Await RunStepAsync(Async Function()
                                   Complete(Await _licensing.ActivateWithCodeAsync(_challenge.ChallengeId, code))
                               End Function)
        End Sub

        Private Async Sub ActivateWithActivationCodeButton_Click(sender As Object, e As RoutedEventArgs) Handles ActivateWithActivationCodeButton.Click
            Dim email = ActivationEmailBox.Text.Trim()
            Dim activationCode = ActivationCodeBox.Text.Trim()
            If Not LooksLikeEmail(email) OrElse activationCode.Length = 0 Then
                ShowError("Enter your email address and the activation code.")
                Return
            End If
            Await RunStepAsync(Async Function()
                                   Complete(Await _licensing.ActivateWithActivationCodeAsync(email, activationCode))
                               End Function)
        End Sub

        Private Sub UseActivationCodeLink_Click(sender As Object, e As RoutedEventArgs) Handles UseActivationCodeLink.Click
            SwitchPath(useActivationCode:=True)
        End Sub

        Private Sub UseEmailCodeLink_Click(sender As Object, e As RoutedEventArgs) Handles UseEmailCodeLink.Click
            SwitchPath(useActivationCode:=False)
        End Sub

        Private Sub SwitchPath(useActivationCode As Boolean)
            HideError()
            EmailCodePanel.Visibility = If(useActivationCode, Visibility.Collapsed, Visibility.Visible)
            ActivationCodePanel.Visibility = If(useActivationCode, Visibility.Visible, Visibility.Collapsed)
            SendCodeButton.IsDefault = Not useActivationCode AndAlso _challenge Is Nothing
            ActivateWithCodeButton.IsDefault = Not useActivationCode AndAlso _challenge IsNot Nothing
            ActivateWithActivationCodeButton.IsDefault = useActivationCode
            If useActivationCode Then
                If ActivationEmailBox.Text.Length = 0 Then ActivationEmailBox.Text = EmailBox.Text
                If ActivationEmailBox.Text.Length = 0 Then ActivationEmailBox.Focus() Else ActivationCodeBox.Focus()
            Else
                EmailBox.Focus()
            End If
        End Sub

        Private Sub Complete(status As LicenceStatus)
            If status Is Nothing OrElse Not status.IsLicensed Then
                ShowError(If(status?.Reason, "Activation didn't complete. Try again."))
                Return
            End If
            Log.Information("Activated licence {LicenseId} for {Customer}", status.LicenseId, status.Customer)
            Result = status
            DialogResult = True
        End Sub

        ' Runs one request with the dialog disabled, and shows failures inside the dialog.
        Private Async Function RunStepAsync(stepAsync As Func(Of Task)) As Task
            HideError()
            InputPanel.IsEnabled = False
            CancelButton.IsEnabled = False
            Mouse.OverrideCursor = Cursors.Wait
            Try
                Await stepAsync()
            Catch ex As LicensingException
                Log.Warning(ex, "Activation step failed with {Code}", ex.Code)
                ShowError(ex.UserMessage())
            Catch ex As Exception
                Log.Error(ex, "Activation step failed")
                ShowError("Something went wrong: " & ex.Message)
            Finally
                Mouse.OverrideCursor = Nothing
                InputPanel.IsEnabled = True
                CancelButton.IsEnabled = True
            End Try
        End Function

        Private Sub ShowError(message As String)
            StatusText.Text = message
            StatusText.Visibility = Visibility.Visible
        End Sub

        Private Sub HideError()
            StatusText.Visibility = Visibility.Collapsed
        End Sub

        Private Shared Function LooksLikeEmail(email As String) As Boolean
            Dim at = email.IndexOf("@"c)
            Return at > 0 AndAlso email.IndexOf("."c, at) > at + 1 AndAlso Not email.EndsWith(".")
        End Function

    End Class

End Namespace
