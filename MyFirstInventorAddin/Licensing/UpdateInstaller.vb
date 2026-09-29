Imports System.ComponentModel
Imports System.Threading.Tasks
Imports Serilog

Namespace iPropertiesController.Licensing

    ' Offers, downloads and starts an update. Never silent: the user confirms first, and the installer
    ' itself raises a User Account Control prompt and offers to close Inventor.
    Public Module UpdateInstaller

        ' Returns True when the installer was started.
        Public Async Function DownloadAndRunAsync(licensing As ILicensingService, offer As UpdateOffer, Optional reportProgress As Action(Of String) = Nothing) As Task(Of Boolean)
            If ShowMessage($"Download and install iPropertiesController {offer.LatestVersion} now?" & vbLf & vbLf &
                           "The installer needs administrator rights, and Inventor has to close before it can finish - the installer will offer to close it.",
                           MsgBoxStyle.YesNo, "Update iPropertiesController") <> MsgBoxResult.Yes Then Return False

            reportProgress?.Invoke($"Downloading {offer.FileName}...")
            ' The service verifies the download's SHA-256 before returning its path.
            Dim installerPath = Await licensing.DownloadUpdateAsync(offer)

            Try
                Process.Start(New ProcessStartInfo(installerPath) With {.UseShellExecute = True})
            Catch ex As Win32Exception When ex.NativeErrorCode = 1223 ' ERROR_CANCELLED: the UAC prompt was declined
                reportProgress?.Invoke("The update was cancelled.")
                Return False
            End Try
            TrackUsage("Install update", offer.LatestVersion)
            Log.Information("Started the installer for version {Version}", offer.LatestVersion)
            Return True
        End Function

    End Module

End Namespace
