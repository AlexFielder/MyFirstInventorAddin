Imports log4net
Imports log4net.Appender
Imports log4net.Repository.Hierarchy
Imports System


Public Class Log4NetFileHelper

        Private DEFAULT_LOG_FILENAME As String = String.Format("application_log_{0}.log", DateTime.Now.ToString("yyyyMMMdd_hhmm"))

        Private root As Logger

        Public Sub New()
        End Sub

        Public Overridable Sub Init()
            root = (CType(LogManager.GetRepository(), Hierarchy)).Root
            root.Repository.Configured = True
        End Sub

        Public Overridable Sub AddConsoleLogging()
            Dim C As ConsoleAppender = GetConsoleAppender()
            AddConsoleLogging(C)
        End Sub

        Public Overridable Sub AddConsoleLogging(ByVal C As ConsoleAppender)
            root.AddAppender(C)
        End Sub

        Public Overridable Function AddFileLogging() As FileAppender
            Return AddFileLogging(DEFAULT_LOG_FILENAME)
        End Function

        Public Overridable Function AddFileLogging(ByVal sFileFullPath As String) As FileAppender
            Return AddFileLogging(sFileFullPath, log4net.Core.Level.All)
        End Function

        Public Overridable Function AddFileLogging(ByVal sFileFullPath As String, ByVal threshold As log4net.Core.Level) As FileAppender
            Return AddFileLogging(sFileFullPath, threshold, True)
        End Function

        Public Overridable Function AddFileLogging(ByVal sFileFullPath As String, ByVal threshold As log4net.Core.Level, ByVal bAppendfile As Boolean) As FileAppender
            Dim appender As FileAppender = GetFileAppender(sFileFullPath, threshold, bAppendfile)
            root.AddAppender(appender)
            Return appender
        End Function

        Public Overridable Function AddRollingFileLogging(ByVal sFileFullPath As String, ByVal threshold As log4net.Core.Level, ByVal bAppendfile As Boolean) As RollingFileAppender
            Dim appender As RollingFileAppender = GetRollingFileAppender(sFileFullPath, threshold, bAppendfile)
            root.AddAppender(appender)
            Return appender
        End Function

        Public Overridable Function AddSMTPLogging(ByVal smtpHost As String, ByVal From As String, ByVal [To] As String, ByVal CC As String, ByVal subject As String, ByVal threshhold As log4net.Core.Level) As SmtpAppender
            Dim appender As SmtpAppender = GetSMTPAppender(smtpHost, From, [To], CC, subject, threshhold)
            root.AddAppender(appender)
            Return appender
        End Function

        Public Function GetLogAppender(ByVal AppenderName As String) As IAppender
            Dim ac As AppenderCollection = (CType(LogManager.GetRepository(), log4net.Repository.Hierarchy.Hierarchy)).Root.Appenders
            For Each appender As log4net.Appender.IAppender In ac
                If appender.Name = AppenderName Then
                    Return appender
                End If
            Next

            Return Nothing
        End Function

        Public Sub CloseAppender(ByVal AppenderName As String)
            Dim appender As log4net.Appender.IAppender = GetLogAppender(AppenderName)
            CloseAppender(appender)
        End Sub

        Private Sub CloseAppender(ByVal appender As log4net.Appender.IAppender)
            appender.Close()
        End Sub

        Private Function GetSMTPAppender(ByVal smtpHost As String, ByVal From As String, ByVal [To] As String, ByVal CC As String, ByVal subject As String, ByVal threshhold As log4net.Core.Level) As SmtpAppender
            Dim lAppender As SmtpAppender = New SmtpAppender()
            lAppender.Cc = CC
            lAppender.[To] = [To]
            lAppender.From = From
            lAppender.SmtpHost = smtpHost
            lAppender.Subject = subject
            lAppender.BufferSize = 512
            lAppender.Lossy = False
            lAppender.Layout = New log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss,fff} %5level [%2thread] %message (%logger{1}:%line)%n")
            lAppender.Threshold = threshhold
            lAppender.ActivateOptions()
            Return lAppender
        End Function

        Private Function GetConsoleAppender() As ConsoleAppender
            Dim lAppender As ConsoleAppender = New ConsoleAppender()
            lAppender.Name = "Console"
            lAppender.Layout = New log4net.Layout.PatternLayout(" %message %n")
            lAppender.Threshold = log4net.Core.Level.All
            lAppender.ActivateOptions()
            Return lAppender
        End Function

        Private Function GetFileAppender(ByVal sFileName As String, ByVal threshhold As log4net.Core.Level, ByVal bFileAppend As Boolean) As FileAppender
            Dim lAppender As FileAppender = New FileAppender()
            lAppender.Name = sFileName
            lAppender.AppendToFile = bFileAppend
            lAppender.File = sFileName
            lAppender.Layout = New log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss} [%logger:%line] %level - %message%newline%exception")
            lAppender.Threshold = threshhold
            lAppender.ActivateOptions()
            Return lAppender
        End Function

        Private Function GetRollingFileAppender(ByVal sFileName As String, ByVal threshhold As log4net.Core.Level, ByVal bFileAppend As Boolean) As RollingFileAppender
            Dim lAppender As RollingFileAppender = New RollingFileAppender()
            lAppender.Name = sFileName
            lAppender.AppendToFile = bFileAppend
            lAppender.File = sFileName
            lAppender.Layout = New log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss} [%logger:%line] %level - %message%newline%exception")
            lAppender.Threshold = threshhold
            lAppender.RollingStyle = RollingFileAppender.RollingMode.Size
            lAppender.MaximumFileSize = "10MB"
            lAppender.MaxSizeRollBackups = 5
            lAppender.StaticLogFileName = True
            lAppender.ActivateOptions()
            Return lAppender
        End Function

        Private Sub ConfigureLog(ByVal sFileName As String)
        End Sub
    End Class


