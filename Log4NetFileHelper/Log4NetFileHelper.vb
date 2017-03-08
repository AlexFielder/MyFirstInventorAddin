Imports log4net
Imports log4net.Appender
Imports log4net.Repository.Hierarchy

''' <summary>
''' Copied originally from this Stackoverflow answer: http://stackoverflow.com/a/16514448/572634
''' </summary>
Public Class Log4NetFileHelper
    Private DEFAULT_LOG_FILENAME As String = String.Format("application_log_{0}.log", DateTime.Now.ToString("yyyyMMMdd_hhmm"))
    Private root As Logger

    Public Sub New()
    End Sub

    Public Overridable Sub Init()
        root = DirectCast(LogManager.GetRepository(), Hierarchy).Root
        'root.AddAppender(GetConsoleAppender());
        'root.AddAppender(GetFileAppender(sFileName));
        root.Repository.Configured = True
    End Sub

#Region "Public Helper Methods"
#Region "Console Logging"
    Public Overridable Sub AddConsoleLogging()
        Dim C As ConsoleAppender = GetConsoleAppender()
        AddConsoleLogging(C)
    End Sub

    Public Overridable Sub AddConsoleLogging(C As ConsoleAppender)
        root.AddAppender(C)
    End Sub
#End Region

#Region "File Logging"
    Public Overridable Function AddFileLogging() As FileAppender
        Return AddFileLogging(DEFAULT_LOG_FILENAME)
    End Function

    Public Overridable Function AddFileLogging(sFileFullPath As String) As FileAppender
        Return AddFileLogging(sFileFullPath, log4net.Core.Level.All)
    End Function

    Public Overridable Function AddFileLogging(sFileFullPath As String, threshold As log4net.Core.Level) As FileAppender
        Return AddFileLogging(sFileFullPath, threshold, True)
    End Function

    Public Overridable Function AddFileLogging(sFileFullPath As String, threshold As log4net.Core.Level, bAppendfile As Boolean) As FileAppender
        Dim appender As FileAppender = GetFileAppender(sFileFullPath, threshold, bAppendfile)
        root.AddAppender(appender)
        Return appender
    End Function
    Public Overridable Function AddRollingFileLogging(sFileFullPath As String, threshold As log4net.Core.Level, bAppendfile As Boolean) As RollingFileAppender
        Dim appender As RollingFileAppender = GetRollingFileAppender(sFileFullPath, threshold, bAppendfile)
        root.AddAppender(appender)
        Return appender
    End Function
    Public Overridable Function AddSMTPLogging(smtpHost As String, From As String, [To] As String, CC As String, subject As String, threshhold As log4net.Core.Level) As SmtpAppender
        Dim appender As SmtpAppender = GetSMTPAppender(smtpHost, From, [To], CC, subject, threshhold)
        root.AddAppender(appender)
        Return appender
    End Function

#End Region


    Public Function GetLogAppender(AppenderName As String) As IAppender
        Dim ac As AppenderCollection = DirectCast(LogManager.GetRepository(), log4net.Repository.Hierarchy.Hierarchy).Root.Appenders

        For Each appender As log4net.Appender.IAppender In ac
            If appender.Name = AppenderName Then
                Return appender
            End If
        Next

        Return Nothing
    End Function

    Public Sub CloseAppender(AppenderName As String)
        Dim appender As log4net.Appender.IAppender = GetLogAppender(AppenderName)
        CloseAppender(appender)
    End Sub

    Private Sub CloseAppender(appender As log4net.Appender.IAppender)
        appender.Close()
    End Sub

#End Region

#Region "Private Methods"

    Private Function GetSMTPAppender(smtpHost As String, From As String, [To] As String, CC As String, subject As String, threshhold As log4net.Core.Level) As SmtpAppender
        Dim lAppender As New SmtpAppender()
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
        Dim lAppender As New ConsoleAppender()
        lAppender.Name = "Console"
        lAppender.Layout = New log4net.Layout.PatternLayout(" %message %n")
        lAppender.Threshold = log4net.Core.Level.All
        lAppender.ActivateOptions()
        Return lAppender
    End Function
    ''' <summary>
    ''' DETAILED Logging 
    ''' log4net.Layout.PatternLayout("%date{dd-MM-yyyy HH:mm:ss,fff} %5level [%2thread] %message (%logger{1}:%line)%n");
    '''  
    ''' </summary>
    ''' <param name="sFileName"></param>
    ''' <param name="threshhold"></param>
    ''' <returns></returns>
    Private Function GetFileAppender(sFileName As String, threshhold As log4net.Core.Level, bFileAppend As Boolean) As FileAppender
        Dim lAppender As New FileAppender()
        lAppender.Name = sFileName
        lAppender.AppendToFile = bFileAppend
        lAppender.File = sFileName
        'lAppender.Layout = new log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss,fff} %5level [%2thread] %message (%logger{1}:%line)%n");
        lAppender.Layout = New log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss} [%logger:%line] %level - %message%newline%exception")
        lAppender.Threshold = threshhold
        lAppender.ActivateOptions()
        Return lAppender
    End Function

    ''' <summary>
    ''' Allows us to create at runtime a RollingFileAppender.
    ''' </summary>
    ''' <param name="sFileName"></param>
    ''' <param name="threshhold"></param>
    ''' <param name="bFileAppend"></param>
    ''' <returns></returns>
    Private Function GetRollingFileAppender(sFileName As String, threshhold As log4net.Core.Level, bFileAppend As Boolean) As RollingFileAppender
        Dim lAppender As New RollingFileAppender()
        lAppender.Name = sFileName
        lAppender.AppendToFile = bFileAppend
        lAppender.File = sFileName
        ' "%date{ABSOLUTE} [%logger] %level - %message%newline%exception"
        lAppender.Layout = New log4net.Layout.PatternLayout("%date{yyyy-MM-dd HH:mm:ss} [%logger:%line] %level - %message%newline%exception")
        'lAppender.Layout = new log4net.Layout.PatternLayout("%date{ABSOLUTE} [%logger] %level - %message%newline%exception");
        lAppender.Threshold = threshhold
        lAppender.RollingStyle = RollingFileAppender.RollingMode.Size
        lAppender.MaximumFileSize = "10MB"
        lAppender.MaxSizeRollBackups = 5
        lAppender.StaticLogFileName = True
        lAppender.ActivateOptions()
        Return lAppender
    End Function

    'private FileAppender GetFileAppender(string sFileName)
    '{
    '    return GetFileAppender(sFileName, log4net.Core.Level.All,true);
    '}

#End Region

    Private Sub ConfigureLog(sFileName As String)


    End Sub
End Class

