Imports NLog
Imports NLog.Config
Imports NLog.Targets


Public Class LoggerHelper
    Private Shared logger As Logger = (New LoggerHelper).ConfigureNLog()

    ' Log Info
    Public Shared Sub LogInfo(message As String)
        logger.Info(message)
    End Sub

    ' Log Warning
    Public Shared Sub LogWarning(message As String)
        logger.Warn(message)
    End Sub

    ' Log Error
    Public Shared Sub LogError(message As String, Optional ex As Exception = Nothing)
        If IsNothing(ex) Then
            logger.Error(message)
        Else
            logger.Error(ex, message)
        End If
    End Sub

    Private Function ConfigureNLog() As Logger
        ' Create a new logging configuration
        Dim config As New LoggingConfiguration()


        ' Define variables
        config.Variables("myvar") = "myvalue"
        config.Variables("LogDirectory") = "${specialfolder:folder=LocalApplicationData}/MyClickOnceApp/EJA_QCC_Applogs"


        ' Create file target with rolling archive
        Dim fileTarget As New FileTarget("logfile") With
            {.FileName = "${var:LogDirectory}/SyncLog_${shortdate}.log",
            .Layout = "${longdate} | ${uppercase:${level}} | ${message} ${exception:format=ToString}",
            .ArchiveFileName = "${basedir}/Logs/Archive/SyncLog.log",
            .ArchiveSuffixFormat = "_{#}",
            .ArchiveEvery = FileArchivePeriod.Day,
            .MaxArchiveFiles = 7,
            .KeepFileOpen = False}


        ' Add the rule that uses the target
        config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget)

        ' Apply the configuration
        LogManager.Configuration = config

        ' Get logger instance
        Return LogManager.GetCurrentClassLogger()
    End Function


End Class
