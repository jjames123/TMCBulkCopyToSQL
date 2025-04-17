Imports System
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Xml



Public Class SQLAction
    Public myReader As SqlDataReader
    '' Private ef As New EfassDataService()

    Dim sServerName, sPort, sUserName, sPassword, sDatabaseName As String
    'Public ConnEquinox As OdbcConnection
    Public gbConnEquinox As SqlConnection
    'Public CmdEquinox As OdbcCommand
    Public gbCmdEquinox As SqlCommand
    Public gbDeleteCmdEquinox As SqlCommand
    'Public adapter As OdbcDataAdapter
    Public gbAdapter As SqlDataAdapter
    'Public myReader As OdbcDataReader

    Public Function INTConnectToEquinox() As Boolean
        Try


            Dim gbConnEquinox As New SqlConnection() ''"Data Source='" & sServerName & "';Port='" & sPort & "';UID='" & sUserName & "';PWD='" & sPassword & "';Database='" & sDatabaseName & "';")
            Dim ConnectionStringName As String

            '' ConnectionStringName = Me.DataWorkspace.EfassData.Details.Name.EfassData.Details.Name
            ''  gbConnEquinox.ConnectionString = ConfigurationManager.ConnectionStrings("b90c97af-d758-4e46-befb-2dd9c9bce687").ConnectionString
            gbConnEquinox.ConnectionString = SQLAction.GetConnectionString("Sql") ''"Data Source=DICL-PC\SQLEXPRESS;Initial Catalog=Efass;User ID=sa;Password=secret123" ''ConfigurationManager.ConnectionStrings("b90c97af-d758-4e46-befb-2dd9c9bce687").ConnectionString
            gbConnEquinox.Open()
        Catch ex As Exception
            'MsgBox("efassConfig NOT FOUND", MsgBoxStyle.Critical)
            MsgBox(ex.Message, MsgBoxStyle.Information)
            Return False
        End Try
        Return True
    End Function
    Public Sub New()

        INTConnectToEquinox()

    End Sub

    Public Shared Function GetConnectionString() As String

        Dim connectionString As String = System.Configuration.ConfigurationManager.ConnectionStrings("egSQLConString").ConnectionString
        Return connectionString


    End Function
    Public Shared Function GetDataSeparator() As String

        Dim dataSeparator As String = System.Configuration.ConfigurationManager.ConnectionStrings("dataSeparator").ConnectionString
        Return dataSeparator


    End Function
    Public Shared Function GetSourceFilePath() As String

        Dim filePath As String = System.Configuration.ConfigurationManager.ConnectionStrings("FileSourcePath").ConnectionString
        Return filePath


    End Function
    Public Shared Function GetLogPath() As String

        Dim logPath As String = System.Configuration.ConfigurationManager.ConnectionStrings("LogPath").ConnectionString
        Return logPath


    End Function


End Class
