Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic.FileIO
Imports System.IO

Public Class CSVSQLBulkLoader

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="fileName">CSV file name with extension </param>
    ''' <param name="DestinationTable"> SQL Server table name</param>
    ''' <param name="numberOfColumns">Number of columns in the CSV file. Corresponding DB Table will have one more auto column ID </param>
    Public Sub Load(fileName As String, DestinationTable As String, numberOfColumns As Int32)
        '' fileName = CSV file path and name
        '' DestinationTable = SQL Server table name
        '' numberOfColumns = Number of columns in the CSV file


        Try

            Dim dt As New DataTable()

            Dim line As String = Nothing

            Dim i As Integer = 0

            Dim separator As String = SQLAction.GetDataSeparator
            Dim filePath As String = SQLAction.GetSourceFilePath
            Dim fullFilePath = filePath + "\" + fileName

            Using sr As StreamReader = File.OpenText(fullFilePath)

                line = sr.ReadLine()

                Do While line IsNot Nothing

                    '' UtilLog.LogInfo("Line: " + line)
                    'UtilLog.LogInfo("Separator: " + separator)

                    Dim data() As String = line.Split(separator)

                    If data.Length > 0 Then

                        If i = 0 Then

                            For Each item In data

                                dt.Columns.Add(New DataColumn())

                            Next item

                            i += 1

                        End If

                        Dim row As DataRow = dt.NewRow()

                        row.ItemArray = data

                        dt.Rows.Add(row)
                        'UtilLog.LogInfo("Row value 1: " + dt.Rows().Item(0).Item(0).ToString)
                        'UtilLog.LogInfo("Row value 2: " + dt.Rows().Item(0).Item(1).ToString)
                        'UtilLog.LogInfo("Row value 3: " + dt.Rows().Item(0).Item(2).ToString)
                        'UtilLog.LogInfo("Row value 4: " + dt.Rows().Item(0).Item(3).ToString)
                        'UtilLog.LogInfo("Row value 5: " + dt.Rows().Item(0).Item(4).ToString)
                        'UtilLog.LogInfo("Row value 6: " + dt.Rows().Item(0).Item(5).ToString)
                        'UtilLog.LogInfo("Row value 7: " + dt.Rows().Item(0).Item(6).ToString)
                        'UtilLog.LogInfo("Row value 8: " + dt.Rows().Item(0).Item(7).ToString)

                    End If

                    line = sr.ReadLine()

                Loop

            End Using
            Using cn As New SqlConnection(SQLAction.GetConnectionString())

                cn.Open()

                Using copy As New SqlBulkCopy(cn)

                    Dim colNum As Integer

                    For colNum = 0 To numberOfColumns - 1

                        ''UtilLog.LogInfo("Column Number: " + colNum.ToString)
                        copy.ColumnMappings.Add(colNum, colNum + 1)

                    Next




                    'copy.ColumnMappings.Add(1, 1)

                    'copy.ColumnMappings.Add(2, 2)


                    UtilLog.LogInfo("DestinationTable: " + DestinationTable)
                    copy.DestinationTableName = DestinationTable

                    copy.WriteToServer(dt)

                    BackUpAndDeleteFile(fileName)

                End Using

            End Using

        Catch ex As Exception
            Throw ex
        End Try


    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="fileName"> fileName with extension</param>

    Public Sub BackUpAndDeleteFile(fileName As String)
        'fileName = fileName with extension
        'directoryPath = directory path of the file

        Try
            Dim filePath As String = SQLAction.GetSourceFilePath

            Dim sNameDate = Now.ToString.Replace("/", "-").Replace(":", "-")
            Dim sCheckFile = filePath + "\" + fileName
            Dim sBackupfile = filePath + "\" + "Backup" + "\" + fileName + " " + sNameDate + ".csv"
            Dim sBackupPath = filePath + "\" + "Backup"
            If My.Computer.FileSystem.FileExists(sCheckFile) Then
                If Not My.Computer.FileSystem.DirectoryExists(sBackupPath) Then
                    My.Computer.FileSystem.CreateDirectory(sBackupPath)
                End If
                My.Computer.FileSystem.CopyFile(sCheckFile, sBackupfile, True)
            End If

            If My.Computer.FileSystem.FileExists(sBackupfile) Then
                My.Computer.FileSystem.DeleteFile(sCheckFile)
            End If
        Catch ex As Exception
            Throw ex
        End Try




        'XMLFileNameReal = "00404_" + DatePart1 + "DFI." + SplitName + "_ORIG" + sXmlFileName



    End Sub

End Class