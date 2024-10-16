Imports System.IO
Imports System.Net.Mail
Imports System.Text
Imports OfficeOpenXml ' EPPlus library for Excel processing
Imports System.Net
Imports System.Data.SqlClient

Public Class OnsemiDump
    Dim EmailSubject As String
    Dim em As New EmailHandler
    Dim em2 As New EmailHandler
    'Dim dsEmail As DataSet = em.GetMailRecipients(141)
    'Dim dsEmail2 As DataSet = em2.GetMailRecipients(141)
    Dim dsEmail As DataSet = em.GetMailRecipients(19)
    Dim dsEmail2 As DataSet = em2.GetMailRecipients(85)

    Private Sub OnsemiDump_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProcessLatestExcelFile("\\192.168.5.9\c$\WebMESReports_IXYS", "C:\WebMESReports_ON", "WaferAllocation_", AddressOf SendEmailWithExcelData)
        Me.Close()
    End Sub

    Public Shared Function CheckDumpedDevices() As Boolean
        Dim sqlhander As New SQLHandler
        Dim sql As String = "[usp_CST_ONSEMI_Dumped_Device_Price]"
        Dim dr As SqlDataReader = Nothing

        If sqlhander.OpenConnection() Then
            If sqlhander.FillDataReader(sql, dr, CommandType.StoredProcedure) Then
                If dr.HasRows Then
                    Return True
                End If
            End If
            sqlhander.CloseConnection()
        End If

        sqlhander = Nothing
        Return False
    End Function

    Private Sub ProcessLatestExcelFile(sourceDir As String, destDir As String, filePrefix As String, sendEmailMethod As Action(Of List(Of String()), String))
        ' Set the LicenseContext property for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        If Not Directory.Exists(destDir) Then
            Directory.CreateDirectory(destDir)
        End If

        Dim latestFile As FileInfo = GetLatestWaferAllocationFile(sourceDir, filePrefix)

        If latestFile IsNot Nothing Then
            Dim destinationPath As String = Path.Combine(destDir, latestFile.Name)
            latestFile.MoveTo(destinationPath)

            Dim excelData As List(Of String()) = ReadDataFromExcel(destinationPath)
            sendEmailMethod(excelData, destinationPath)
        End If
    End Sub

    Private Sub NoDevice()
        ProcessLatestExcelFile("D:\waferAllocation", "C:\WebMESReports_ON", "ONSemi_PriceAlert_", AddressOf SendEmailWithExcelDataNoDevice)
    End Sub

    Public Sub FetchReport()
        Dim tempFolder_Storage As String = "D:\waferAllocation"
        Dim ReportServerUrl As String = "http://192.168.5.12/ReportServer?/OnSemi/ONSemi+Price+Alert&rs:Format=EXCELOPENXML"
        Dim Username As String = "Administrator"
        Dim Password As String = "Dnhk$%07232022"
        Dim Domain As String = "atecphil3.com"
        Dim FileName As String = $"ONSemi_PriceAlert_{Date.Now:MMddyyyy_HHmmss}.xlsx"
        Dim FilePath As String = Path.Combine(tempFolder_Storage, FileName)

        Try
            Using webClient As New WebClient()
                webClient.Credentials = New NetworkCredential(Username, Password, Domain)
                webClient.DownloadFile(ReportServerUrl, FilePath)
            End Using
        Catch ex As Exception
            ' Handle exceptions here (logging or notifying)
        End Try
    End Sub

    Private Function GetLatestWaferAllocationFile(sourceDir As String, filePrefix As String) As FileInfo
        Dim dirInfo As New DirectoryInfo(sourceDir)
        Dim files As FileInfo() = dirInfo.GetFiles($"{filePrefix}*.xlsx").OrderByDescending(Function(f) f.LastWriteTime).ToArray()

        Return If(files.Any(), files(0), Nothing)
    End Function

    Private Function ReadDataFromExcel(filePath As String) As List(Of String())
        Dim data As New List(Of String())

        If File.Exists(filePath) Then
            Using package As New ExcelPackage(New FileInfo(filePath))
                Dim worksheet As ExcelWorksheet = package.Workbook?.Worksheets.FirstOrDefault()
                If worksheet IsNot Nothing Then
                    Dim rowCount As Integer = worksheet.Dimension.Rows
                    Dim colCount As Integer = Math.Min(worksheet.Dimension.Columns, 11)

                    For row As Integer = 1 To rowCount
                        Dim rowData(colCount - 1) As String
                        For col As Integer = 1 To colCount
                            rowData(col - 1) = If(worksheet.Cells(row, col).Value?.ToString(), "")
                        Next
                        data.Add(rowData)
                    Next
                Else
                    MessageBox.Show("Worksheet not found in Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        Else
            MessageBox.Show("Excel file not found at specified path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Return data
    End Function

    Private Sub SendEmailWithExcelData(data As List(Of String()), filePath As String)
        EmailSubject = "Wafer Allocation Details " & DateTime.Now.ToString()

        Dim message As New StringBuilder()
        message.AppendLine("<html><body><table border='1' cellpadding='5' cellspacing='0'><tr>")
        For Each columnName In {"Customer", "Device", "Die Part", "ProductID", "Package", "Lead Count", "Wafer Lot", "PID", "Quantity", "Start Date", "Load Qty"}
            message.AppendLine($"<th style='background-color: #007bff; color: #ffffff; font-weight: bold; font-size: 14px; font-family: Times New Roman, Times, serif;'>{columnName}</th>")
        Next
        message.AppendLine("</tr>")

        For rowIdx As Integer = 1 To data.Count - 1
            message.AppendLine("<tr>")
            For Each cellValue In data(rowIdx)
                message.AppendLine($"<td style='font-family: Times New Roman, Times, serif; font-size: 11pt;'>{cellValue}</td>")
            Next
            message.AppendLine("</tr>")
        Next

        message.AppendLine("</table></body></html><p><strong>Please see attached allocated PID for today.</strong></p>")
        em.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
        If CheckDumpedDevices() Then
            FetchReport()
            NoDevice()
        End If

    End Sub

    Private Sub SendEmailWithExcelDataNoDevice(data As List(Of String()), filePath As String)
        EmailSubject = "onsemi price alert " & DateTime.Now.ToString()

        Dim message As New StringBuilder()
        message.AppendLine("<html><body><table border='1' cellpadding='5' cellspacing='0'><tr>")
        For Each columnName In {"1", "2", "3", "4", "5", "6"}
            message.AppendLine($"<th style='background-color: #007bff; color: #ffffff; font-weight: bold; font-size: 14px; font-family: Times New Roman, Times, serif;'>{columnName}</th>")
        Next
        message.AppendLine("</tr>")

        For rowIdx As Integer = 1 To data.Count - 1
            message.AppendLine("<tr>")
            For Each cellValue In data(rowIdx)
                message.AppendLine($"<td style='font-family: Times New Roman, Times, serif; font-size: 11pt;'>{cellValue}</td>")
            Next
            message.AppendLine("</tr>")
        Next

        message.AppendLine("</table></body></html><p><strong>Please see attached file for Price Matrix of ONSemi Dumped Devices as of " & DateTime.Now & "</strong></p>")
        em2.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
    End Sub
End Class



'Imports System.IO
'Imports System.Net.Mail
'Imports System.Text
'Imports OfficeOpenXml ' EPPlus library for Excel processing
'Imports System.Net
'Imports System.Data.SqlClient

'Public Class OnsemiDump
'    Dim EmailSubject As String
'    Dim em As New EmailHandler
'    Dim dsEmail As DataSet = em.GetMailRecipients(141)
'    'Dim dsEmail As DataSet = em.GetMailRecipients(19)

'    Private Sub OnsemiDump_Load(sender As Object, e As EventArgs) Handles MyBase.Load

'        ProcessLatestExcelFile()
'        If CheckDumpedDevices() Then
'            FetchReport()
'            NoDevice()
'        Else
'            'MessageBox.Show("No dumped devices found, report will not be generated.")
'        End If
'        Me.Close()
'    End Sub

'    Public Shared Function CheckDumpedDevices() As Boolean
'        Dim sqlhander As New SQLHandler
'        Dim sql As String = "[usp_CST_ONSEMI_Dumped_Device_Price]"
'        Dim dr As SqlDataReader = Nothing

'        If sqlhander.OpenConnection() Then
'            If sqlhander.FillDataReader(sql, dr, CommandType.StoredProcedure) Then
'                If dr.HasRows Then
'                    ' Data found, proceed with generating report
'                    Return True
'                End If
'            End If
'            sqlhander.CloseConnection()
'        End If

'        sqlhander = Nothing
'        Return False
'    End Function



'    'Public Shared Function CheckDumpedDevices() As Boolean
'    '    Dim SqlHander As New SQLHandler
'    '    Dim sql As String = "[]"
'    '    Dim dr As SqlDataReader = Nothing

'    '    If SqlHander.OpenConnection() Then
'    '        If SqlHander.FillDataReader(sql, dr, CommandType.StoredProcedure) Then
'    '            If dr.HasRows Then
'    '                Return True
'    '            End If
'    '        End If
'    '    End If
'    'End Function




'    Private Sub ProcessLatestExcelFile()
'        ' Set the LicenseContext property for EPPlus
'        ExcelPackage.LicenseContext = LicenseContext.NonCommercial ' Set to LicenseContext.Commercial if you have a commercial license

'        ' Define the source and destination directories
'        Dim sourceDir As String = "\\192.168.5.9\c$\WebMESReports_IXYS"
'        Dim destDir As String = "C:\WebMESReports_ON"
'        'Dim paths As DataSet = GetPathsFromDB()
'        'Dim sourceDir As String = paths.Tables(0).Rows(0)("SourcePath").ToString()
'        'Dim destDir As String = paths.Tables(0).Rows(0)("DestinationPath").ToString()


'        ' Ensure destination directory exists
'        If Not Directory.Exists(destDir) Then
'            Directory.CreateDirectory(destDir)
'        End If

'        ' Find the latest WaferAllocation file
'        Dim latestFile As FileInfo = GetLatestWaferAllocationFile(sourceDir)

'        If latestFile IsNot Nothing Then
'            Dim destinationPath As String = Path.Combine(destDir, latestFile.Name)
'            latestFile.MoveTo(destinationPath)

'            Dim excelData As List(Of String()) = ReadDataFromExcel(destinationPath)

'            ' Process and send email with Excel data
'            SendEmailWithExcelData(excelData, destinationPath)
'        Else
'            'MessageBox.Show("No WaferAllocation files found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'    End Sub

'    Private Sub NoDevice()
'        ' Set the LicenseContext property for EPPlus
'        ExcelPackage.LicenseContext = LicenseContext.NonCommercial ' Set to LicenseContext.Commercial if you have a commercial license

'        ' Define the source and destination directories
'        Dim sourceDir As String = "D:\waferAllocation"
'        Dim destDir As String = "C:\WebMESReports_ON"
'        'Dim paths As DataSet = GetPathsFromDB()
'        'Dim sourceDir As String = paths.Tables(0).Rows(0)("SourcePath").ToString()
'        'Dim destDir As String = paths.Tables(0).Rows(0)("DestinationPath").ToString()


'        ' Ensure destination directory exists
'        If Not Directory.Exists(destDir) Then
'            Directory.CreateDirectory(destDir)
'        End If

'        ' Find the latest WaferAllocation file
'        Dim latestFile As FileInfo = GetLatestWaferAllocationFile2(sourceDir)

'        If latestFile IsNot Nothing Then
'            Dim destinationPath As String = Path.Combine(destDir, latestFile.Name)
'            latestFile.MoveTo(destinationPath)

'            Dim excelData As List(Of String()) = ReadDataFromExcel2(destinationPath)

'            ' Process and send email with Excel data
'            SendEmailWithExcelDataNoDevice(excelData, destinationPath)
'        Else
'            'MessageBox.Show("No WaferAllocation files found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'    End Sub

'    Public Sub FetchReport()
'        Dim tempFolder_Storage As String = "D:\waferAllocation" ' Replace with the actual folder path

'        ' URL to access the report directly in Excel format
'        Dim ReportServerUrl As String = "http://192.168.5.12/ReportServer?/OnSemi/ONSemi+Price+Alert&rs:Format=EXCELOPENXML"
'        Dim Username As String = "Administrator"
'        Dim Password As String = "Dnhk$%07232022"
'        Dim Domain As String = "atecphil3.com"
'        Dim FileName As String = "ONSemi_PriceAlert_" & Date.Now().ToString("MMddyyyy_HHmmss") & ".xlsx"
'        Dim FilePath As String = Path.Combine(tempFolder_Storage, FileName)

'        Try
'            Using webClient As New WebClient()
'                webClient.Credentials = New NetworkCredential(Username, Password, Domain)
'                webClient.DownloadFile(ReportServerUrl, FilePath)
'                ' Console.WriteLine("Report downloaded successfully at: " & FilePath)
'            End Using

'        Catch ex As Exception
'            'Console.WriteLine("Error downloading report: " & ex.Message)
'        End Try
'    End Sub

'    'Private Function GetPathsFromDB() As DataSet
'    '    Dim dsPaths As New DataSet()
'    '    Dim strSQL As String = "usp_GetFilePaths" ' Replace with your actual stored procedure name
'    '    Dim sql_handler As New SQLHandler()

'    '    ' Assuming AutoEmailCode is a required parameter, adjust accordingly
'    '    sql_handler.CreateParameter(1)
'    '    sql_handler.SetParameterValues(0, "@AutoEmailCode", SqlDbType.BigInt, autoEmailCode) ' Adjust if AutoEmailCode is not needed

'    '    ' Execute the stored procedure to fill the dataset with file paths
'    '    If sql_handler.FillDataSet(strSQL, dsPaths, CommandType.StoredProcedure) Then
'    '        Return dsPaths
'    '    Else
'    '        MessageBox.Show("Failed to retrieve paths from the database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'    '        Return Nothing
'    '    End If
'    'End Function

'    Private Function GetLatestWaferAllocationFile(sourceDir As String) As FileInfo
'        ' Get all files starting with "WaferAllocation_" in the source directory
'        Dim dirInfo As New DirectoryInfo(sourceDir)
'        Dim files As FileInfo() = dirInfo.GetFiles("WaferAllocation_*.xlsx").OrderByDescending(Function(f) f.LastWriteTime).ToArray()

'        ' Return the latest file (first file in the ordered list)
'        If files.Any() Then
'            Return files(0)
'        End If

'        Return Nothing
'    End Function

'    Private Function GetLatestWaferAllocationFile2(sourceDir As String) As FileInfo
'        ' Get all files starting with "WaferAllocation_" in the source directory
'        Dim dirInfo As New DirectoryInfo(sourceDir)
'        Dim files As FileInfo() = dirInfo.GetFiles("ONSemi_PriceAlert_*.xlsx").OrderByDescending(Function(f) f.LastWriteTime).ToArray()

'        ' Return the latest file (first file in the ordered list)
'        If files.Any() Then
'            Return files(0)
'        End If

'        Return Nothing
'    End Function

'    Private Function ReadDataFromExcel(filePath As String) As List(Of String())
'        Dim data As New List(Of String())

'        ' Check if the file exists
'        If File.Exists(filePath) Then
'            Using package As New ExcelPackage(New FileInfo(filePath))
'                If package.Workbook IsNot Nothing Then
'                    ' Get the first worksheet (adjust as needed)
'                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()
'                    If worksheet IsNot Nothing Then
'                        Dim rowCount As Integer = worksheet.Dimension.Rows
'                        Dim colCount As Integer = Math.Min(worksheet.Dimension.Columns, 11) ' maximum of 11 columns

'                        For row = 1 To rowCount
'                            Dim rowData(colCount - 1) As String
'                            For col = 1 To colCount

'                                Dim cellValue As Object = worksheet.Cells(row, col).Value
'                                If cellValue IsNot Nothing Then
'                                    rowData(col - 1) = cellValue.ToString()
'                                Else
'                                    rowData(col - 1) = "" ' or handle null case as needed
'                                End If
'                            Next
'                            data.Add(rowData)
'                        Next
'                    Else
'                        MessageBox.Show("Worksheet not found in Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'                    End If
'                Else
'                    MessageBox.Show("Workbook not found in Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'                End If
'            End Using
'        Else
'            MessageBox.Show("Excel file not found at specified path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If

'        Return data
'    End Function

'    Private Function ReadDataFromExcel2(filePath As String) As List(Of String())
'        Dim data As New List(Of String())

'        ' Check if the file exists
'        If File.Exists(filePath) Then
'            Using package As New ExcelPackage(New FileInfo(filePath))
'                If package.Workbook IsNot Nothing Then
'                    ' Get the first worksheet (adjust as needed)
'                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()
'                    If worksheet IsNot Nothing Then
'                        Dim rowCount As Integer = worksheet.Dimension.Rows
'                        Dim colCount As Integer = Math.Min(worksheet.Dimension.Columns, 11) ' maximum of 11 columns

'                        For row = 1 To rowCount
'                            Dim rowData(colCount - 1) As String
'                            For col = 1 To colCount

'                                Dim cellValue As Object = worksheet.Cells(row, col).Value
'                                If cellValue IsNot Nothing Then
'                                    rowData(col - 1) = cellValue.ToString()
'                                Else
'                                    rowData(col - 1) = "" ' or handle null case as needed
'                                End If
'                            Next
'                            data.Add(rowData)
'                        Next
'                    Else
'                        MessageBox.Show("Worksheet not found in Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'                    End If
'                Else
'                    MessageBox.Show("Workbook not found in Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'                End If
'            End Using
'        Else
'            MessageBox.Show("Excel file not found at specified path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If

'        Return data
'    End Function

'    Private Sub SendEmailWithExcelData(data As List(Of String()), filePath As String)
'        Dim Datenow As DateTime = DateTime.Now
'        EmailSubject = "Wafer Allocation Details " & Datenow.ToString()

'        Dim message As New StringBuilder()
'        message.AppendLine("<html><body>")
'        message.AppendLine("<table border='1' cellpadding='5' cellspacing='0'>")

'        message.AppendLine("<tr>")
'        For Each columnName In {"Customer", "Device", "Die Part", "ProductID", "Package", "Lead Count", "Wafer Lot", "PID", "Quantity", "Start Date", "Load Qty"}
'            message.AppendLine("<th style='background-color: #007bff; color: #ffffff; font-weight: bold; font-size: 14px; font-family: Times New Roman, Times, serif;'>" & columnName & "</th>")
'        Next
'        message.AppendLine("</tr>")

'        For rowIdx = 1 To data.Count - 1
'            message.AppendLine("<tr>")
'            For Each cellValue In data(rowIdx)
'                message.AppendLine("<td style='font-family: Times New Roman, Times, serif; font-size: 11pt;'>" & cellValue & "</td>")
'            Next
'            message.AppendLine("</tr>")
'        Next

'        message.AppendLine("</table>")
'        message.AppendLine("</body></html>")

'        message.AppendLine("<p><strong>Please see attached allocated PID for today.</strong></p>")

'        ' Send email
'        em.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
'    End Sub



'    Private Sub SendEmailWithExcelDataNoDevice(data As List(Of String()), filePath As String)
'    Dim Datenow As DateTime = DateTime.Now
'    EmailSubject = "onsemi price alert " & Datenow.ToString()

'        Dim message As New StringBuilder()
'        message.AppendLine("<html><body>")
'        message.AppendLine("<table border='1' cellpadding='5' cellspacing='0'>")

'        message.AppendLine("<tr>")
'        For Each columnName In {"1", "2", "3", "4", "5", "6"}
'            message.AppendLine("<th style='background-color: #007bff; color: #ffffff; font-weight: bold; font-size: 14px; font-family: Times New Roman, Times, serif;'>" & columnName & "</th>")
'        Next
'        message.AppendLine("</tr>")

'        For rowIdx = 1 To data.Count - 1
'        message.AppendLine("<tr>")
'        For Each cellValue In data(rowIdx)
'            message.AppendLine("<td style='font-family: Times New Roman, Times, serif; font-size: 11pt;'>" & cellValue & "</td>")
'        Next
'        message.AppendLine("</tr>")
'    Next

'    message.AppendLine("</table>")
'    message.AppendLine("</body></html>")

'        message.AppendLine("<p><strong>Please see attached file for Price Matrix of ONSemi Dumped Devices as of " & Datenow.ToString & "</strong></p>")

'        ' Send email
'        em.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
'End Sub
'End Class