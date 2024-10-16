Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Data.SqlClient
Imports System.Data

''' <summary>
''' Author:         Ramil Grajo
''' Created Date:   1/24/2013
''' Purpose:        Revised SQL_Class.
''' </summary>
Public Class SQLHandler
    Private err_msg As String
    Private server As String, database As String, username As String, password As String
    Private sql_conn As SqlConnection
    Private cmd_param As SqlParameter()

    Private cmdOutputParams As List(Of String)
    Private cmdOutputValues As List(Of String)

    Public Sub New()
        'server = "192.168.5.153"
        'database = "WEBMES_CONN"
        'username = "sa"
        'password = "dnhk0723$%"

        server = "MSDYNAMICS-DB\AXDB"
        database = "MES_ATEC"
        username = "sa"
        password = "p@ssw0rd"

        'server = "(local)"
        'database = "MES_ATEC"
        'username = ""
        'password = ""

        'server = "ATEC-MES2"
        'database = "MES_ATEC"
        'username = "sa"
        'password = "enola845&*"

        sql_conn = New SqlConnection()
        err_msg = ""
    End Sub

    Public Sub SetToATECLogsheet()
        server = "MSDYNAMICS-DB\AXDB"
        database = "ATEC_Logsheets"
        username = "sa"
        password = "p@ssw0rd"
    End Sub


    Public Sub SetToCentralAccess()
        server = "MSDYNAMICS-DB\AXDB"
        database = "CentralAccess"
        username = "sa"
        password = "p@ssw0rd"

        'server = "(local)"
        'database = "CentralAccess"
        'username = ""
        'password = ""

        'server = "ATEC-MES2"
        'database = "CentralAccess"
        'username = "sa"
        'password = "enola845&*"
    End Sub

    Public Sub SetToMESSupport()
        'server = "192.168.5.153"
        'database = "WEBMES_CONN_Support"
        'username = "sa"
        'password = "dnhk0723$%"

        server = "MSDYNAMICS-DB\AXDB"
        database = "MES_ATEC_Support"
        username = "sa"
        password = "p@ssw0rd"

        'server = "(local)"
        'database = "MES_ATEC_Support"
        'username = ""
        'password = ""

        'server = "ATEC-MES2"
        'database = "MES_ATEC_Support"
        'username = "sa"
        'password = "enola845&*"
    End Sub

    Public Sub SetToAXDB()
        server = "MSDYNAMICS-DB\AXDB"
        database = "AX2009DB"
        username = "sa"
        password = "p@ssw0rd"
    End Sub

    Public Sub SetToNAVISION()
        server = "NAVISION-SERVER"
        database = "SmartTrack 3.25"
        username = "sa"
        password = "25hpw2k30304$"
    End Sub

    Public Sub SetToEMMS1()
        server = "MSDYNAMICS-DB\AXDB"
        database = "EMMS1"
        username = "sa"
        password = "p@ssw0rd"
    End Sub

    Public Function GetErrorMessage() As String
        Return "Error on " & err_msg
    End Function

    Public Function OpenConnection() As Boolean
        Try
            ' Check if the connection is already open and close it if necessary
            If sql_conn.State = ConnectionState.Open Then
                sql_conn.Close()
            End If

            Dim connection As String = BuildConnectionString()

            ' Set the connection string and open the connection
            sql_conn.ConnectionString = connection
            sql_conn.Open()

            ' Return True to indicate successful connection
            Return True
        Catch ex As Exception
            ' Capture the error message for further analysis
            err_msg = "Open Connection: " & ex.Message
            ' Return False to indicate connection failure
            Return False
        End Try
    End Function

    Private Function BuildConnectionString() As String
        If String.IsNullOrEmpty(username) Then
            ' Use Windows Authentication
            Return $"server={server}; database={database}; connection timeout=30; Trusted_Connection=Yes;"
        Else
            ' Use SQL Server Authentication
            Return $"server={server}; database={database}; user id={username}; password={password}; connection timeout=30"
        End If
    End Function


    Public Function CloseConnection() As Boolean
        Try
            sql_conn.Close()
            Return True
        Catch ex As Exception
            err_msg = ex.Message
            Return False
        End Try
    End Function

    Public Function CreateParameter(ByVal size As Integer, Optional ByVal _redim As Boolean = False) As Boolean
        Try
            If size = 0 Then
                err_msg = "Create Parameter: Invalid size of parameters"
                Return False
            End If
            cmdOutputParams = New List(Of String)()
            cmdOutputValues = New List(Of String)()
            If _redim Then
                ReDim Preserve cmd_param(size - 1)
            Else
                cmd_param = New SqlParameter(size - 1) {}
            End If

            Return True
        Catch ex As Exception
            err_msg = "Create Parameter: " & ex.Message
            Return False
        End Try
    End Function

    Public Function SetParameterValues(ByVal position As Integer, ByVal paramName As String, ByVal type As System.Data.SqlDbType, ByVal value As Object, Optional ByVal direction As ParameterDirection = ParameterDirection.Input) As Boolean
        Try
            If cmd_param Is Nothing Then
                err_msg = "Set Parameter Values: Invalid size of parameters"
                Return False
            End If

            cmd_param(position) = New SqlParameter(paramName, type)
            cmd_param(position).Direction = direction

            If direction = ParameterDirection.Output Then
                cmdOutputParams.Add(paramName)
                If type = SqlDbType.NVarChar OrElse type = SqlDbType.VarChar Then
                    cmd_param(position).Size = 4000
                End If
            Else
                cmd_param(position).Value = value
            End If


            Return True
        Catch ex As Exception
            err_msg = "Set Parameter Values: " & ex.Message
            Return False
        End Try

    End Function

    Private Function AttachParameter(ByRef cmd As SqlCommand) As Boolean
        Try
            If cmd_param IsNot Nothing AndAlso cmd_param.Length > 0 Then
                For Each p As Object In cmd_param
                    cmd.Parameters.Add(p)
                Next
            Else
                err_msg = "Attach Parameter: Invalid size of parameters"
                Return False
            End If
            Return True
        Catch ex As Exception
            err_msg = "Attach Parameter: " & ex.Message
            Return False
        End Try
    End Function

    Private Sub SetOutputParamValues(ByVal cmd As SqlCommand)
        If cmdOutputParams IsNot Nothing Then
            For Each output As String In cmdOutputParams
                cmdOutputValues.Add(cmd.Parameters(output).Value.ToString())
            Next
        End If
    End Sub

    Public Function GetOutputParamValue(ByVal paramName As String) As String
        For i As Integer = 0 To cmdOutputParams.Count - 1
            If cmdOutputParams(i) = paramName Then
                Return cmdOutputValues(i)
            End If
        Next
        Return ""
    End Function

    Public Function ExecuteNonQuery(ByVal sql_string As String, ByVal command_type As System.Data.CommandType) As Boolean
        Try
            If OpenConnection() Then
                Dim command As New SqlCommand(sql_string, sql_conn)
                command.CommandTimeout = 99999
                command.CommandType = command_type

                If cmd_param IsNot Nothing Then
                    AttachParameter(command)
                End If

                command.ExecuteNonQuery()
                CloseConnection()

                If cmd_param IsNot Nothing Then
                    SetOutputParamValues(command)
                End If

                command = Nothing

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            err_msg = "Execute Non Query: " & ex.Message
            Return False
        End Try
    End Function

    Public Function FillDataReader(ByVal sql_string As String, ByRef dr As SqlDataReader, ByVal command_type As System.Data.CommandType) As Boolean
        Try
            Dim command As New SqlCommand(sql_string, sql_conn)
            command.CommandTimeout = 999999
            command.CommandType = command_type

            If cmd_param IsNot Nothing Then
                AttachParameter(command)
            End If

            dr = command.ExecuteReader(CommandBehavior.CloseConnection)

            If cmd_param IsNot Nothing Then
                SetOutputParamValues(command)
            End If

            command = Nothing

            Return True
        Catch ex As Exception
            err_msg = "Fill DataReader: " & ex.Message
            Return False
        End Try
    End Function

    Public Function FillDataSet(ByVal sql_string As String, ByRef ds As DataSet, ByVal command_type As System.Data.CommandType) As Boolean
        Try
            If OpenConnection() Then
                Dim command As New SqlCommand(sql_string, sql_conn)
                Dim da As SqlDataAdapter = Nothing
                command.CommandTimeout = 999999
                command.CommandType = command_type

                If cmd_param IsNot Nothing Then
                    AttachParameter(command)
                End If

                da = New SqlDataAdapter(command)
                da.Fill(ds)

                command = Nothing
                da = Nothing

                CloseConnection()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            err_msg = "Fill DataSet: " & ex.Message
            Return False
        End Try
    End Function

    Public Function GetSQLDbType(ByVal dataType As String) As SqlDbType
        Select Case dataType.ToLower
            Case "binary" : Return SqlDbType.VarBinary
            Case "bit" : Return SqlDbType.Bit
            Case "char" : Return SqlDbType.Char
            Case "date" : Return SqlDbType.Date
            Case "datetime" : Return SqlDbType.DateTime
            Case "datetime2" : Return SqlDbType.DateTime2
            Case "datetimeoffset" : Return SqlDbType.DateTimeOffset
            Case "decimal" : Return SqlDbType.Decimal
            Case "float" : Return SqlDbType.Float
            Case "image" : Return SqlDbType.Binary
            Case "int" : Return SqlDbType.Int
            Case "money" : Return SqlDbType.Money
            Case "nchar" : Return SqlDbType.NChar
            Case "ntext" : Return SqlDbType.NText
            Case "numeric" : Return SqlDbType.Decimal
            Case "nvarchar" : Return SqlDbType.NVarChar
            Case "real" : Return SqlDbType.Real
            Case "rowversion" : Return SqlDbType.Timestamp
            Case "smalldatetime" : Return SqlDbType.DateTime
            Case "smallint" : Return SqlDbType.SmallInt
            Case "smallmoney" : Return SqlDbType.SmallMoney
            Case "sql_variant" : Return SqlDbType.Variant
            Case "text" : Return SqlDbType.Text
            Case "time" : Return SqlDbType.Time
            Case "timestamp" : Return SqlDbType.Timestamp
            Case "tinyint" : Return SqlDbType.TinyInt
            Case "uniqueidentifier" : Return SqlDbType.UniqueIdentifier
            Case "varbinary" : Return SqlDbType.VarBinary
            Case "varchar" : Return SqlDbType.VarChar
            Case "xml" : Return SqlDbType.Xml
        End Select
        Return SqlDbType.VarChar
    End Function
End Class
