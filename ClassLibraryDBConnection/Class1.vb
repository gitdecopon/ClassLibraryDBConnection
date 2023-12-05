Imports System.Data.SQLite
Imports System.Data.SqlClient
Public Class ClassSqliteAccess
    Private _sqlSeesion As SQLiteConnection
    Private _sqlCommand As SQLiteCommand
    Private _sqlDataAdapter As SQLiteDataAdapter
    Private _sqlTrn As SQLiteTransaction

    Public Sub New(ByVal ConnectionString As String)
        Dim connection As New SQLiteConnection
        _sqlSeesion = connection
        _sqlSeesion.ConnectionString = ConnectionString
        _sqlSeesion.Open()

        _sqlCommand = New SQLiteCommand
        _sqlCommand.Connection = _sqlSeesion

    End Sub

    Public Sub DBClose()
        _sqlCommand.Dispose()
        _sqlSeesion.Close()
        _sqlSeesion.Dispose()
    End Sub

    Public Sub DBExec(ByVal sql As String)
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text
        _sqlCommand.ExecuteNonQuery()
    End Sub

    Public Function DBExecuteReader(ByVal sql As String) As SQLiteDataReader
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text

        Return _sqlCommand.ExecuteReader()
    End Function

    Public Function DBExecuteScalar(ByVal sql As String) As Object
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text

        Return _sqlCommand.ExecuteScalar()
    End Function

    Public Function DBDataAdapter(ByVal sql As String) As SQLiteDataAdapter
        _sqlDataAdapter = New SQLiteDataAdapter(sql, _sqlSeesion)

        Return _sqlDataAdapter
    End Function

    Public Sub DBExecuteNonQuery(ByVal sql As String)
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text
        _sqlCommand.ExecuteNonQuery()
    End Sub

    Public Sub DBBigintran()
        _sqlTrn = _sqlSeesion.BeginTransaction(IsolationLevel.ReadCommitted)
    End Sub

    Public Sub DBCommit()
        _sqlTrn.Commit()
    End Sub

    Public Sub DBRollback()
        _sqlTrn.Rollback()
    End Sub

    Public Sub DBVacum()
        _sqlCommand.CommandText = "Vacum"
        _sqlCommand.CommandType = CommandType.Text
        _sqlCommand.ExecuteNonQuery()
    End Sub
End Class

Public Class ClassSqlserverAccess
    Private _sqlSeesion As SqlConnection
    Private _sqlCommand As SqlCommand
    Private _sqlDataAdapter As SqlDataAdapter
    Private _sqlTrn As SqlTransaction

    Public Sub New(ByVal ConnectionString As String)
        Dim connection As New SqlConnection
        _sqlSeesion = connection
        _sqlSeesion.ConnectionString = ConnectionString
        _sqlSeesion.Open()

        _sqlCommand = New SqlCommand
        _sqlCommand.Connection = _sqlSeesion

    End Sub

    Public Sub DBClose()
        _sqlCommand.Dispose()
        _sqlSeesion.Close()
        _sqlSeesion.Dispose()
    End Sub

    Public Sub DBExec(ByVal sql As String)
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text
        _sqlCommand.ExecuteNonQuery()
    End Sub

    Public Function DBExecuteReader(ByVal sql As String) As SqlDataReader
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text

        Return _sqlCommand.ExecuteReader()
    End Function

    Public Function DBExecuteScalar(ByVal sql As String) As Object
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text

        Return _sqlCommand.ExecuteScalar()
    End Function

    Public Function DBDataAdapter(ByVal sql As String) As SqlDataAdapter
        _sqlDataAdapter = New SqlDataAdapter(sql, _sqlSeesion)

        Return _sqlDataAdapter
    End Function

    Public Sub DBExecuteNonQuery(ByVal sql As String)
        _sqlCommand.CommandText = sql
        _sqlCommand.CommandType = CommandType.Text
        _sqlCommand.ExecuteNonQuery()
    End Sub

    Public Sub DBBigintran()
        _sqlTrn = _sqlSeesion.BeginTransaction(IsolationLevel.ReadCommitted)
    End Sub

    Public Sub DBCommit()
        _sqlTrn.Commit()
    End Sub

    Public Sub DBRollback()
        _sqlTrn.Rollback()
    End Sub
End Class
