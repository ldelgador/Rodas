Public Class BaseDAO

    Private mCadenaConexion As String 'Cadena de conexion a la base de datos
    Protected WithEvents mConexion As OleDb.OleDbConnection 'Conexion a la base de datos

    Private mNivelLog As Integer = 0
    Private mFileLog As String = "DAO.Log"

    'Constructor
    Public Sub New(ByVal CadenaConexion As String, Optional ByVal pNivelLog As Integer = 0, Optional ByVal pFileLog As String = "OraclePresenciaDAO.Log")
        mCadenaConexion = CadenaConexion
        mNivelLog = pNivelLog
        mFileLog = pFileLog
    End Sub

    Private Sub mConexion_InfoMessage(ByVal sender As Object, ByVal e As System.Data.OleDb.OleDbInfoMessageEventArgs) Handles mConexion.InfoMessage
        If mNivelLog > 0 Then
            Haz_Log("Mensaje: " & e.Message, mFileLog)
        End If
    End Sub


    Public Function ConectaDAO() As Boolean
        'conecta a la base de datos

        Try
            mConexion = New OleDb.OleDbConnection(mCadenaConexion)
            mConexion.Open()
            Return True
        Catch ex As Exception
            Trata_Error("Error al Conectar", ex)
            Return False
        End Try

    End Function

    Public Function DesConectaDAO() As Boolean
        'desconecta de la base de datos

        Try
            If Not IsNothing(mConexion) Then
                mConexion.Close()
            End If
            Return True
        Catch ex As Exception
            Trata_Error("Error al Desconectar", ex)
            Return False
        End Try

    End Function

    Protected Sub Trata_Error(ByVal Mensaje As String, ByVal pex As Exception, Optional ByVal pSQL As String = "")
        'rutina de tratamiento de errores
        'hace un log en el archivo
        Try
            If mNivelLog > 0 Then
                If pSQL <> "" Then
                    Haz_Log("SQL: " & pSQL, mFileLog)
                End If
                Haz_Log("Mensaje:" & Mensaje & ": " & pex.Message, mFileLog)
                Haz_Log("Pila:" & pex.StackTrace, mFileLog)
            End If
        Catch ex As Exception

        End Try

        'vuelve a lanzar el error
        Throw pex

    End Sub

    Private Sub Haz_Log(ByVal Mensaje As String, ByVal pFile As String)
        'genera un registro en el archivo de log de la clase

        Try

            Dim i As String


            i = FileSystem.FreeFile()
            FileSystem.FileOpen(i, pFile, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
            FileSystem.WriteLine(i, Now & ": " & Mensaje)
            FileSystem.FileClose(i)

        Catch ex As Exception

        End Try

    End Sub
    Protected Sub Haz_Log(ByVal Mensaje As String, Optional ByVal pNivelLog As Integer = 0)
        'genera un registro en el archivo de log de la clase

        Try

            Dim i As String

            If pNivelLog <= mNivelLog Then
                i = FileSystem.FreeFile()
                FileSystem.FileOpen(i, mFileLog, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
                FileSystem.WriteLine(i, Now & ": " & Mensaje)
                FileSystem.FileClose(i)
            End If
        Catch ex As Exception

        End Try

    End Sub


End Class
