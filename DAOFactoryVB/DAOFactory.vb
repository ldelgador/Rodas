Public Class DAOFactory

    Public Enum TipoFactory
        Oracle = 1
        SQLServer = 2
        MySQL = 3
    End Enum

    Public Shared NivelLog As Integer = 0
    Public Shared FileLog As String = ""

    Public Shared Function GetFactory(ByVal Tipo As TipoFactory) As DAOFactory

        Select Case Tipo
            Case TipoFactory.Oracle
                Return New OracleFactory()
            Case TipoFactory.SQLServer
                Return New SQLServerFactory
                'Case TipoFactory.MySQL
                '    Return New MySQLFactory
            Case Else
                Return Nothing
        End Select

    End Function

    'funciones a sobreescribir

    Public Overridable Function getPresenciaDAO(ByVal CadenaConexion As String) As DAO.PresenciaDAO

    End Function

    Public Overridable Function getBolsaPresenciaDAO(ByVal CadenaConexion As String) As DAO.BolsaPresenciaDAO

    End Function

    Public Overridable Function getParteTrabajoDAO(ByVal CadenaConexion As String) As DAO.ParteTrabajoDAO

    End Function
    Public Overridable Function getDepandDAO(ByVal CadenaConexion As String) As DAO.DepandDAO

    End Function

    Public Overridable Function getAlertasDAO(ByVal CadenaConexion As String) As DAO.AlertasDAO

    End Function

End Class
