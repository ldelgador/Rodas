Public Class MySQLFactory
    Inherits DAOFactory

    Public Overrides Function getPresenciaDAO(ByVal CadenaConexion As String) As DAO.PresenciaDAO
        Return New DAO.MySQLPresenciaDAO(CadenaConexion, NivelLog, FileLog)
    End Function

    Public Overrides Function getParteTrabajoDAO(ByVal CadenaConexion As String) As DAO.ParteTrabajoDAO
        'Return New DAO.OracleParteTrabajoDAO(CadenaConexion)
    End Function


    Public Overrides Function getDepandDAO(ByVal CadenaConexion As String) As DAO.DepandDAO
        'Return New DAO.MySQLDepandDAO(CadenaConexion, NivelLog, FileLog)
    End Function

    Public Overrides Function getAlertasDAO(ByVal CadenaConexion As String) As DAO.AlertasDAO
        'Return New DAO.MySQLAlertasDAO(CadenaConexion, NivelLog, FileLog)
    End Function
End Class
