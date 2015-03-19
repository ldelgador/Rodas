Public Class OracleFactory
    Inherits DAOFactory

    Public Overrides Function getPresenciaDAO(ByVal CadenaConexion As String) As DAO.PresenciaDAO
        Return New DAO.OraclePresenciaDAO(CadenaConexion, NivelLog, FileLog)
    End Function

    Public Overrides Function getBolsaPresenciaDAO(ByVal CadenaConexion As String) As DAO.BolsaPresenciaDAO
        Return New DAO.OracleBolsaNegocioDAO(CadenaConexion, NivelLog, FileLog)
    End Function

    Public Overrides Function getParteTrabajoDAO(ByVal CadenaConexion As String) As DAO.ParteTrabajoDAO
        'Return New DAO.OracleParteTrabajoDAO(CadenaConexion)
    End Function


    Public Overrides Function getDepandDAO(ByVal CadenaConexion As String) As DAO.DepandDAO
        Return New DAO.OracleDepandDAO(CadenaConexion, NivelLog, FileLog)
    End Function

    Public Overrides Function getAlertasDAO(ByVal CadenaConexion As String) As DAO.AlertasDAO
        Return New DAO.OracleAlertasDAO(CadenaConexion, NivelLog, FileLog)
    End Function
End Class
