Public Class SQLServerFactory
    Inherits DAOFactory

    Public Overrides Function getPresenciaDAO(ByVal CadenaConexion As String) As DAO.PresenciaDAO
        'Return New SQLServerPresenciaDAO(CadenaConexion)
    End Function

    Public Overrides Function getParteTrabajoDAO(ByVal CadenaConexion As String) As DAO.ParteTrabajoDAO
        Return New DAO.SQLServerParteTrabajoDAO(CadenaConexion, NivelLog, FileLog)
    End Function

End Class
