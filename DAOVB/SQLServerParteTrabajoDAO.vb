Public Class SQLServerParteTrabajoDAO
    Inherits BaseDAO
    Implements ParteTrabajoDAO

    Public Sub New(ByVal CadenaConexion As String, Optional ByVal pNivelLog As Integer = 0, Optional ByVal pFileLog As String = "OraclePresenciaDAO.Log")
        MyBase.New(CadenaConexion, pNivelLog, pFileLog)
    End Sub

    Public Function Conecta() As Boolean Implements ParteTrabajoDAO.Conecta
        'conecta a la base de datos

        Return MyBase.ConectaDAO()

    End Function

    Public Function DesConecta() As Boolean Implements ParteTrabajoDAO.DesConecta
        'desconecta de la base de datos

        Return MyBase.DesConectaDAO()

    End Function


    Public Function Lista_Tareas(Optional ByVal pIDUsuario As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pID_Tarea As Long = -1, Optional ByVal pID_Proyecto As String = "", Optional ByVal lista_Orden As String = "", Optional ByVal pSiguienteResponsable As String = "", Optional ByVal pUltimoResponsable As String = "", Optional ByVal lista_Tipo As String = "", Optional ByVal lista_Estado As String = "", Optional ByVal pEstado_Nulo As String = Nothing, Optional ByVal pFin_Nulo As String = Nothing) As OleDb.OleDbDataReader _
    Implements ParteTrabajoDAO.Lista_Tareas

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand()
        Dim mreader As OleDb.OleDbDataReader

        Try

            mConsulta = "SELECT ID_Tarea, ID_Usuario, ID_Proyecto, ID_Tipo, Comienzo, fin"
            mConsulta = mConsulta & " ,Codigo_Tarea, Descripcion, Observaciones, Estado, Modificado, ID_SiguienteResponsable "
            mConsulta = mConsulta & " FROM Tarea "
            If pIDUsuario <> "" Then
                mWhere = mWhere & " WHERE ID_Usuario = '" & pIDUsuario & "'"
            End If
            If pFechaDesde <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " Comienzo <= cast('" & pFechaDesde & " 23:59' as smalldatetime)"
            End If
            If pFechaHasta <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " Comienzo >= cast('" & pFechaHasta & " 00:00' as smalldatetime)"
            End If
            If pID_Tarea > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Tarea = " & pID_Tarea
            End If

            If pID_Proyecto <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pSiguienteResponsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_SiguienteResponsable  ='" & pSiguienteResponsable & "'"
            End If
            If pUltimoResponsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_UltimoResponsable ='" & pUltimoResponsable & "'"
            End If

            If lista_Estado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND ("
                Else
                    mWhere = mWhere & " WHERE ("
                End If
                mWhere = mWhere & " Estado  IN (" & lista_Estado & ")"
                If IsNothing(pEstado_Nulo) Then mWhere = mWhere & ")"
            End If

            If Not IsNothing(pEstado_Nulo) Then
                If lista_Estado <> "" Then
                    mWhere = mWhere & " OR "
                Else
                    If mWhere <> "" Then
                        mWhere = mWhere & " AND "
                    Else
                        mWhere = mWhere & " WHERE "
                    End If
                End If
                If pEstado_Nulo = "S" Then
                    mWhere = mWhere & " Estado IS NULL"
                Else
                    mWhere = mWhere & " Estado IS NOT NULL"
                End If
                If lista_Estado <> "" Then mWhere = mWhere & ")"
            End If

            If lista_Tipo <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Tipo  IN ('" & lista_Tipo & "')"
            End If

            If Not IsNothing(pFin_Nulo) Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                If pFin_Nulo = "S" Then
                    mWhere = mWhere & " Fin IS NULL"
                Else
                    mWhere = mWhere & " Fin IS NOT NULL"
                End If
            End If
            mConsulta = mConsulta & mWhere
            If lista_Orden = "" Then
                mConsulta = mConsulta & " ORDER BY Comienzo"
            Else
                mConsulta = mConsulta & " ORDER BY " & lista_Orden
            End If
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Inserta_Proyecto(ByVal pID_Proyecto As String, ByVal pNombre As String, ByVal pDescripcion As String, ByVal pID_Cliente As String, ByVal pFecha_Inicio As String, ByVal pFecha_Final As String, ByVal pHoras_Total As Integer, ByVal pObservaciones As String) As Boolean _
    Implements ParteTrabajoDAO.Inserta_Proyecto

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand()

        Try

            mConsulta = "INSERT INTO Proyecto(ID_Proyecto, Nombre, Descripcion, ID_Cliente,  Fecha_Inicio, Fecha_Final, Horas_Total, Observaciones)"
            mConsulta = mConsulta & " VALUES('" & pID_Proyecto & "','" & pNombre & "','" & pDescripcion & "','"
            mConsulta = mConsulta & pID_Cliente & "','" & pFecha_Inicio & "',"
            If pFecha_Final <> "" Then
                mConsulta = mConsulta & "'" & pFecha_Final & "'"
            Else
                mConsulta = mConsulta & "null"
            End If
            mConsulta = mConsulta & "," & pHoras_Total
            If pObservaciones <> "" Then
                mConsulta = mConsulta & ",'" & pObservaciones & "')"
            Else
                mConsulta = mConsulta & ",null)"
            End If

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            mConexion.Close()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Asignaciones_Proyectos(ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pID_Responsable As String) As OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Asignaciones_Proyectos
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand()
        Dim mreader As OleDb.OleDbDataReader

        Try

            mConsulta = "SELECT ID_Usuario, ID_Proyecto, ID_Responsable FROM Asignacion_Proyecto "
            If pID_Usuario <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Usuario = '" & pID_Usuario & "'"
            End If
            If pID_Proyecto <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Responsable = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & mWhere
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Usuarios(Optional ByVal pID_Usuario As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Usuarios
        Dim mSQL As String

        Try

            mSQL = "SELECT ID_Usuario, Nombre, Apellidos, Clave, Observaciones, Admin, Email "
            mSQL = mSQL & " FROM Usuario"
            If pID_Usuario <> "" Then
                mSQL = mSQL & " WHERE ID_Usuario IN ('" & pID_Usuario & "')"
            Else
                mSQL = mSQL & " ORDER BY Apellidos, nombre"
            End If

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Actualiza_Fin_Tarea(ByVal pID_Tarea As Long, ByVal pFin As Date) As Boolean Implements ParteTrabajoDAO.Actualiza_Fin_Tarea
        Dim mSQL As String
        Try
            mSQL = "UPDATE Tarea SET Fin = '" & pFin & "'"
            mSQL = mSQL & " WHERE ID_Tarea = " & pID_Tarea
            Dim mCommand As New OleDb.OleDbCommand()
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Proyectos(Optional ByVal pID_Proyecto As String = "", Optional ByVal pID_Cliente As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Proyectos
        Dim mConsulta As String
        Dim mWhere As String
        'mSQL = "SELECT ID_Proyecto, Nombre, Descripcion, ID_Cliente, Fecha_Inicio, Fecha_Final, Horas_Total, Observaciones"
        'mSQL = mSQL & " FROM Proyecto"
        'If pID_Proyecto <> "" Then
        '    mSQL = mSQL & " WHERE ID_Proyecto = '" & pID_Proyecto & "'"
        'End If
        'mSQL = mSQL & " ORDER BY Nombre"

        Try

            mConsulta = "SELECT Proyecto.ID_Proyecto, Proyecto.Nombre AS Proyecto_Nombre, Descripcion, Cliente.ID_Cliente, Cliente.Nombre AS Cliente_Nombre, "
            mConsulta = mConsulta & " Fecha_Inicio, Fecha_Final, Observaciones"
            mConsulta = mConsulta & " FROM Proyecto INNER JOIN Cliente on Proyecto.ID_Cliente = Cliente.ID_Cliente "
            If pID_Proyecto <> "" Then
                mWhere = " WHERE ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pID_Cliente <> "" Then
                If mWhere <> "" Then
                    mWhere = " WHERE "
                Else
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " Proyecto.ID_Cliente = '" & pID_Cliente & "'"
            End If
            mConsulta = mConsulta & mWhere
            mConsulta = mConsulta & " order by Proyecto.Nombre "

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Busca_Proyectos(Optional ByVal pID_Proyecto As String = "", Optional ByVal pID_Cliente As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Busca_Proyectos
        Dim mConsulta As String
        Dim mWhere As String
        'mSQL = "SELECT ID_Proyecto, Nombre, Descripcion, ID_Cliente, Fecha_Inicio, Fecha_Final, Horas_Total, Observaciones"
        'mSQL = mSQL & " FROM Proyecto"
        'If pID_Proyecto <> "" Then
        '    mSQL = mSQL & " WHERE ID_Proyecto = '" & pID_Proyecto & "'"
        'End If
        'mSQL = mSQL & " ORDER BY Nombre"

        Try

            mConsulta = "SELECT Proyecto.ID_Proyecto, Proyecto.Nombre AS Proyecto_Nombre, Descripcion, Cliente.ID_Cliente, Cliente.Nombre AS Cliente_Nombre, "
            mConsulta = mConsulta & " Fecha_Inicio, Fecha_Final, Observaciones"
            mConsulta = mConsulta & " FROM Proyecto INNER JOIN Cliente on Proyecto.ID_Cliente = Cliente.ID_Cliente "
            If pID_Proyecto <> "" Then
                mWhere = " WHERE ID_Proyecto Like '%" & pID_Proyecto & "%'"
            End If
            If pID_Cliente <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE "
                Else
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " Proyecto.ID_Cliente = '" & pID_Cliente & "'"
            End If
            mConsulta = mConsulta & mWhere
            mConsulta = mConsulta & " order by Proyecto.Nombre "

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Clientes(Optional ByVal pID_Cliente As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Clientes
        Dim mSQL As String

        Try

            mSQL = "SELECT ID_Cliente, Nombre"
            mSQL = mSQL & " FROM Cliente"
            If pID_Cliente <> "" Then
                mSQL = mSQL & " WHERE ID_Cliente = '" & pID_Cliente & "'"
            End If
            mSQL = mSQL & " ORDER BY Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Tarea_Actual(ByVal pID_Usuario As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tarea_Actual
        Dim mConsulta As String

        Try

            'consulta la tarea actual
            mConsulta = "SELECT Tarea.ID_Tarea, Tarea.ID_Proyecto, Proyecto.Nombre, Tarea.ID_Tipo, Tipo.Nombre, Tarea.Descripcion, Tarea.Comienzo, DateDiff(""mi"",Tarea.Comienzo, dbo.mysysdate(getdate())) as Duracion, Tarea.codigo_tarea"
            mConsulta = mConsulta & " FROM Tipo INNER JOIN (Proyecto INNER JOIN Tarea ON Proyecto.ID_Proyecto = Tarea.ID_Proyecto) ON Tipo.ID_Tipo = Tarea.ID_Tipo"
            mConsulta = mConsulta & " where Tarea.ID_Usuario = '" & pID_Usuario & "'"
            mConsulta = mConsulta & " and fin is null"
            mConsulta = mConsulta & " and cast(str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo)) as smalldatetime) = '" & Format(Ahora(), "dd/MM/yyyy") & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Tareas_Pendientes_Aprobar_por_Usuario(ByVal pID_Responsables As String, ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Pendientes_Aprobar_por_Usuario
        Dim mConsulta As String

        Try

            mConsulta = mConsulta & " SELECT Tarea.ID_Usuario, Tarea.ID_Proyecto, Sum(DATEDIFF(mi, Tarea.Comienzo, Tarea.Fin)) AS Duracion,"
            mConsulta = mConsulta & "(SELECT Usuario.Apellidos + ', ' + Usuario.Nombre AS Nombre FROM Usuario WHERE usuario.ID_Usuario = Tarea.id_usuario) as Nombre,"
            mConsulta = mConsulta & "(SELECT Proyecto.Nombre FROM Proyecto WHERE Proyecto.ID_Proyecto = Tarea.ID_Proyecto)AS Proyecto,"
            mConsulta = mConsulta & "(SELECT Cliente.Id_Cliente FROM Cliente,Proyecto WHERE Cliente.ID_Cliente = Proyecto.ID_Cliente and Proyecto.ID_Proyecto = Tarea.ID_Proyecto) AS Cliente,"
            mConsulta = mConsulta & "(SELECT Cliente.Nombre FROM Cliente,Proyecto WHERE Cliente.ID_Cliente = Proyecto.ID_Cliente and Proyecto.ID_Proyecto = Tarea.ID_Proyecto) AS NCliente"
            mConsulta = mConsulta & " FROM Tarea"
            mConsulta = mConsulta & " WHERE "
            mConsulta = mConsulta & "(Estado is null or Estado = 0)  AND "
            mConsulta = mConsulta & "Fin is not null "
            If pID_Responsables <> "" Then
                mConsulta = mConsulta & "AND ID_SiguienteResponsable IN ('" & pID_Responsables & "')"
            End If

            'filtros por fecha
            mConsulta = mConsulta & " AND Tarea.Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mConsulta = mConsulta & " AND Tarea.Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"

            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            mConsulta = mConsulta & " GROUP BY Tarea.ID_Usuario, Tarea.ID_Proyecto"
            mConsulta = mConsulta & " ORDER BY Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Ultima_Tarea(ByVal pID_Usuario As String, ByVal pFecha As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Ultima_Tarea

        Dim mConsulta As String

        Try

            'consulta la tarea actual
            'obtengo la ultima tarea realizada en esta fecha
            mConsulta = "Select * FROM Tarea WHERE ID_Usuario = '" & pID_Usuario & "'"
            mConsulta = mConsulta & " and Comienzo <= cast('" & Format(pFecha, "dd/MM/yyyy") & " 23:59' as smalldatetime)"
            mConsulta = mConsulta & " order by Comienzo Desc"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
			Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Ultimo_Codigo_Tarea() As OleDb.OleDbDataReader Implements ParteTrabajoDAO.Ultimo_Codigo_Tarea
        Dim mConsulta As String


        Try

            'consulta la tarea actual
            'obtengo la ultima tarea realizada en esta fecha
            mConsulta = "Select MAX(ID_Tarea) FROM Tarea"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Inserta_Tarea(ByVal pID_Tarea As Long, ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pID_Tipo As String, ByVal pCodigo_Tarea As String, ByVal pDescripcion As String, ByVal pFecha_Inicio As String, ByVal pFecha_Final As String, ByVal pID_SiguienteResponsable As String, Optional ByVal pEstado As Integer = 0) As Boolean Implements ParteTrabajoDAO.Inserta_Tarea

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand()

        Try

            mConsulta = "INSERT INTO Tarea(ID_Tarea,ID_Usuario,ID_Proyecto,ID_Tipo,Codigo_Tarea,Descripcion,Comienzo,Fin,ID_SiguienteResponsable, Estado)"

            mConsulta = mConsulta & " VALUES(" & pID_Tarea & ",'" & pID_Usuario & "','" & pID_Proyecto & "','" & pID_Tipo & "',"

            If pCodigo_Tarea <> "" Then
                mConsulta = mConsulta & " '" & pCodigo_Tarea & " ',"
            Else
                mConsulta = mConsulta & " Null,"
            End If

            If pDescripcion <> "" Then
                mConsulta = mConsulta & " '" & pDescripcion & " ',"
            Else
                mConsulta = mConsulta & " Null,"
            End If

            mConsulta = mConsulta & "'" & pFecha_Inicio & "',"

            If pFecha_Final <> "" Then
                mConsulta = mConsulta & "'" & pFecha_Final & "',"
            Else
                mConsulta = mConsulta & "Null,"
            End If

            If pID_SiguienteResponsable <> "" Then
                mConsulta = mConsulta & "'" & pID_SiguienteResponsable & "'"
            Else
                mConsulta = mConsulta & "Null"
            End If
            mConsulta = mConsulta & "," & pEstado & ")"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            mConexion.Close()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Lista_Proyectos_asignados(ByVal pID_Usuario As String, Optional ByVal pSoloProyectosActivosFecha As String = "", Optional ByVal pID_Cliente As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Proyectos_asignados

        Dim mCommand As New OleDb.OleDbCommand()
        Dim mreader As OleDb.OleDbDataReader

        Dim mconsulta As String

        Try

            'carga la lista de proyectos asignados
            mconsulta = "SELECT Proyecto.ID_Proyecto, Proyecto.Nombre as Proyecto_Nombre"
            mconsulta = mconsulta & " FROM Proyecto INNER JOIN Asignacion_Proyecto ON Proyecto.ID_Proyecto = Asignacion_Proyecto.ID_Proyecto"
            mconsulta = mconsulta & " WHERE Asignacion_Proyecto.ID_Usuario = '" & pID_Usuario & "'"
            If IsDate(pSoloProyectosActivosFecha) Then
                mconsulta = mconsulta & " AND (Proyecto.Fecha_final >= cast('" & pSoloProyectosActivosFecha & "' as datetime) "
                mconsulta = mconsulta & " OR Proyecto.Fecha_final is null)"
            End If
            If pID_Cliente <> "" Then
                mconsulta = mconsulta & " AND Proyecto.ID_Cliente = '" & pID_Cliente & "'"
            End If
            mconsulta = mconsulta & " ORDER BY Proyecto.Nombre"

            mCommand.Connection = mConexion
            mCommand.CommandText = mconsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mconsulta)

        End Try

    End Function

    Public Function Lista_Tipos(Optional ByVal pID_Tipo As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Tipos
        Dim mConsulta As String

        Try

            mConsulta = "SELECT ID_Tipo, Nombre"
            mConsulta = mConsulta & " FROM Tipo"
            If pID_Tipo <> "" Then
                mConsulta = mConsulta & " WHERE ID_Tipo = '" & pID_Tipo & "'"
            End If
            mConsulta = mConsulta & " ORDER BY Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Actualiza_Tarea(ByVal pLista_ID_Tarea As String, Optional ByVal pID_Usuario As String = Nothing, Optional ByVal pID_Proyecto As String = Nothing, Optional ByVal pID_Tipo As String = Nothing, Optional ByVal pCodigo_Tarea As String = Nothing, Optional ByVal pComienzo As String = Nothing, Optional ByVal pFin As String = Nothing, Optional ByVal pDescripcion As String = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pEstado As String = Nothing, Optional ByVal pModificado As String = Nothing, Optional ByVal pID_SiguienteResponsable As String = Nothing, Optional ByVal pID_UltimoResponsable As String = Nothing) As Boolean Implements ParteTrabajoDAO.Actualiza_Tarea

        Dim mConsulta As String

        Try

            mConsulta = "UPDATE Tarea SET"

            If Not pID_Usuario Is Nothing Then
                mConsulta = mConsulta & " ID_Tipo = '" & QuitaComilla(pID_Usuario) & "',"
            End If

            If Not pID_Proyecto Is Nothing Then
                mConsulta = mConsulta & " ID_Proyecto = '" & QuitaComilla(pID_Proyecto) & "', "
            End If

            If Not pID_Tipo Is Nothing Then
                mConsulta = mConsulta & " ID_Tipo = '" & QuitaComilla(pID_Tipo) & "',"
            End If

            If Not pCodigo_Tarea Is Nothing Then
                If pCodigo_Tarea = "" Then
                    mConsulta = mConsulta & " Codigo_Tarea = Null,"
                Else
                    mConsulta = mConsulta & " Codigo_Tarea = '" & QuitaComilla(pCodigo_Tarea) & "',"
                End If
            End If

            If Not pComienzo Is Nothing Then
                If pComienzo = "" Then
                    mConsulta = mConsulta & " Comienzo = Null,"
                Else
                    mConsulta = mConsulta & " Comienzo = '" & QuitaComilla(pComienzo) & "',"
                End If
            End If
            If Not pFin Is Nothing Then
                If pFin = "" Then
                    mConsulta = mConsulta & " Fin = Null,"
                Else
                    mConsulta = mConsulta & " Fin = '" & QuitaComilla(pFin) & "',"
                End If
            End If

            If Not pDescripcion Is Nothing Then
                If pDescripcion = "" Then
                    mConsulta = mConsulta & " Descripcion = Null,"
                Else
                    mConsulta = mConsulta & " Descripcion = '" & QuitaComilla(pDescripcion) & "',"
                End If

            End If

            If Not pObservaciones Is Nothing Then
                If pObservaciones = "" Then
                    mConsulta = mConsulta & " Observaciones = Null,"
                Else
                    mConsulta = mConsulta & " Observaciones = '" & QuitaComilla(pObservaciones) & "',"
                End If
            End If

            If Not pModificado Is Nothing Then
                mConsulta = mConsulta & " Modificado = '" & QuitaComilla(pModificado) & "',"
            End If

            If Not pID_SiguienteResponsable Is Nothing Then
                If pID_SiguienteResponsable = "" Then
                    mConsulta = mConsulta & " ID_SiguienteResponsable = Null,"
                Else
                    mConsulta = mConsulta & " ID_SiguienteResponsable = '" & QuitaComilla(pID_SiguienteResponsable) & "',"
                End If

            End If
            If Not pID_UltimoResponsable Is Nothing Then
                If pID_UltimoResponsable = "" Then
                    mConsulta = mConsulta & " ID_UltimoResponsable = Null,"
                Else
                    mConsulta = mConsulta & " ID_UltimoResponsable = '" & QuitaComilla(pID_UltimoResponsable) & "',"
                End If

            End If
            If Not pEstado Is Nothing Then
                mConsulta = mConsulta & " Estado = '" & pEstado & "',"
            End If

            mConsulta = Left(mConsulta, mConsulta.Length - 1)
            'mConsulta = mConsulta & " WHERE ID_Tarea = " & pID_Tarea
            mConsulta = mConsulta & " WHERE ID_Tarea IN (" & pLista_ID_Tarea & ")"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Borra_Tarea(ByVal pID_Tarea As Long) As Boolean Implements ParteTrabajoDAO.Borra_Tarea
        Dim mConsulta As String
        Dim nAfectados As Integer

        Try

            mConsulta = " DELETE FROM Tarea WHERE ID_Tarea = " & pID_Tarea

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            nAfectados = mCommand.ExecuteNonQuery()

            If nAfectados > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Proyectos_Asignados_con_Tareas_Pendientes(ByVal pID_Responsables As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Proyectos_Asignados_con_Tareas_Pendientes
        Dim mConsulta As String

        Try

            mConsulta = "SELECT distinct Proyecto.ID_Proyecto, Proyecto.Nombre"
            mConsulta = mConsulta & " FROM (Asignacion_Proyecto INNER JOIN Proyecto "
            mConsulta = mConsulta & " ON Asignacion_Proyecto.ID_Proyecto = Proyecto.ID_Proyecto) INNER JOIN TAREA"
            mConsulta = mConsulta & " ON Tarea.id_proyecto = Proyecto.id_proyecto"
            mConsulta = mConsulta & " WHERE Asignacion_Proyecto.ID_Responsable IN ('" & pID_Responsables & "')"
            mConsulta = mConsulta & " and tarea.id_siguienteResponsable  = Asignacion_Proyecto.ID_Responsable"
            mConsulta = mConsulta & " and (Estado is null or Estado = 0)  AND Fin is not null "
            mConsulta = mConsulta & " ORDER BY Proyecto.Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Tareas_Pendientes_Aprobar_por_Usuario_y_Tipo(ByVal pID_Responsable As String, ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Pendientes_Aprobar_por_Usuario_y_Tipo
        Dim mSQL As String

        Try

            mSQL = "select Tarea.ID_Tipo, Tipo.Nombre "
            mSQL = mSQL & " ,sum(DateDiff(""mi"", Tarea.Comienzo, Tarea.Fin)) as Horas "
            mSQL = mSQL & " ,min(Comienzo) as Inicio "
            mSQL = mSQL & " ,max(Fin) as Final "
            mSQL = mSQL & " from Tarea left join Tipo on Tarea.ID_Tipo = Tipo.ID_Tipo"
            mSQL = mSQL & " where ID_Usuario = '" & pID_Usuario & "'"
            mSQL = mSQL & " and ID_Proyecto = '" & pID_Proyecto & "'"
            mSQL = mSQL & " and ID_SiguienteResponsable IN ('" & pID_Responsable & "')"
            mSQL = mSQL & " and Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"
            mSQL = mSQL & " Group by Tarea.ID_Tipo, Tipo.Nombre"


            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Tareas_Aprobadas_por_Usuario_y_Tipo(ByVal pID_Responsable As String, ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Aprobadas_por_Usuario_y_Tipo
        Dim mSQL As String

        Try
            mSQL = "select Tarea.ID_Tipo, Tipo.Nombre "
            mSQL = mSQL & " ,sum(DateDiff(""mi"", Tarea.Comienzo, Tarea.Fin)) as Horas "
            mSQL = mSQL & " ,min(Comienzo) as Inicio "
            mSQL = mSQL & " ,max(Fin) as Final "
            mSQL = mSQL & " from Tarea left join Tipo on Tarea.ID_Tipo = Tipo.ID_Tipo"
            mSQL = mSQL & " where ID_Usuario = '" & pID_Usuario & "'"
            mSQL = mSQL & " and ID_Proyecto = '" & pID_Proyecto & "'"
            mSQL = mSQL & " and ID_UltimoResponsable IN ('" & pID_Responsable & "')"
            mSQL = mSQL & " and Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"
            mSQL = mSQL & " Group by Tarea.ID_Tipo, Tipo.Nombre"


            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Proyecto_Asignado_A_Usuario(ByVal pID_Usuario As String, ByVal pID_Proyecto As String) As Boolean Implements ParteTrabajoDAO.Proyecto_Asignado_A_Usuario
        Dim mConsulta As String

        Try

            mConsulta = "SELECT * FROM Asignacion_Proyecto WHERE ID_Usuario ='" & pID_Usuario & "'"
            mConsulta = mConsulta & " AND ID_PRoyecto = '" & pID_Proyecto & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mReader As OleDb.OleDbDataReader
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta

            mReader = mCommand.ExecuteReader
            If mReader.Read Then
                'ya esta asignado
                Proyecto_Asignado_A_Usuario = True
            Else
                'no esta asignado
                Proyecto_Asignado_A_Usuario = False
            End If
            mReader.Close()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Inserta_Asignacion_Proyecto(ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pID_Responsable As String) As Boolean Implements ParteTrabajoDAO.Inserta_Asignacion_Proyecto
        Dim mConsulta As String

        Try

            mConsulta = "INSERT INTO Asignacion_Proyecto (ID_Usuario, ID_Proyecto, ID_Responsable)"
            mConsulta = mConsulta & " VALUES ('" & pID_Usuario & "','" & pID_Proyecto & "','" & pID_Responsable & "')"

            Dim mCommand As New OleDb.OleDbCommand()

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Borra_Asignacion_Proyecto(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Proyecto As String = "", Optional ByVal pID_Responsable As String = "") As Boolean Implements ParteTrabajoDAO.Borra_Asignacion_Proyecto
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "DELETE Asignacion_Proyecto "
            If pID_Usuario <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Usuario = '" & pID_Usuario & "'"
            End If

            If pID_Proyecto <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Proyecto = '" & pID_Proyecto & "'"
            End If


            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_Responsable = '" & pID_Responsable & "'"
            End If


            Dim mCommand As New OleDb.OleDbCommand()


            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta + mWhere
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Responsables_Maximos_de_Proyecto(ByVal pID_Proyecto As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Responsables_Maximos_de_Proyecto
        ' Usuarios que son responsable de sí mismos
        Dim mSQL As String

        Try

            mSQL = "select Asignacion_Proyecto.ID_Usuario , Usuario.Nombre, Usuario.Apellidos from Asignacion_Proyecto inner join Usuario "
            mSQL = mSQL & " ON Asignacion_Proyecto.ID_Usuario = Usuario.ID_Usuario"
            mSQL = mSQL & " where Asignacion_Proyecto.ID_Usuario = Asignacion_Proyecto.ID_Responsable"
            mSQL = mSQL & " and Asignacion_Proyecto.ID_Proyecto = '" & pID_Proyecto & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mReader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader
            Return mReader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Responsables_Maximos_de_Proyecto2(ByVal pID_Proyecto As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Responsables_Maximos_de_Proyecto2
        ' Responsables que no están asignados al proyecto como usuarios
        Dim mSQL As String

        Try

            mSQL = "select Asignacion_Proyecto.ID_Responsable , Usuario.Nombre, Usuario.Apellidos from Asignacion_Proyecto inner join Usuario "
            mSQL = mSQL & " ON Asignacion_Proyecto.ID_Responsable = Usuario.ID_Usuario"
            mSQL = mSQL & " where Asignacion_Proyecto.ID_Responsable not in (SELECT ID_Usuario FROM Asignacion_Proyecto"
            mSQL = mSQL & " where Asignacion_Proyecto.ID_Proyecto = '" & pID_Proyecto & "')"
            mSQL = mSQL & " and Asignacion_Proyecto.ID_Proyecto = '" & pID_Proyecto & "'"
            mSQL = mSQL & " Group by Asignacion_Proyecto.ID_Responsable , Usuario.Nombre, Usuario.Apellidos "

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mReader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader
            Return mReader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Subordinados_de_Proyecto(ByVal pID_Proyecto As String, ByVal pID_Responsable As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Subordinados_de_Proyecto
        Dim mSQL As String


        Try

            mSQL = "select Asignacion_Proyecto.ID_Usuario , Usuario.Nombre, Usuario.Apellidos from Asignacion_Proyecto inner join Usuario "
            mSQL = mSQL & " ON Asignacion_Proyecto.ID_Usuario = Usuario.ID_Usuario"
            mSQL = mSQL & " where Asignacion_Proyecto.ID_Responsable = '" & pID_Responsable & "'"
            mSQL = mSQL & " and Asignacion_Proyecto.ID_Usuario <> '" & pID_Responsable & "'"
            mSQL = mSQL & " and Asignacion_Proyecto.ID_Proyecto = '" & pID_Proyecto & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mReader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader
            Return mReader

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Actualiza_Proyecto(ByVal pID_Proyecto_Ant As String, Optional ByVal pID_Proyecto As String = Nothing, Optional ByVal pNombre As String = Nothing, Optional ByVal pDescripcion As String = Nothing, Optional ByVal pID_Cliente As String = Nothing, Optional ByVal pFecha_Inicio As String = Nothing, Optional ByVal pFecha_Final As String = Nothing, Optional ByVal pObservaciones As String = Nothing) As Boolean Implements ParteTrabajoDAO.Actualiza_Proyecto
        Dim mconsulta As String

        Try

            mconsulta = "UPDATE Proyecto SET"

            If Not pID_Proyecto Is Nothing Then
                mconsulta = mconsulta & " ID_Proyecto = '" & QuitaComilla(pID_Proyecto) & "', "
            End If


            If Not pNombre Is Nothing Then
                If pNombre = "" Then
                    mconsulta = mconsulta & " Nombre = Null,"
                Else
                    mconsulta = mconsulta & " Nombre = '" & QuitaComilla(pNombre) & "',"
                End If
            End If

            If Not pDescripcion Is Nothing Then
                If pDescripcion = "" Then
                    mconsulta = mconsulta & " Descripcion = Null,"
                Else
                    mconsulta = mconsulta & " Descripcion = '" & QuitaComilla(pDescripcion) & "',"
                End If
            End If

            If Not pID_Cliente Is Nothing Then
                If pID_Cliente = "" Then
                    mconsulta = mconsulta & " ID_Cliente = Null,"
                Else
                    mconsulta = mconsulta & " ID_Cliente = '" & QuitaComilla(pID_Cliente) & "',"
                End If
            End If
            If Not pFecha_Inicio Is Nothing Then
                If pFecha_Inicio = "" Then
                    mconsulta = mconsulta & " Fecha_Inicio = Null,"
                Else
                    mconsulta = mconsulta & " Fecha_Inicio = '" & QuitaComilla(pFecha_Inicio) & "',"
                End If
            End If
            If Not pFecha_Final Is Nothing Then
                If pFecha_Final = "" Then
                    mconsulta = mconsulta & " Fecha_Final = Null,"
                Else
                    mconsulta = mconsulta & " Fecha_Final = '" & QuitaComilla(pFecha_Final) & "',"
                End If
            End If
            If Not pObservaciones Is Nothing Then
                If pObservaciones = "" Then
                    mconsulta = mconsulta & " Observaciones = Null,"
                Else
                    mconsulta = mconsulta & " Observaciones = '" & QuitaComilla(pObservaciones) & "',"
                End If
            End If
            mconsulta = Left(mconsulta, mconsulta.Length - 1)
            mconsulta = mconsulta & " WHERE ID_Proyecto = '" & pID_Proyecto_Ant & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mconsulta


            If mCommand.ExecuteNonQuery() > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mconsulta)

        End Try

    End Function


    Public Function Borra_Proyecto(ByVal pID_Proyecto As String) As Boolean Implements ParteTrabajoDAO.Borra_Proyecto
        Dim mConsulta As String
        Dim nAfectados As Integer

        Try

            mConsulta = " DELETE FROM Proyecto WHERE ID_Proyecto = '" & pID_Proyecto & "'"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            nAfectados = mCommand.ExecuteNonQuery()

            If nAfectados > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Descripcion_Tareas(ByVal pID_Proyecto As String) As OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Descripcion_Tareas
        Dim mConsulta As String

        Try

            mConsulta = "select descripcion"
            mConsulta = mConsulta & " from tarea"
            mConsulta = mConsulta & " where descripcion is not null"
            mConsulta = mConsulta & " and descripcion <> ''"
            mConsulta = mConsulta & " and fin is not null"
            mConsulta = mConsulta & " and ID_Proyecto = '" & pID_Proyecto & "'"

            mConsulta = mConsulta & " group by descripcion"
            'mConsulta = mConsulta & " order by max(fin) desc"
            mConsulta = mConsulta & " order by descripcion"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            Return mCommand.ExecuteReader()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function
    Public Function Lista_Codigo_Tareas(ByVal pID_Proyecto As String) As OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Codigo_Tareas

        Dim mConsulta As String

        Try

            mConsulta = "select codigo_tarea"
            mConsulta = mConsulta & " from tarea"
            mConsulta = mConsulta & " where codigo_tarea is not null"
            mConsulta = mConsulta & " and codigo_tarea <> ''"
            mConsulta = mConsulta & " and fin is not null"
            'If pID_Usuario <> "" Then
            '    mConsulta = mConsulta & " and ID_Usuario = '" & pID_Usuario & "'"
            'End If
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " and ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            mConsulta = mConsulta & " group by codigo_tarea"
            'mConsulta = mConsulta & " order by max(fin) desc"
            mConsulta = mConsulta & " order by codigo_tarea"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            Return mCommand.ExecuteReader()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Partes(ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pFechaHasta As Date, ByVal pFechaDesde As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Partes
        Dim mConsulta As String

        Try

            'mConsulta = "SELECT Tarea.ID_Tarea as Parte, Proyecto.Nombre as Proyecto, Tipo.Nombre as Tipo, Tarea.Codigo_Tarea as Tarea, format(Tarea.Comienzo,""dd/MM/yyyy hh:mm"") as Comienzo, format(Tarea.Fin,""hh:mm"") as Fin, format(Tarea.Fin - Tarea.Comienzo ,""hh:mm"") as Duración"
            mConsulta = "SELECT Tarea.ID_Tarea as Parte, Proyecto.Nombre as Proyecto, Tipo.Nombre as Tipo"
            mConsulta = mConsulta & ", Tarea.Comienzo, Tarea.Fin, Tarea.Fin - Tarea.Comienzo as Duración, Estado, Tarea.Codigo_Tarea "
            mConsulta = mConsulta & " FROM Tipo INNER JOIN (Proyecto INNER JOIN Tarea ON Proyecto.ID_Proyecto = Tarea.ID_Proyecto) ON Tipo.ID_Tipo = Tarea.ID_Tipo"
            mConsulta = mConsulta & " where Tarea.ID_Usuario = '" & pID_Usuario & "'"
            mConsulta = mConsulta & " and cast(str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo)) as smalldatetime) <= cast('" & Format(pFechaDesde, "dd/MM/yyyy") & "' as smalldatetime)"
            mConsulta = mConsulta & " and cast(str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo)) as smalldatetime) >= cast('" & Format(pFechaHasta, "dd/MM/yyyy") & "' as smalldatetime)"
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            mConsulta = mConsulta & " order by Comienzo"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            Return mCommand.ExecuteReader()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Partes_Totales(ByVal pID_Usuario As String, ByVal pID_Proyecto As String, ByVal pFechaDesde As Date, ByVal pFechaHasta As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Partes_Totales
        Dim mConsulta As String

        Try

            mConsulta = "SELECT count(Tarea.ID_Tarea) as cuenta, "
            mConsulta = mConsulta & " sum(datediff(""mi"",Tarea.Comienzo,Tarea.Fin)) as Duración "
            mConsulta = mConsulta & " FROM Tipo INNER JOIN (Proyecto INNER JOIN Tarea ON Proyecto.ID_Proyecto = Tarea.ID_Proyecto) ON Tipo.ID_Tipo = Tarea.ID_Tipo"
            mConsulta = mConsulta & " where Tarea.ID_Usuario = '" & pID_Usuario & "'"
            mConsulta = mConsulta & " and cast(str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo)) as smalldatetime) <= cast('" & Format(pFechaDesde, "dd/MM/yyyy") & "' as smalldatetime)"
            mConsulta = mConsulta & " and cast(str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo)) as smalldatetime) >= cast('" & Format(pFechaHasta, "dd/MM/yyyy") & "' as smalldatetime)"
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            Return mCommand.ExecuteReader()

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Tareas_Pendientes_Aprobar(ByVal pID_Responsable As String, ByVal pID_Usuario As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pID_Proyecto As String, ByVal pID_Tipo As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Pendientes_Aprobar
        Dim mConsulta As String

        Try

            mConsulta = "SELECT ID_Tarea, "
            'mConsulta = mConsulta & " Tarea.ID_Usuario, Usuario.Apellidos + ', ' + Usuario.Nombre as Nombre,"
            'mConsulta = mConsulta & " Tarea.ID_Proyecto, Proyecto.Nombre as PNombre, Proyecto.ID_Cliente, Cliente.Nombre as CNombre, "
            'mConsulta = mConsulta & " Tarea.ID_Tipo, Tipo.Nombre as TNombre,"
            mConsulta = mConsulta & " Tarea.Codigo_Tarea,"
            mConsulta = mConsulta & " Tarea.Comienzo, Tarea.Fin, "
            mConsulta = mConsulta & " Tarea.Fin - Tarea.Comienzo as Duracion,"
            mConsulta = mConsulta & " Tarea.Descripcion"
            mConsulta = mConsulta & " FROM"
            mConsulta = mConsulta & " Tarea " ', Usuario, Proyecto, Cliente, Tipo"
            mConsulta = mConsulta & " WHERE " 'Tarea.ID_Usuario = Usuario.ID_Usuario"
            'mConsulta = mConsulta & " AND Tarea.ID_Proyecto = Proyecto.ID_Proyecto"
            'mConsulta = mConsulta & " AND Proyecto.ID_Cliente = Cliente.ID_Cliente"
            'mConsulta = mConsulta & " AND Tarea.ID_Tipo = Tipo.ID_Tipo"
            mConsulta = mConsulta & " (Estado is null or Estado = 0)"
            mConsulta = mConsulta & " AND Fin is not null"
            mConsulta = mConsulta & " AND ID_SiguienteResponsable ='" & pID_Responsable & "'"
            If pID_Usuario <> "" Then
                mConsulta = mConsulta & " AND ID_Usuario ='" & pID_Usuario & "'"
            End If
            'filtros por fecha
            mConsulta = mConsulta & " AND Tarea.Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mConsulta = mConsulta & " AND Tarea.Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Tipo = " & pID_Tipo
            End If
            mConsulta = mConsulta & " ORDER BY Tarea.ID_Usuario, Comienzo"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Usuarios_Xml(ByVal pID_Usuario As String) As String Implements ParteTrabajoDAO.Lista_Usuarios_Xml

        Dim mSQL As String
        'Dim mConexion As New OleDb.OleDbConnection(CTE_Cadena_conexionPT)
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim mXML As String
        Dim auxstr As String

        Try

            'obtengo los datos de los usuarios
            mSQL = "SELECT ID_Usuario, Nombre, Apellidos, Observaciones, Email"
            mSQL = mSQL & " FROM Usuario "
            If Not pID_Usuario Is Nothing Then
                mSQL = mSQL & " where ID_Usuario in ('" & pID_Usuario & "')"
            End If
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)

            mDataset = New DataSet()
            mDataAdapter.Fill(mDataset, "Usuario")
            mDataset.DataSetName = "Usuarios"
            auxstr = mDataset.GetXml()
            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Totales_Usuario_Proyecto_Xml(ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pLista_Usuarios As String, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Usuario_Proyecto_Xml

        Dim mSQL As String

        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por usuario y proyecto
            mSQL = "SELECT usuario.ID_Usuario, tarea.ID_Proyecto, proyecto.nombre, sum(datediff(mi,tarea.comienzo, tarea.fin)) as total"
            mSQL = mSQL & " FROM (Usuario inner join tarea"
            mSQL = mSQL & " on usuario.id_usuario = tarea.id_usuario)"
            mSQL = mSQL & " inner join proyecto on "
            mSQL = mSQL & " tarea.id_proyecto = proyecto.id_proyecto"
            mSQL = mSQL & " where(Tarea.Fin Is Not null)"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            If Not pLista_Usuarios Is Nothing Then
                mSQL = mSQL & " and usuario.ID_Usuario in ('" & pLista_Usuarios & "')"
            End If
            mSQL = mSQL & " group by usuario.ID_Usuario, tarea.ID_Proyecto, proyecto.nombre"
            mSQL = mSQL & " order by sum(datediff(mi,tarea.comienzo, tarea.fin)) desc"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)

            mDataset = New DataSet()
            mDataAdapter.Fill(mDataset, "UsuarioProyecto")
            mDataset.DataSetName = "UsuarioProyectos"
            auxstr = mDataset.GetXml()
            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Totales_Usuario_Tipo_Xml(ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pLista_Usuarios As String, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Usuario_Tipo_Xml

        Dim mSQL As String

        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por usuario y tipo
            mSQL = "SELECT usuario.ID_Usuario, tarea.ID_tipo, tipo.nombre, sum(datediff(mi,tarea.comienzo, tarea.fin)) as total"
            mSQL = mSQL & " FROM (Usuario inner join tarea"
            mSQL = mSQL & " on usuario.id_usuario = tarea.id_usuario)"
            mSQL = mSQL & " inner join tipo on "
            mSQL = mSQL & " tarea.id_tipo = tipo.id_tipo"
            mSQL = mSQL & " where(Tarea.Fin Is Not null)"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            If Not pLista_Usuarios Is Nothing Then
                mSQL = mSQL & " and usuario.ID_Usuario in ('" & pLista_Usuarios & "')"
            End If
            mSQL = mSQL & " group by usuario.ID_Usuario, tarea.ID_tipo, tipo.nombre"
            mSQL = mSQL & " order by sum(datediff(mi,tarea.comienzo, tarea.fin)) desc"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataAdapter.Fill(mDataset, "UsuarioTipo")
            mDataset.DataSetName = "UsuarioTipos"

            mDataset = New DataSet()
            auxstr = mDataset.GetXml()
            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Totales_Usuario_Proyecto_Tipo_Xml(ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pLista_Usuarios As String, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Usuario_Proyecto_Tipo_Xml
        Dim mSQL As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try
            'totales por usuario y proyecto y tipo 
            mSQL = "SELECT usuario.ID_Usuario, tarea.id_proyecto, proyecto.nombre, "
            mSQL = mSQL & " tarea.ID_tipo, tipo.nombre, sum(datediff(mi,tarea.comienzo, tarea.fin)) as total"
            mSQL = mSQL & " FROM ((Usuario inner join tarea"
            mSQL = mSQL & " on usuario.id_usuario = tarea.id_usuario)"
            mSQL = mSQL & " inner join tipo on "
            mSQL = mSQL & " tarea.id_tipo = tipo.id_tipo) "
            mSQL = mSQL & " inner join proyecto"
            mSQL = mSQL & " on tarea.id_proyecto = proyecto.id_proyecto"
            mSQL = mSQL & " where(Tarea.Fin Is Not null)"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            If Not pLista_Usuarios Is Nothing Then
                mSQL = mSQL & " and usuario.ID_Usuario in ('" & pLista_Usuarios & "')"
            End If
            mSQL = mSQL & " group by usuario.ID_Usuario, tarea.id_proyecto, proyecto.nombre, tarea.ID_tipo, tipo.nombre"
            'mSQL = mSQL & " order by usuario.ID_Usuario, proyecto.nombre, tipo.nombre"
            mSQL = mSQL & " order by usuario.ID_Usuario asc, proyecto.nombre asc, sum(datediff(mi,tarea.comienzo, tarea.fin)) desc"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "UsuarioProyectoTipos"
            mDataAdapter.Fill(mDataset, "UsuarioProyectoTipo")
            auxstr = mDataset.GetXml()
            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Tareas_Realizadas_Xml(ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pLista_Usuarios As String, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Tareas_Realizadas_Xml
        Dim mSQL As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'tareas realizadas
            mSQL = "SELECT ID_Tarea, Tarea.ID_Usuario, Proyecto.ID_Proyecto, Proyecto.Nombre, Tipo.Nombre as Tipo, Codigo_Tarea, Comienzo , Fin, Tarea.Descripcion, Tarea.Observaciones, cast(datediff(""mi"",comienzo,fin)/60.00 as numeric(10,2)) as Duracion"
            mSQL = mSQL & " FROM (Tarea LEFT JOIN Tipo ON Tarea.ID_Tipo = Tipo.ID_Tipo)"
            mSQL = mSQL & " LEFT JOIN Usuario ON Tarea.ID_Usuario = Usuario.ID_Usuario"
            mSQL = mSQL & " LEFT JOIN Proyecto ON Proyecto.ID_Proyecto = Tarea.ID_Proyecto"
            mSQL = mSQL & " where(Tarea.Fin Is Not null)"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            If Not pLista_Usuarios Is Nothing Then
                mSQL = mSQL & " and Tarea.ID_Usuario in ('" & pLista_Usuarios & "')"
            End If
            mSQL = mSQL & " order by Proyecto.Nombre,comienzo"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "Tareas"
            mDataAdapter.Fill(mDataset, "Tarea")
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Lista_Proyectos_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Proyectos_Xml
        Dim mSQL As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'obtengo los datos del proyecto
            mSQL = "SELECT  "
            mSQL = mSQL & " Proyecto.ID_Proyecto, Proyecto.Nombre, Proyecto.Descripcion,  "
            mSQL = mSQL & " Cliente.Nombre as CNombre, "
            mSQL = mSQL & " Proyecto.Fecha_Inicio, Proyecto.Fecha_Final,"
            mSQL = mSQL & " cast(sum(dateDiff(""mi"",tarea.comienzo, Tarea.fin))/60.00 as numeric(9,2)) as Total"
            mSQL = mSQL & " FROM"
            mSQL = mSQL & " (Proyecto LEFT JOIN (Select * from Tarea where Tarea.fin is not null "
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " ) as tarea ON Proyecto.ID_Proyecto = Tarea.ID_Proyecto), Cliente"
            mSQL = mSQL & " WHERE"
            mSQL = mSQL & " Proyecto.ID_Cliente = cliente.ID_Cliente"
            mSQL = mSQL & " and Proyecto.ID_Proyecto = '" & pID_Proyecto & "'"
            mSQL = mSQL & " and tarea.fin is not null"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            mSQL = mSQL & " GROUP BY Proyecto.ID_Proyecto, Proyecto.Nombre, Proyecto.Descripcion,  "
            mSQL = mSQL & " Cliente.Nombre , "
            mSQL = mSQL & " Proyecto.Fecha_Inicio, Proyecto.Fecha_Final"

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "Proyectos"
            mDataAdapter.Fill(mDataset, "Proyecto")
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Lista_Totales_Proyecto_Usuario_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Proyecto_Usuario_Xml
        Dim mSQL As String
        Dim mSQL2 As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por usuario

            mSQL = "SELECT Usuario.ID_Usuario, Nombre, Apellidos, "
            mSQL = mSQL & " Count(1) as Cuenta, count(distinct str(day(comienzo)) + '/' + str(month(comienzo)) + '/' + str(year(comienzo))) as Dias,"
            mSQL = mSQL & " cast(sum(datediff(""mi"",comienzo,fin))/60.00 as numeric(10,2)) as Total  "

            'subconsulta de total de todos los proyectos
            mSQL2 = "(Select  cast(sum(datediff(""mi"",comienzo,fin))/60.00 as numeric(10,2)) from Tarea where Tarea.ID_Usuario = Usuario.ID_Usuario"
            If Not pIncluirNoAprobadas Then
                mSQL2 = mSQL2 & " and tarea.estado = 1"
            End If
            mSQL2 = mSQL2 & " and Tarea.fin Is Not null"
            'condiciones de fecha
            mSQL2 = mSQL2 & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            mSQL2 = mSQL2 & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            mSQL2 = mSQL2 & " ) as TotalOTros "

            mSQL = mSQL & "," & mSQL2

            mSQL = mSQL & " FROM Usuario,Tarea "
            mSQL = mSQL & " WHERE Usuario.ID_Usuario= Tarea.ID_Usuario "
            mSQL = mSQL & " AND Tarea.Fin is not null"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            'condiciones de fecha
            If Not IsDBNull(pFecha_Desde) Then
                mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            End If
            If Not IsDBNull(pFecha_Hasta) Then
                mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            End If

            mSQL = mSQL & " GROUP BY Usuario.ID_Usuario, Nombre, Apellidos"
            mSQL = mSQL & " ORDER BY Usuario.Apellidos"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "Usuarios"
            mDataAdapter.Fill(mDataset, "Usuario")
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Lista_Totales_Proyecto_Tipo_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Proyecto_Tipo_Xml
        Dim mSQL As String
        Dim mSQL2 As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por tipo
            mSQL = "SELECT Proyecto.Nombre, Tipo.Nombre, "
            mSQL = mSQL & " cast(sum(DateDiff(""mi"", comienzo, fin))/60.00 as numeric(10,2))"
            mSQL = mSQL & " FROM (Tarea left join Tipo"
            mSQL = mSQL & " ON Tarea.ID_Tipo = Tipo.ID_Tipo"
            mSQL = mSQL & " ) LEFT JOIN Proyecto"
            mSQL = mSQL & " ON Tarea.ID_Proyecto = Proyecto.ID_Proyecto"
            mSQL = mSQL & " where Tarea.fin Is Not null"
            mSQL = mSQL & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If


            'condiciones de fecha
            If Not IsDBNull(pFecha_Desde) Then
                mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            End If
            If Not IsDBNull(pFecha_Hasta) Then
                mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            End If

            mSQL = mSQL & " group by Proyecto.Nombre,  Tipo.Nombre"
            'mSQL = mSQL & " order by Tipo.Nombre"
            mSQL = mSQL & " order by cast(sum(DateDiff(""mi"", comienzo, fin))/60.00 as numeric(10,2)) desc"

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataAdapter.Fill(mDataset, "Tipo")
            mDataset.DataSetName = "TotTipo"
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try

    End Function

    Public Function Lista_Totales_Proyecto_Codigo_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean, Optional ByVal pID_Usuario As String = "") As String Implements ParteTrabajoDAO.Lista_Totales_Proyecto_Codigo_Xml
        Dim mSQL As String
        Dim mSQL2 As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por Codigo_Tarea
            mSQL = "SELECT Tarea.ID_Proyecto, Proyecto.Nombre, Codigo_Tarea, "
            mSQL = mSQL & " cast(sum(DateDiff(""mi"", comienzo, fin))/60.00 as numeric(10,2)),"
            mSQL = mSQL & " convert(nvarchar,min(comienzo),3) + ' ' + left(convert(nvarchar,min(comienzo),8),5) as Inicio,"
            mSQL = mSQL & " convert(nvarchar,max(fin),3) + ' ' + left(convert(nvarchar,max(fin),8),5) as Final, "
            mSQL = mSQL & " cast(Datediff(""mi"", min(comienzo),max(fin))/1440.00 as numeric(10,2)) as dias"
            If pID_Usuario <> "" Then
                mSQL = mSQL & ",Tarea.ID_Usuario"
            End If
            mSQL = mSQL & " FROM Tarea left join Proyecto "
            mSQL = mSQL & " on tarea.ID_Proyecto = Proyecto.ID_Proyecto"
            mSQL = mSQL & " where Tarea.fin Is Not null"
            If pID_Proyecto <> "" Then
                mSQL = mSQL & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pID_Usuario <> "" Then
                mSQL = mSQL & " and Tarea.ID_Usuario in ('" & pID_Usuario & "')"
            End If
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If

            'condiciones de fecha
            If Not IsDBNull(pFecha_Desde) Then
                mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            End If
            If Not IsDBNull(pFecha_Hasta) Then
                mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            End If

            mSQL = mSQL & " group by Tarea.ID_Proyecto, Proyecto.Nombre, Codigo_Tarea"
            If pID_Usuario <> "" Then
                mSQL = mSQL & ",Tarea.ID_Usuario"
            End If
            'mSQL = mSQL & " order by cast(sum(DateDiff(""mi"", comienzo, fin))/60.00 as numeric(10,2)) desc"
            mSQL = mSQL & " order by Proyecto.Nombre, Codigo_Tarea"

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataAdapter.Fill(mDataset, "CodigoTarea")
            mDataset.DataSetName = "TotCodigoTarea"
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Lista_Totales_Proyecto_Usuario_Tipo_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Totales_Proyecto_Usuario_Tipo_Xml
        Dim mSQL As String
        Dim mSQL2 As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'totales por usuario y tipo
            mSQL = "SELECT Usuario.Nombre, Usuario.Apellidos, Proyecto.Nombre, Tipo.Nombre, "
            mSQL = mSQL & " cast(sum(DateDiff(""mi"", comienzo, fin))/60.00 as numeric(10,2))"
            mSQL = mSQL & " FROM ((Tarea left join Tipo"
            mSQL = mSQL & " ON Tarea.ID_Tipo = Tipo.ID_Tipo"
            mSQL = mSQL & " ) LEFT JOIN Proyecto"
            mSQL = mSQL & " ON Tarea.ID_Proyecto = Proyecto.ID_Proyecto"
            mSQL = mSQL & " ) LEFT JOIN Usuario"
            mSQL = mSQL & " ON Tarea.ID_USuario = Usuario.ID_Usuario"
            mSQL = mSQL & " where(Tarea.fin Is Not null)"
            mSQL = mSQL & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If

            'condiciones de fecha
            If Not IsDBNull(pFecha_Desde) Then
                mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            End If
            If Not IsDBNull(pFecha_Hasta) Then
                mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            End If

            mSQL = mSQL & " group by Usuario.Nombre, Usuario.Apellidos, Proyecto.Nombre,  Tipo.Nombre"
            mSQL = mSQL & " order by usuario.apellidos, Tipo.Nombre"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "TotUsuarioTipo"
            mDataAdapter.Fill(mDataset, "UsuarioTipo")
            auxstr = mDataset.GetXml()

            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Lista_Tareas_Realizadas_Proyecto_Xml(ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pIncluirNoAprobadas As Boolean) As String Implements ParteTrabajoDAO.Lista_Tareas_Realizadas_Proyecto_Xml
        Dim mSQL As String
        Dim mSQL2 As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim auxstr As String

        Try

            'tareas
            mSQL = "SELECT ID_Tarea, Tarea.ID_Usuario, Usuario.Nombre, Usuario.Apellidos, Tipo.Nombre as Tipo, Codigo_Tarea, Comienzo , Fin, Descripcion, Tarea.Observaciones, cast(datediff(""mi"",comienzo,fin)/60.00 as numeric(10,2)) as Duracion"
            mSQL = mSQL & " FROM (Tarea LEFT JOIN Tipo ON Tarea.ID_Tipo = Tipo.ID_Tipo)"
            mSQL = mSQL & " LEFT JOIN Usuario ON Tarea.ID_Usuario = Usuario.ID_Usuario"
            mSQL = mSQL & " where(Tarea.Fin Is Not null)"
            If Not pIncluirNoAprobadas Then
                mSQL = mSQL & " and Tarea.Estado = 1"
            End If
            mSQL = mSQL & " and Tarea.ID_Proyecto = '" & pID_Proyecto & "'"

            'condiciones de fecha
            If Not IsDBNull(pFecha_Desde) Then
                mSQL = mSQL & " and Tarea.Comienzo >= '" & Format(pFecha_Desde, "dd/MM/yyyy") & " 00:00'"
            End If
            If Not IsDBNull(pFecha_Hasta) Then
                mSQL = mSQL & " and Tarea.Fin <= '" & Format(pFecha_Hasta, "dd/MM/yyyy") & " 23:59'"
            End If

            mSQL = mSQL & " order by comienzo"
            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)
            mDataset = New DataSet()
            mDataset.DataSetName = "Tareas"
            mDataAdapter.Fill(mDataset, "Tarea")
            auxstr = mDataset.GetXml()
            Return auxstr & vbNewLine

        Catch ex As Exception
            Trata_Error("", ex, mSQL)

        End Try


    End Function

    Public Function Inserta_Aprobacion(ByVal pID_Tarea As String, ByVal pID_Responsable As String, Optional ByVal pFecha As String = Nothing) As Boolean Implements ParteTrabajoDAO.Inserta_Aprobacion
        Dim mConsulta As String

        Try

            mConsulta = " INSERT INTO Aprobaciones([ID_Tarea], [ID_Responsable], [Fecha]) "
            If pFecha Is Nothing Then
                mConsulta = mConsulta & " VALUES(" & pID_Tarea & ",'" & pID_Responsable & "', dbo.mysysdate(getdate()))"
            Else
                mConsulta = mConsulta & " VALUES(" & pID_Tarea & ",'" & pID_Responsable & "', '" & pFecha & "')"
            End If


            Dim mCommand As New OleDb.OleDbCommand()

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try


    End Function

    '############################################3
    Public Function Ahora() As Date Implements ParteTrabajoDAO.Ahora
        'coge la fecha de la base de datos
        'a traves de la funcion mysysdate (da valores distintos para usuarios en canarias)
        Dim mConsulta As String

        Try

            Dim mSalida As DateTime

            mConsulta = "SELECT dbo.mysysdate(getdate())"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()
            mreader.Read()
            mSalida = NVL(mreader(0), Now)
            mreader.Close()

            Return mSalida

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Lista_Proyectos_Asignados_con_Tareas_Autorizadas(ByVal pID_Responsables As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_Proyectos_Asignados_con_Tareas_Autorizadas

        Dim mConsulta As String

        Try

            mConsulta = "SELECT distinct Proyecto.ID_Proyecto, Proyecto.Nombre"
            mConsulta = mConsulta & " FROM (Asignacion_Proyecto INNER JOIN Proyecto "
            mConsulta = mConsulta & " ON Asignacion_Proyecto.ID_Proyecto = Proyecto.ID_Proyecto) INNER JOIN TAREA"
            mConsulta = mConsulta & " ON Tarea.id_proyecto = Proyecto.id_proyecto"
            mConsulta = mConsulta & " WHERE Asignacion_Proyecto.ID_Responsable IN('" & pID_Responsables & "')"
            mConsulta = mConsulta & " and tarea.id_UltimoResponsable  = Asignacion_Proyecto.ID_Responsable"
            'mConsulta = mConsulta & " and (Estado is null or Estado = 0)  AND Fin is not null "
            mConsulta = mConsulta & " ORDER BY Proyecto.Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Tareas_Aprobadas_por_Usuario(ByVal pID_Responsables As String, ByVal pID_Proyecto As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Aprobadas_por_Usuario
        Dim mConsulta As String

        Try

            mConsulta = mConsulta & " SELECT Tarea.ID_Usuario, Tarea.ID_Proyecto, Sum(DATEDIFF(mi, Tarea.Comienzo, Tarea.Fin)) AS Duracion,"
            mConsulta = mConsulta & "(SELECT Usuario.Apellidos + ', ' + Usuario.Nombre AS Nombre FROM Usuario WHERE usuario.ID_Usuario = Tarea.id_usuario) as Nombre,"
            mConsulta = mConsulta & "(SELECT Proyecto.Nombre FROM Proyecto WHERE Proyecto.ID_Proyecto = Tarea.ID_Proyecto)AS Proyecto,"
            mConsulta = mConsulta & "(SELECT Cliente.Id_Cliente FROM Cliente,Proyecto WHERE Cliente.ID_Cliente = Proyecto.ID_Cliente and Proyecto.ID_Proyecto = Tarea.ID_Proyecto) AS Cliente,"
            mConsulta = mConsulta & "(SELECT Cliente.Nombre FROM Cliente,Proyecto WHERE Cliente.ID_Cliente = Proyecto.ID_Cliente and Proyecto.ID_Proyecto = Tarea.ID_Proyecto) AS NCliente"
            mConsulta = mConsulta & " FROM Tarea"
            mConsulta = mConsulta & " WHERE "
            'mConsulta = mConsulta & "(Estado is null or Estado = 0)  AND "
            mConsulta = mConsulta & " Fin is not null "
            If pID_Responsables <> "" Then
                mConsulta = mConsulta & "AND ID_UltimoResponsable IN ('" & pID_Responsables & "')"
            End If

            'filtros por fecha
            mConsulta = mConsulta & " AND Tarea.Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mConsulta = mConsulta & " AND Tarea.Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"

            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            mConsulta = mConsulta & " GROUP BY Tarea.ID_Usuario, Tarea.ID_Proyecto"
            mConsulta = mConsulta & " ORDER BY Nombre"

            Dim mCommand As New OleDb.OleDbCommand()
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Tareas_Aprobadas(ByVal pID_Responsable As String, ByVal pID_Usuario As String, ByVal pFecha_Desde As Date, ByVal pFecha_Hasta As Date, ByVal pID_Proyecto As String, ByVal pID_Tipo As String) As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Tareas_Aprobadas
        Dim mConsulta As String

        Try

            mConsulta = "SELECT ID_Tarea, "
            'mConsulta = mConsulta & " Tarea.ID_Usuario, Usuario.Apellidos + ', ' + Usuario.Nombre as Nombre,"
            'mConsulta = mConsulta & " Tarea.ID_Proyecto, Proyecto.Nombre as PNombre, Proyecto.ID_Cliente, Cliente.Nombre as CNombre, "
            'mConsulta = mConsulta & " Tarea.ID_Tipo, Tipo.Nombre as TNombre,"
            mConsulta = mConsulta & " Tarea.Codigo_Tarea,"
            mConsulta = mConsulta & " Tarea.Comienzo, Tarea.Fin, "
            mConsulta = mConsulta & " Tarea.Fin - Tarea.Comienzo as Duracion,"
            mConsulta = mConsulta & " Tarea.Descripcion"
            mConsulta = mConsulta & " FROM"
            mConsulta = mConsulta & " Tarea " ', Usuario, Proyecto, Cliente, Tipo"
            mConsulta = mConsulta & " WHERE " 'Tarea.ID_Usuario = Usuario.ID_Usuario"
            'mConsulta = mConsulta & " AND Tarea.ID_Proyecto = Proyecto.ID_Proyecto"
            'mConsulta = mConsulta & " AND Proyecto.ID_Cliente = Cliente.ID_Cliente"
            'mConsulta = mConsulta & " AND Tarea.ID_Tipo = Tipo.ID_Tipo"
            'mConsulta = mConsulta & " (Estado is null or Estado = 0)"
            'mConsulta = mConsulta & " AND Fin is not null"
            mConsulta = mConsulta & " ID_UltimoResponsable ='" & pID_Responsable & "'"
            If pID_Usuario <> "" Then
                mConsulta = mConsulta & " AND ID_Usuario ='" & pID_Usuario & "'"
            End If
            'filtros por fecha
            mConsulta = mConsulta & " AND Tarea.Comienzo >= '" & pFecha_Desde.ToString("dd/MM/yyyy") & " 00:00'"
            mConsulta = mConsulta & " AND Tarea.Fin <= '" & pFecha_Hasta.ToString("dd/MM/yyyy") & " 23:59'"
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Proyecto = '" & pID_Proyecto & "'"
            End If
            If pID_Proyecto <> "" Then
                mConsulta = mConsulta & " AND Tarea.ID_Tipo = " & pID_Tipo
            End If
            mConsulta = mConsulta & " ORDER BY Tarea.ID_Usuario, Comienzo"

            Dim mCommand As New OleDb.OleDbCommand
            Dim mreader As OleDb.OleDbDataReader

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mreader = mCommand.ExecuteReader()

            Return mreader

        Catch ex As Exception

            Trata_Error("", ex, mConsulta)

        End Try

    End Function

    Public Function Borra_Aprobacion(ByVal pLista_ID_Tarea As String, ByVal pID_Responsable As String) As Boolean Implements ParteTrabajoDAO.Borra_Aprobacion
        Dim mConsulta As String

        Try

            mConsulta = " DELETE Aprobaciones "
            mConsulta = mConsulta & " WHERE ID_Tarea IN (" & pLista_ID_Tarea & ")"
            mConsulta = mConsulta & " AND ID_Responsable = '" & pID_Responsable & "'"

            Dim mCommand As New OleDb.OleDbCommand

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try


    End Function


    Public Function Lista_AprobacionesTarea(Optional ByVal pID_Tarea As Long = -1, Optional ByVal pID_Responsable As String = "") As System.Data.OleDb.OleDbDataReader Implements ParteTrabajoDAO.Lista_AprobacionesTarea
        'da lista de aprobaciones
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As OleDb.OleDbDataReader

        Try

            mConsulta = "SELECT *"
            mConsulta = mConsulta & " FROM aprobaciones "
            If pID_Tarea >= 0 Then
                mWhere = "WHERE ID_Tarea = " & pID_Tarea
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " id_siguienteresponsable = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & " " & mWhere & " order by ID_Tarea, FECHA "

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("", ex, mConsulta)

        End Try

    End Function


    Public Function Lista_Proyecto_Asignaciones(ByVal pDatos As System.Data.DataSet, ByVal pID_Proyecto As String, Optional ByVal ID_Tarea As String = Nothing, Optional ByVal pID_Usuario As String = Nothing, Optional ByVal pOrden As String = Nothing, Optional ByRef pError As String = Nothing) As Boolean Implements ParteTrabajoDAO.Lista_Proyecto_Asignaciones

        Dim mSQL As String
        Dim mWHERE As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim mXML As String
        Dim auxstr As String

        Try

            mSQL = "SELECT * "
            mSQL = mSQL & " FROM Proyecto_Tarea_Asignacion "

            mWHERE = "WHERE ID_PROYECTO = '" & pID_Proyecto & "'"
            If ID_Tarea <> "" Then
                mWHERE = mWHERE & " AND ID_TAREA = '" & ID_Tarea & "'"
            End If
            If pID_Usuario <> "" Then
                mWHERE = mWHERE & " AND ID_Usuario = '" & pID_Usuario & "'"
            End If
            mSQL = mSQL & mWHERE
            If pOrden <> "" Then
                mSQL = mSQL & " ORDER BY " & pOrden
            End If

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)

            mDataset = New DataSet
            mDataAdapter.Fill(mDataset, "Proyecto_Tarea_Asignaciones")
            mDataset.DataSetName = "Proyecto_Tarea_Asignaciones"
            auxstr = mDataset.GetXml()
            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mSQL)
            Return False
        End Try


    End Function

    Public Function Lista_Proyecto_Tareas(ByRef pDatos As System.Data.DataSet, ByVal pID_Proyecto As String, Optional ByVal ID_Tarea As String = Nothing, Optional ByVal pOrden As String = Nothing, Optional ByRef pError As String = Nothing) As Boolean Implements ParteTrabajoDAO.Lista_Proyecto_Tareas
        Dim mSQL As String
        Dim mWHERE As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim mXML As String
        Dim auxstr As String

        Try

            mSQL = "SELECT * "
            mSQL = mSQL & " FROM Proyecto_Tarea"

            mWHERE = " WHERE ID_PROYECTO = '" & pID_Proyecto & "'"
            If ID_Tarea <> "" Then
                mWHERE = mWHERE & " AND ID_TAREA = '" & ID_Tarea & "'"
            End If
            mSQL = mSQL & mWHERE
            If pOrden <> "" Then
                mSQL = mSQL & " ORDER BY " & pOrden
            End If

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)

            pDatos = New DataSet
            mDataAdapter.Fill(pDatos, "Proyecto_Tarea")
            pDatos.DataSetName = "Proyecto_Tarea"

            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_Proyecto_Tareas(ByVal pID_Proyecto As String, ByVal ID_Tarea As String, ByVal pNombre As String, Optional ByVal pDescripcion As String = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pFechaInicio As String = Nothing, Optional ByVal pFechaFin As String = Nothing, Optional ByVal pHorasTrabajo As String = Nothing, Optional ByRef pError As String = "") As Boolean Implements ParteTrabajoDAO.Inserta_Proyecto_Tareas
        Dim mConsulta As String
        Try
            mConsulta = " INSERT INTO Proyecto_Tarea(ID_Proyecto, ID_Tarea, Nombre, Descripcion, Observaciones, Fecha_Inicio, Fecha_Fin, Horas_trabajo)"
            mConsulta = mConsulta & " VALUES('" & pID_Proyecto & "','" & ID_Tarea & "', "

            If pNombre <> "" Then
                mConsulta = mConsulta & "'" & QuitaComilla(pNombre) & "',"
            Else
                mConsulta = mConsulta & " NULL,"
            End If
            If pDescripcion <> "" Then
                mConsulta = mConsulta & "'" & QuitaComilla(pDescripcion) & "',"
            Else
                mConsulta = mConsulta & " NULL,"
            End If
            If pObservaciones <> "" Then
                mConsulta = mConsulta & "'" & QuitaComilla(pObservaciones) & "',"
            Else
                mConsulta = mConsulta & " NULL,"
            End If
            If IsDate(pFechaInicio) Then
                mConsulta = mConsulta & "'" & pFechaInicio & "',"
            Else
                mConsulta = mConsulta & " NULL,"
            End If
            If IsDate(pFechaFin) Then
                mConsulta = mConsulta & "'" & pFechaFin & "',"
            Else
                mConsulta = mConsulta & " NULL,"
            End If
            If IsNumeric(pHorasTrabajo) Then
                mConsulta = mConsulta & pHorasTrabajo & ")"
            Else
                mConsulta = mConsulta & " NULL)"
            End If

            Dim mCommand As New OleDb.OleDbCommand

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_Proyecto_Tareas(ByVal pID_Proyecto As String, ByVal pID_Tarea As String, ByVal pNombre As String, Optional ByVal pDescripcion As String = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pFechaInicio As String = Nothing, Optional ByVal pFechaFin As String = Nothing, Optional ByVal pHorasTrabajo As String = Nothing, Optional ByRef pError As String = "") As Boolean Implements ParteTrabajoDAO.Actualiza_Proyecto_Tareas
        Dim mConsulta As String
        Dim mSET As String
        Dim mWHERE As String
        Try
            mConsulta = " UPDATE Proyecto_Tarea"
            mWHERE = " WHERE ID_Proyecto = '" & pID_Proyecto & "' AND ID_TAREA = '" & pID_Tarea & "'"
            If Not IsNothing(pNombre) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                If pNombre <> "" Then
                    mSET = mSET & "Nombre = '" & QuitaComilla(pNombre) & "'"
                Else
                    mSET = mSET & "Nombre = NULL"
                End If

            End If
            If Not IsNothing(pDescripcion) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                If pDescripcion <> "" Then
                    mSET = mSET & "Descripcion = '" & QuitaComilla(pDescripcion) & "'"
                Else
                    mSET = mSET & "Descripcion = NULL"
                End If
            End If
            If Not IsNothing(pObservaciones) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                If pObservaciones <> "" Then
                    mSET = mSET & "Observaciones = '" & QuitaComilla(pObservaciones) & "'"
                Else
                    mSET = mSET & "Observaciones = NULL"
                End If
            End If
            If IsDate(pFechaInicio) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                mSET = mSET & "Fecha_Inicio = '" & QuitaComilla(pFechaInicio) & "'"
            End If
            If IsDate(pFechaFin) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                mSET = mSET & "Fecha_Fin = '" & QuitaComilla(pFechaFin) & "'"
            End If
            If IsNumeric(pHorasTrabajo) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                mSET = mSET & "Horas_Trabajo= " & pHorasTrabajo
            End If

            Dim mCommand As New OleDb.OleDbCommand

            mConsulta = mConsulta & " " & mSET & " " & mWHERE
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Proyecto_Tarea(ByVal pID_Proyecto As String, ByVal ID_Tarea As String, Optional ByRef pError As String = "") As Boolean Implements ParteTrabajoDAO.Elimina_Proyecto_Tarea
        Dim mConsulta As String
        Try
            mConsulta = " DELETE Proyecto_Tarea"
            mConsulta = " WHERE ID_Proyecto = '" & pID_Proyecto & "' AND ID_TAREA = '" & ID_Tarea & "'"

            Dim mCommand As New OleDb.OleDbCommand

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            If (mCommand.ExecuteNonQuery() > 0) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_ResumenUsuarioProyecto(ByRef pDatos As System.Data.DataSet, Optional ByVal pListaUsuarios As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pAprobadas As Boolean = True, Optional ByRef pError As String = Nothing) As Boolean Implements ParteTrabajoDAO.Lista_ResumenUsuarioProyecto

        Dim mSQL As String
        Dim mDataAdapter As OleDb.OleDbDataAdapter
        Dim mDataset As DataSet
        Dim i As Integer

        Try

            mSQL = "SELECT Tarea.[ID_Usuario], [ID_Proyecto], sum(DateDiff(mi, [Comienzo], [Fin])) as tiempo"
            mSQL = mSQL & " FROM Tarea,Usuario"
            mSQL = mSQL & " WHERE Usuario.id_usuario = tarea.id_usuario "
            If pFechaDesde <> "" Then
                mSQL = mSQL & " and Comienzo >= cast('" & pFechaDesde & " 00:00' as smalldatetime)"
            End If
            If pFechaHasta <> "" Then
                mSQL = mSQL & " and Comienzo <= cast('" & pFechaHasta & " 23:59' as smalldatetime)"
            End If
            If pAprobadas Then
                mSQL = mSQL & " and estado = 1"
            End If
            If pListaUsuarios <> "" Then
                mSQL = mSQL & " and Usuario.id_usuario in ('" & pListaUsuarios & "')"
            End If
            mSQL = mSQL & " group by Tarea.[ID_Usuario], [ID_Proyecto]"
            mSQL = mSQL & " order by [ID_Proyecto]"

            mDataAdapter = New OleDb.OleDbDataAdapter(mSQL, mConexion)

            pDatos = New DataSet
            mDataAdapter.Fill(pDatos, "ResumenUsuarioProyecto")
            pDatos.DataSetName = "ResumenUsuarioProyecto"

            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("", ex, mSQL)
            Return False
        End Try

    End Function



End Class

