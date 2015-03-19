Public Class OraclePresenciaDAO
    Inherits BaseDAO
    Implements PresenciaDAO


    'implementación para ORACLE
    Public Sub New(ByVal CadenaConexion As String, Optional ByVal pNivelLog As Integer = 0, Optional ByVal pFileLog As String = "OraclePresenciaDAO.Log")
        MyBase.New(CadenaConexion, pNivelLog, pFileLog)
    End Sub

    Public Function Conecta() As Boolean Implements PresenciaDAO.Conecta
        'conecta a la base de datos

        Return MyBase.ConectaDAO()

    End Function

    Public Function DesConecta() As Boolean Implements PresenciaDAO.DesConecta
        'desconecta de la base de datos

        Return MyBase.DesConectaDAO()

    End Function

    '######################################

    Public Function ObtenerDni(Optional ByVal pUsuario As String = "") As Data.DataSet Implements DAO.PresenciaDAO.ObtenerDNI

        Dim mConsulta As String

        Try

            mConsulta = "select DNI from empleados where email = '" & pUsuario & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en CargaCampo", ex, mConsulta)
        End Try
    End Function

    Public Function Elimina_PicadaUsuario(Optional ByVal DNI As String = "", Optional ByVal pan_tarjeta As String = "") As Data.DataSet Implements PresenciaDAO.Elimina_PicadaUsuario

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "Delete FROM eventos where "
            mConsulta = mConsulta & "dni_empl='" & DNI & "' and pan_tarjeta ='" & pan_tarjeta & "' and fecha=to_char(sysdate,'dd/mm/yyyy') and hora= to_char(sysdate,'hh24:mi') and permitido='S'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Elimina_PicadaUsuario", ex, mConsulta)
        End Try

    End Function

    Public Function Comprueba_PermisoTarjeta(Optional ByVal pDni As String = "", Optional ByVal pTarjeta As String = "") As Data.DataSet Implements PresenciaDAO.Comprueba_PermisoTarjeta

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT cod_recurso FROM (select cod_recurso from aattarjetasautorizadas a, "
            mConsulta = mConsulta & "(select t.pan_tarjeta, e.calcula_Saldo, e.dni from empleados e, tarjetasasociadas t where e.dni = t.DNI_EMPL and t.FECHA_HORA_BAJA is null) b where a.pan_tarjeta = " & pTarjeta & " AND trim(translate(a.pan_tarjeta,'0123456789',' ')) is null  "
            mConsulta = mConsulta & " ) where cod_recurso in (1,2,3,4,5,6,7,8,9) order by cod_recurso,2"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
                        'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Comprueba_PermisoTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Comprueba_TarjetaUsuario(Optional ByVal DNI As String = "", Optional ByVal pan_tarjeta As String = "") As Data.DataSet Implements PresenciaDAO.Comprueba_TarjetaUsuario

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT * FROM eventos where "
            mConsulta = mConsulta & "dni_empl='" & DNI & "' and pan_tarjeta ='" & pan_tarjeta & "' and fecha=to_char(sysdate,'dd/mm/yyyy') and permitido='S'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Comprueba_TarjetaUsuario", ex, mConsulta)
        End Try

    End Function

    Public Function Comprueba_Permiso(Optional ByVal DNI As String = "", Optional ByVal tipo As String = "") As Data.DataSet Implements PresenciaDAO.Comprueba_Permiso

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT * FROM AATTARJETASAUTORIZADAS,tarjetasasociadas where "
            If tipo = "E" Then
                mConsulta = mConsulta & " tarjetasasociadas.dni_empl "
            ElseIf tipo = "P" Then
                mConsulta = mConsulta & " tarjetasasociadas.dni_prov "
            End If
            mConsulta = mConsulta & "='" & DNI & "' and AATTARJETASAUTORIZADAS.cod_recurso=98 and tarjetasasociadas.pan_tarjeta=AATTARJETASAUTORIZADAS.pan_tarjeta"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Comprueba_Permiso", ex, mConsulta)
        End Try

    End Function

    Public Function Inserta_PANVisitante(Optional ByVal DNI As String = "", Optional ByVal pPanTarjeta As String = "") As Boolean Implements PresenciaDAO.Inserta_PANVisitante

        Dim mSQL As String

        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "UPDATE VISITANTES SET PAN_TARJETA ='" & pPanTarjeta & "' WHERE DNI='" & DNI & "'"
            'Conectarse, insertar datos y desconectarse de la base de datos 
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Inserta_Visitantes", ex, mSQL)
        End Try

    End Function

    Public Function Inserta_Visita(Optional ByVal DNI_VISITANTE As String = "", Optional ByVal DNI_VISITADO As String = "") As Boolean Implements PresenciaDAO.Inserta_Visita

        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim Fecha_Hora As DateTime

        Try

            mSQL = "INSERT INTO VISITAS (DNI_VISITANTE, DNI_EMPL_VISITADO, FECHA, DNI_EMPL_OPERARIO,MOTIVO) "
            mSQL = mSQL & " VALUES ('" & DNI_VISITANTE & "', '" & DNI_VISITADO & "', to_date('" & Fecha_Hora.Now.ToString() & "','dd/mm/yyyy hh24:mi:ss'),'" & DNI_VISITADO & "','')"


            'Conectarse, insertar datos y desconectarse de la base de datos 
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Inserta_Visitantes", ex, mSQL)
        End Try

    End Function

    Public Function Inserta_Visitante(Optional ByVal DNI As String = "", Optional ByVal NOMBRE As String = "", Optional ByVal APE1 As String = "", Optional ByVal APE2 As String = "", Optional ByVal PAN_TARJETA As String = "", Optional ByVal EMPRESA As String = "") As Boolean Implements PresenciaDAO.Inserta_Visitante

        Dim mSQL As String

        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "INSERT INTO VISITANTES (DNI, NOMBRE, APE1, APE2, PAN_TARJETA, EMPRESA) "
            mSQL = mSQL & " VALUES ('" & DNI & "', '" & NOMBRE & "', '" & APE1 & "', '" & APE2 & "', " & PAN_TARJETA & ", '" & EMPRESA & "')"

          
            'Conectarse, insertar datos y desconectarse de la base de datos 
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Inserta_Visitantes", ex, mSQL)
        End Try

    End Function

    Public Function Actualiza_Visitante(Optional ByVal DNI As String = "", Optional ByVal NOMBRE As String = "", Optional ByVal APE1 As String = "", Optional ByVal APE2 As String = "", Optional ByVal PAN_TARJETA As String = "", Optional ByVal EMPRESA As String = "") As Boolean Implements PresenciaDAO.Actualiza_Visitante

        Dim mSQL As String

        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "Update VISITANTES SET NOMBRE='" & NOMBRE & "', APE1='" & APE1 & "', APE2='" & APE2 & "', PAN_TARJETA='" & PAN_TARJETA & "', EMPRESA='" & EMPRESA & "' WHERE DNI='" & DNI & "'"


            'Conectarse, insertar datos y desconectarse de la base de datos 
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Inserta_Visitantes", ex, mSQL)
        End Try

    End Function


    Public Function Comprueba_Visitante(Optional ByVal DNI As String = "") As Data.DataSet Implements PresenciaDAO.Comprueba_Visitante

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from Visitantes"

            mWhere = " SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & DNI & "%'))"
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitados_Filtros_Todos(Optional ByVal Filtro As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitados_Filtros_Todos

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from empleados"
            If Filtro <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(EMPRESA)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitados_Filtros_varios(Optional ByVal FiltroNombre As String = "", Optional ByVal FiltroApellidos As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitados_Filtros_varios

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from empleados"
            If FiltroNombre <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & FiltroNombre & "%'))"
            End If
            If FiltroApellidos <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & FiltroApellidos & "%')) OR SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & FiltroApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitados_Filtro(Optional ByVal FiltroDNI As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitados_Filtro
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from empleados"
            mWhere = " SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & FiltroDNI & "%'))"
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitantes_Filtros_Todos(Optional ByVal Filtro As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitantes_Filtros_Todos

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from visitantes"
            If Filtro <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
                mWhere = mWhere + " OR SUPR_ACCENT(UPPER(EMPRESA)) LIKE SUPR_ACCENT(UPPER('" & Filtro & "%'))"
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function



    Public Function Lista_Visitantes_Filtros_varios(Optional ByVal FiltroNombre As String = "", Optional ByVal FiltroApellidos As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitantes_Filtros_varios

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from visitantes"
            If FiltroNombre <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & FiltroNombre & "%'))"
            End If
            If FiltroApellidos <> "" Then
                mWhere = " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & FiltroApellidos & "%')) OR SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & FiltroApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitantes_Filtro(Optional ByVal FiltroDNI As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitantes_Filtro
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from visitantes"
            mWhere = " SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & FiltroDNI & "%'))"
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Visitantes(Optional ByVal FiltroDNI As String = "") As Data.DataSet Implements PresenciaDAO.Lista_Visitantes
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Empresa from visitantes"
            mWhere = " DNI > '5000000'"
            If FiltroDNI <> "" Then
                If InStr(FiltroDNI, ",") > 0 Then
                    mWhere = " dni in (" & FiltroDNI & ")"
                Else
                    If UCase(Left(FiltroDNI, 7)) = "SELECT " Then
                        mWhere = " dni in (" & FiltroDNI & ")"
                    Else
                        mWhere = " dni ='" & FiltroDNI & "'"
                    End If

                End If

            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere & " ORDER BY DNI ASC"
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Visitantes", ex, mConsulta)
        End Try

    End Function

    Function CompruebaNombreyClave(ByVal pDni As String, ByVal pClave As String) As Data.DataSet Implements DAO.PresenciaDAO.CompruebaNombreyClave
        Dim mConsulta As String

        Try

            mConsulta = "SELECT * FROM empleados where dni = '" & pDni & "' and clave_web='" & pClave & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Existe_DniUsuario", ex, mConsulta)
        End Try
    End Function

    Function CompruebaAcceso(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.CompruebaAcceso
        Dim mConsulta As String

        Try

            mConsulta = "SELECT clave_web FROM empleados where dni = '" & pDni & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Existe_DniUsuario", ex, mConsulta)
        End Try
    End Function

    Function UsuarioTieneClave(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.UsuarioTieneClave
        Dim mConsulta As String

        Try

            mConsulta = "SELECT clave_web FROM empleados where dni = '" & pDni & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Existe_DniUsuario", ex, mConsulta)
        End Try
    End Function

    Function Existe_Clave(ByVal pClave As String, ByVal pTipo As String) As Data.DataSet Implements DAO.PresenciaDAO.Existe_Clave
        Dim mConsulta As String

        Try
            If pTipo = "1" Then
                mConsulta = "SELECT * FROM empleados where clave_emp = '" & pClave & "'"
            End If
            If pTipo = "2" Then
                mConsulta = "SELECT * FROM empleados where to_number(clave_emp) = " & Val(pClave)
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Existe_DniUsuario", ex, mConsulta)
        End Try
    End Function

    Function Existe_DniUsuario(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.Existe_DniUsuario
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM EMPLEADOS_PROVEEDORES WHERE dni ='" & pDni & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Existe_DniUsuario", ex, mConsulta)
        End Try
    End Function

    Function TarjetaAsociada(ByVal pNumTarjeta As String) As Data.DataSet Implements DAO.PresenciaDAO.TarjetaAsociada
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM tarjetasasociadas where pan_tarjeta = '" & pNumTarjeta & "' and " & " fecha_hora_baja is null"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function TarjetaAsociada2(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.TarjetaAsociada2
        Dim mConsulta As String

        Try
            mConsulta = "SELECT Pan_tarjeta from tarjetasasociadas WHERE DNI_EMPL ='" & pDni & "' AND FECHA_HORA_BAJA IS NULL"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en TarjetaAsociada2", ex, mConsulta)
        End Try
    End Function

    Function TarjetaTemporal(ByVal pNumTarjeta As String) As Data.DataSet Implements DAO.PresenciaDAO.TarjetaTemporal
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM TEMPORALES WHERE PAN_TARJETA = '" & pNumTarjeta & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function TarjetaAlta(ByVal pNumTarjeta As String) As Data.DataSet Implements DAO.PresenciaDAO.TarjetaAlta
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM Tarjetas where pan_tarjeta = '" & pNumTarjeta & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function CompruebaUltimaActiva(ByVal pDni As String, ByVal pActivas As String) As Data.DataSet Implements DAO.PresenciaDAO.CompruebaUltimaActiva
        Dim mConsulta As String

        Try
            If pActivas = "1" Then
                mConsulta = "SELECT * FROM TARJETASASOCIADAS WHERE dni_empl='" & pDni & "' AND fecha_hora_baja is NULL"

            Else
                mConsulta = "SELECT * FROM TARJETASASOCIADAS WHERE dni_empl = '" & pDni & "' ORDER BY fecha_hora_alta desc"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en CompruebaUltimaActiva", ex, mConsulta)
        End Try
    End Function

    Function CompruebaUltimaActivaVisitantes(ByVal pDni As String, ByVal pActivas As String) As Data.DataSet Implements DAO.PresenciaDAO.CompruebaUltimaActivaVisitantes
        Dim mConsulta As String

        Try
            If pActivas = "1" Then
                mConsulta = "SELECT * FROM TARJETASASOCIADAS WHERE dni_vis='" & pDni & "' AND (fecha_hora_baja is NULL or fecha_hora_baja >= sysdate)"

            Else
                mConsulta = "SELECT * FROM TARJETASASOCIADAS WHERE dni_vis = '" & pDni & "' ORDER BY fecha_hora_alta desc"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en CompruebaUltimaActivaVisitantes", ex, mConsulta)
        End Try
    End Function

    Function DevuelveSecuenciaTarjetas(ByVal pPan As String, ByVal pDni As String, ByVal tipo As String, ByVal pNumTarjeta As String) As Data.DataSet Implements DAO.PresenciaDAO.DevuelveSecuenciaTarjetas
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM TARJETASASOCIADAS WHERE "

            If tipo = "1" Then
                mConsulta = mConsulta & "dni_empl ='" & pDni & "'"
            ElseIf tipo = "2" Then
                mConsulta = mConsulta & "dni_prov ='" & pDni & "'"
            Else
                mConsulta = mConsulta & "dni_vis ='" & pDni & "'"
            End If
            mConsulta = mConsulta & " AND SUBSTR(PAN_TARJETA,8,6)='" & pNumTarjeta & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en DevuelveSecuenciaTarjetas", ex, mConsulta)
        End Try
    End Function

    Function ExisteTarjeta(ByVal pPan As String, ByVal tipo As String) As Data.DataSet Implements DAO.PresenciaDAO.ExisteTarjeta
        Dim mConsulta As String

        Try
            If tipo = "1" Then
                mConsulta = "SELECT count(1) from tarjetas where pan_tarjeta = '" & pPan & "'"
            Else
                mConsulta = "SELECT count(1) from tarjetasasociadas where pan_tarjeta = '" & pPan & "'"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function NuevoNumero(ByVal tipo As String) As Data.DataSet Implements DAO.PresenciaDAO.NuevoNumero
        Dim mConsulta As String

        Try
            mConsulta = "SELECT MAX(SUBSTR(PAN_TARJETA,8,6)) AS NUMERO FROM TARJETAS WHERE SUBSTR(PAN_TARJETA,14,1) = '" & tipo & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function NuevoNumeroDIP() As Data.DataSet Implements DAO.PresenciaDAO.NuevoNumeroDIP
        Dim mConsulta As String

        Try
            mConsulta = "SELECT MAX(to_number(PAN_TARJETA)) AS NUMERO FROM TARJETAS WHERE LENGTH(pan_tarjeta) = 12"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try
    End Function

    Function NuevoClaveEmpleado() As Data.DataSet Implements DAO.PresenciaDAO.NuevoClaveEmpleado
        Dim mConsulta As String

        Try
            mConsulta = "SELECT max(clave_emp) FROM empleados"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try

    End Function

    Function InsertaTarjeta(ByVal pPan As String, ByVal pEstado As String, ByVal pTipo As String, ByVal pFechaCaducidad As String) As Boolean Implements DAO.PresenciaDAO.InsertaTarjeta
        Dim mConsulta As String
        Try
            mConsulta = "INSERT INTO TARJETAS (PAN_TARJETA,ESTADO,TIPO_TARJETA,FECHA_CADUCIDAD) VALUES ('" & pPan & "','" & pEstado & "'," & Val(pTipo) & ",to_date('" & pFechaCaducidad & "','dd/mm/yyyy'))"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function ActualizaTarjeta(ByVal pNumTarjeta As String, ByVal pEstado As String, ByVal pEstadoModificado As String) As Boolean Implements DAO.PresenciaDAO.ActualizaTarjeta
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE Tarjetas set estado = '" & pEstado & "',estadomodificado = '" & pEstadoModificado & "' where pan_tarjeta = '" & pNumTarjeta & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function ActualizaTarjetasAsociadas(ByVal pFechaHora As String, ByVal pDni As String) As Boolean Implements DAO.PresenciaDAO.ActualizaTarjetasAsociadas
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE TARJETASASOCIADAS SET FECHA_HORA_BAJA = to_date('" & pFechaHora & "','dd/mm/yyyy hh24:mi:ss') WHERE DNI_EMPL ='" & pDni & "' AND FECHA_HORA_BAJA IS NULL"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function ActualizaTarjetasAsociadasVisitantes(ByVal pFechaHora As String, ByVal pDni As String) As Boolean Implements DAO.PresenciaDAO.ActualizaTarjetasAsociadasVisitantes
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE TARJETASASOCIADAS SET FECHA_HORA_BAJA = to_date('" & pFechaHora & "','dd/mm/yyyy hh24:mi:ss') WHERE DNI_VIS ='" & pDni & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function EliminaTarjeta(ByVal pPan As String) As Boolean Implements PresenciaDAO.EliminaTarjeta
        Dim mConsulta As String
        Try
            mConsulta = "DELETE tarjetas WHERE pan_tarjeta='" & pPan & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function EliminaTarjetaAsociada(ByVal pPan As String) As Boolean Implements PresenciaDAO.EliminaTarjetaAsociada
        Dim mConsulta As String
        Try
            mConsulta = "DELETE tarjetasasociadas where pan_tarjeta='" & pPan & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function EliminaTarjetaTemporal(ByVal pPan As String) As Boolean Implements PresenciaDAO.EliminaTarjetaTemporal
        Dim mConsulta As String
        Try
            mConsulta = "DELETE temporales where pan_tarjeta = '" & pPan & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaTemporal", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function ActualizaEmpTarjeta(ByVal pNumTarjeta As String, ByVal pDni As String) As Boolean Implements PresenciaDAO.ActualizaEmpTarjeta
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE empleados SET pan_tarjeta ='" & pNumTarjeta & "' WHERE dni ='" & pDni & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function ActualizaAnulacion(ByVal pNumTarjeta As String, ByVal pFechaAnulacion As String) As Boolean Implements DAO.PresenciaDAO.ActualizaAnulacion
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE TARJETAS SET FECHA_CADUCIDAD = to_date('" & pFechaAnulacion & "','dd/mm/yyyy') WHERE PAN_TARJETA ='" & pNumTarjeta & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function InsertaTAsociada(ByVal pPan As String, ByVal pDni As String, ByVal pFechaHoraAlta As String, ByVal pFechaHoraBaja As String) As Boolean Implements DAO.PresenciaDAO.InsertaTAsociada
        Dim mConsulta As String
        Try
            If pFechaHoraBaja = "null" Then
                mConsulta = "INSERT INTO TARJETASASOCIADAS (PAN_TARJETA,DNI_EMPL,FECHA_HORA_ALTA) VALUES ('" & pPan & "','" & pDni & "',to_date('" & pFechaHoraAlta & "','dd/mm/yyyy hh24:mi:ss'))"
            Else
                mConsulta = "INSERT INTO TARJETASASOCIADAS (PAN_TARJETA,DNI_EMPL,FECHA_HORA_ALTA, FECHA_HORA_BAJA) VALUES ('" & pPan & "','" & pDni & "',to_date('" & pFechaHoraAlta & "','dd/mm/yyyy hh24:mi:ss'),to_date('" & pFechaHoraBaja & "','dd/mm/yyyy hh24:mi:ss'))"
            End If

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Function InsertaTAsociadaVis(ByVal pPan As String, ByVal pDni As String, ByVal pFechaHoraAlta As String, ByVal pFechaHoraBaja As String) As Boolean Implements DAO.PresenciaDAO.InsertaTAsociadaVis
        Dim mConsulta As String
        Try
            If pFechaHoraBaja = "null" Then
                mConsulta = "INSERT INTO TARJETASASOCIADAS (PAN_TARJETA,DNI_VIS,FECHA_HORA_ALTA) VALUES ('" & pPan & "','" & pDni & "',to_date('" & pFechaHoraAlta & "','dd/mm/yyyy hh24:mi:ss'))"
            Else
                mConsulta = "INSERT INTO TARJETASASOCIADAS (PAN_TARJETA,DNI_VIS,FECHA_HORA_ALTA, FECHA_HORA_BAJA) VALUES ('" & pPan & "','" & pDni & "',to_date('" & pFechaHoraAlta & "','dd/mm/yyyy hh24:mi:ss'),to_date('" & pFechaHoraBaja & "','dd/mm/yyyy hh24:mi:ss'))"
            End If

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EliminaTarjetaAsociada", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function MeteGruposConsultaHijos(ByVal pCodigo As String) As Data.DataSet Implements DAO.PresenciaDAO.MeteGruposConsultaHijos
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM gruposconsulta WHERE grupo_padre = " & pCodigo & " ORDER BY desc_grupo "

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MeteGruposConsultaHijos", ex, mConsulta)
        End Try

    End Function

    


    Public Function Lista_GruposConsulta() As Data.DataSet Implements DAO.PresenciaDAO.Lista_GruposConsulta
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM gruposconsulta WHERE grupo_padre is null order by desc_grupo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposConsulta", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_GruposTrabajo() As Data.DataSet Implements DAO.PresenciaDAO.Lista_GruposTrabajo
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM grupotrabajo order by desc_grupotrabajo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposTrabajo", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_GruposPrivilegio() As Data.DataSet Implements DAO.PresenciaDAO.Lista_GruposPrivilegio
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM gruposprivilegios order by desc_grupo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposPrivilegio", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_GruposConsultaPerteneceA(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.Lista_GruposConsultaPerteneceA
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM gruposconsulta t1,pertenecena t2 WHERE t1.cod_grupo = t2.cod_grupo AND"
            mConsulta = mConsulta & " DNI_EMPL = '" & pDni & "'"
            mConsulta = mConsulta & " ORDER BY desc_grupo"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposPrivilegio", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_AsociaGrupoTrabajo(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.Lista_AsociaGrupoTrabajo
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM asociausuariogrupotrabajo t1, grupotrabajo t2 WHERE t1.cod_grupotrabajo = t2.cod_grupotrabajo AND "
            mConsulta = mConsulta & " t1.dni_EMPL "
            mConsulta = mConsulta & "' and t1.fecha_desde <= to_date('SYSDATE','DD/MM/YYYY') AND " & "(t1.fecha_hasta IS NULL OR t1.fecha_hasta > to_date('SYSDATE','DD/MM/YYYY'))"
            mConsulta = mConsulta & " ORDER BY desc_grupotrabajo"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposPrivilegio", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_GruposPrivilegioPerteneceA(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.Lista_GruposPrivilegioPerteneceA
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM gruposprivilegios t1,pertenecena t2 WHERE t1.cod_grupo = t2.cod_grupo AND"
            mConsulta = mConsulta & " DNI_EMPL = '" & pDni & "'"
            mConsulta = mConsulta & " ORDER BY desc_grupo"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Lista_GruposPrivilegio", ex, mConsulta)
        End Try

    End Function



    '######################################
    Public Function Guardar_ValorScapini_Personal(ByVal pDni As String, ByVal pCampo As String, ByVal pValor As String, ByVal pTipo As String) As Boolean Implements DAO.PresenciaDAO.Guardar_ValorScapini_Personal

        Dim mConsulta As String

        Try

            If pTipo = "insertar" Then
                mConsulta = "INSERT INTO empleados(DNI) values('" & pDni & "')"
            End If
            If pTipo = "actualizar" Then
                mConsulta = "UPDATE empleados SET " & pCampo & " = '" & pValor & "'" & " where dni = '" & pDni & "'"
            End If

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Guarda_DatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Guardar_ValorScapini(ByVal pDni As String, ByVal pCampo As String, ByVal pValor As String, ByVal pTipo As String) As Boolean Implements DAO.PresenciaDAO.Guardar_ValorScapini

        Dim mConsulta As String

        Try

            If pTipo = "insertar" Then
                mConsulta = "INSERT INTO SCAPINI (DNI) values('" & pDni & "')"
            End If
            If pTipo = "actualizar" Then
                mConsulta = "UPDATE SCAPINI SET " & pCampo & " = '" & pValor & "'" & " where dni = '" & pDni & "'"
            End If

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Guarda_DatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function ValorScapini_Personal(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.ValorScapini_Personal
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM empleados where dni = '" & pDni & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en ValorScapini_Personal", ex, mConsulta)
        End Try

    End Function

    Public Function ValorScapini_Personal2(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.ValorScapini_Personal2
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM empleados where dni = '0'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en ValorScapini_Personal2", ex, mConsulta)
        End Try

    End Function

    Public Function ValorScapini1(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.ValorScapini1
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM SCAPINI where dni = '" & pDni & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en ValorScapini1", ex, mConsulta)
        End Try

    End Function

    Public Function ValorScapini2(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.ValorScapini2
        Dim mConsulta As String

        Try
            mConsulta = "SELECT * FROM SCAPINI where dni = '0'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en ValorScapini2", ex, mConsulta)
        End Try

    End Function

    Public Function Empleados_ProveedoresTipo(ByVal pDni As String) As Data.DataSet Implements DAO.PresenciaDAO.Empleados_ProveedoresTipo
        Dim mConsulta As String

        Try
            mConsulta = "SELECT tipo from empleados_proveedores where dni ='" & pDni & "'"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Empleados_ProveedoresTipo", ex, mConsulta)
        End Try

    End Function

    Public Function CargaCampo(ByVal pCPerfil As String, ByVal pCDato As String) As Data.DataSet Implements DAO.PresenciaDAO.CargaCampo

        Dim mConsulta As String
        Try
            mConsulta = "SELECT * from dato_perfil_tarjeta where cod_perfil = " & pCPerfil & " and cod_dato = " & pCDato

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en CargaCampo", ex, mConsulta)
        End Try
    End Function

    Public Function Comprueba_clave(Optional ByVal pClave As String = "", Optional ByVal pDni As String = "") As Data.DataSet Implements DAO.PresenciaDAO.Comprueba_clave

        Dim mConsulta As String

        Try

            mConsulta = "select * from empleados where clave_emp = " & pClave
            If pDni <> "" Then
                mConsulta = mConsulta & " and dni <> '" & pDni & "'"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en CargaCampo", ex, mConsulta)
        End Try
    End Function

    Public Function Obten_Valor(ByVal PCadSQL As String) As Data.DataSet Implements DAO.PresenciaDAO.Obten_Valor

        Dim mConsulta As String

        Try
            mConsulta = PCadSQL

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en Obten_Valor", ex, mConsulta)
        End Try

    End Function

    Public Function CargaPerfiles() As Data.DataSet Implements DAO.PresenciaDAO.CargaPerfiles

        Dim mConsulta As String

        Try
            mConsulta = "SELECT cod_perfil,nombre_perfil from perfil_tarjeta order by cod_perfil"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en CargaPerfiles", ex, mConsulta)
        End Try

    End Function

    Public Function VisualizaTarjeta(ByVal pCodigoPerfil As String) As Data.DataSet Implements DAO.PresenciaDAO.VisualizaTarjeta

        Dim mConsulta As String

        Try

            mConsulta = "select * from dato_perfil_tarjeta where cod_perfil = " & pCodigoPerfil & " order by cod_dato"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en VisualizaTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function VisualizaEmpleadoTarjeta() As Data.DataSet Implements DAO.PresenciaDAO.VisualizaEmpleadoTarjeta

        Dim mConsulta As String

        Try

            mConsulta = "Select fondo from perfil_tarjeta where cod_perfil = (select min(cod_perfil) from perfil_tarjeta)"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en VisualizaEmpleadoTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function MuestraHistorial(ByVal pDNI As String, ByVal tipo As Boolean) As Data.DataSet Implements DAO.PresenciaDAO.MuestraHistorial
        Dim mConsulta As String

        Try

            mConsulta = "SELECT PAN_TARJETA,to_char(FECHA_HORA_ALTA,'dd/mm/yyyy hh24:mi:ss') FECHA_ALTA, to_char(FECHA_HORA_BAJA,'dd/mm/yyyy hh24:mi:ss') FECHA_BAJA FROM TARJETASASOCIADAS WHERE DNI_EMPL = '" & pDNI & "' ORDER BY FECHA_HORA_BAJA ASC"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en MuestraHistorial", ex, mConsulta)
        End Try

    End Function

    Public Function VisualizarTarjeta(ByVal pCodigo As String) As Data.DataSet Implements DAO.PresenciaDAO.VisualizarTarjeta

        Dim mConsulta As String

        Try


            mConsulta = "Select * from perfil_tarjeta where cod_perfil = " & pCodigo

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en VisualizarTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Guarda_DatosPerfilTarjeta(ByVal pCPerfil As String, ByVal pCDato As String, ByVal pFuente As String, ByVal pNegrita As String, ByVal pFoto As String, ByVal pX As String, ByVal pY As String, ByVal pAlto As String, ByVal pAncho As String, ByVal pTamano As String, ByVal pColor As String, ByVal pSource As String) As Boolean Implements DAO.PresenciaDAO.Guarda_DatosPerfilTarjeta

        Dim mConsulta As String

        Try

            mConsulta = "update dato_perfil_tarjeta set x = " & pX & ", y = " & pY & ", alto = " & pAlto & ", ancho = " & pAncho & ", tamano = " & pTamano
            mConsulta = mConsulta & ", nombre_tipo = " & pFuente & ""
            mConsulta = mConsulta & ", negrita = " & pNegrita & ""
            mConsulta = mConsulta & ", foto = " & pFoto & ""
            mConsulta = mConsulta & ",color = " & pColor
            mConsulta = mConsulta & ", source = '" & pSource & "'"
            mConsulta = mConsulta & " where cod_perfil = " & pCPerfil & " and cod_dato = " & pCDato

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Guarda_DatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Guarda_DatosPerfilTarjeta_Coordenadas(ByVal pCPerfil As String, ByVal pCDato As String, ByVal pX As String, ByVal pY As String) As Boolean Implements DAO.PresenciaDAO.Guarda_DatosPerfilTarjeta_Coordenadas

        Dim mConsulta As String

        Try

            mConsulta = "update dato_perfil_tarjeta set x = " & pX & _
            ", y = " & pY & "where cod_perfil = " & pCPerfil & " and cod_dato = " & pCDato

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Guarda_DatosPerfilTarjeta_Coordenadas", ex, mConsulta)
        End Try

    End Function

    Public Function Muestra_DatosPerfilTarjeta(ByVal pCPerfil As String, ByVal pCDato As String) As Data.DataSet Implements DAO.PresenciaDAO.Muestra_DatosPerfilTarjeta

        Dim mConsulta As String
        Try

            mConsulta = "select * from dato_perfil_tarjeta where cod_perfil = " & pCPerfil & " and cod_dato = " & pCDato

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Muestra_DatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function ListaDatosPerfilTarjeta(ByVal pCodigo As String) As Data.DataSet Implements DAO.PresenciaDAO.ListaDatosPerfilTarjeta

        Dim mConsulta As String
        Try

            mConsulta = "select max(cod_dato) from dato_perfil_tarjeta where cod_perfil = " & pCodigo

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en ListaDatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function InsertaDatosPerfilTarjeta(ByVal pCodigo As String, ByVal pNuevoCodigo As String) As Boolean Implements DAO.PresenciaDAO.InsertaDatosPerfilTarjeta

        Dim mConsulta As String

        Try

            mConsulta = "INSERT into dato_perfil_tarjeta(cod_perfil,cod_dato) values(" & pCodigo & "," & pNuevoCodigo & ")"
            
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en InsertaDatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function CargaDatosPerfilTarjeta(ByVal pCodigo As Integer) As Data.DataSet Implements DAO.PresenciaDAO.CargaDatosPerfilTarjeta

        Dim mConsulta As String
        Try

            mConsulta = "select * from dato_perfil_tarjeta where cod_perfil = " & pCodigo & " order by cod_dato"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en CargaDatosPerfilTarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function ListaPerfilesTarjeta() As Data.DataSet Implements DAO.PresenciaDAO.ListaPerfilesTarjeta

        Dim mConsulta As String
        Try

            mConsulta = "select max(cod_perfil) from perfil_tarjeta"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet

        Catch ex As Exception
            Trata_Error("Error en ListaPerfilesTarjeta", ex, mConsulta)
        End Try


    End Function

    Public Function Elimina_DatosPerfilTarjeta(ByVal pCodigo As String, ByVal pDato As String) As Boolean Implements DAO.PresenciaDAO.Elimina_DatosPerfilTarjeta

        Dim mConsulta As String

        Try

            mConsulta = "Delete dato_perfil_tarjeta where cod_perfil = " & pCodigo & " and cod_dato = " & pDato

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Elimina_DatosPerfilTarjeta", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Elimina_Perfil(ByVal pCodigo As String) As Boolean Implements DAO.PresenciaDAO.Elimina_Perfil

        Dim mConsulta As String

        Try

            mConsulta = "Delete perfil_tarjeta where cod_perfil = " & pCodigo

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Elimina_Perfil", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Muestra_Perfil(ByVal pCodigo As String) As Data.DataSet Implements DAO.PresenciaDAO.Muestra_Perfil

        Dim mConsulta As String
        Try

            mConsulta = "select * from perfil_tarjeta where cod_perfil = " & pCodigo

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Muestra_Perfil", ex, mConsulta)
        End Try

    End Function

    Public Function Guardar_Perfil(ByVal pCodigo As Integer, ByVal pNombrePerfil As String, ByVal pFondo As String) As Boolean Implements DAO.PresenciaDAO.Guardar_Perfil

        Dim mConsulta As String

        Try

            mConsulta = "update perfil_tarjeta set nombre_perfil = '" & pNombrePerfil & "'"
            mConsulta = mConsulta & ", fondo = '" & pFondo & "'"
            mConsulta = mConsulta & " where cod_perfil = " & pCodigo

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Guardar_Perfil", ex, mConsulta)
        End Try

    End Function

    Public Function Insertar_Perfil(ByVal pCodigo As String, ByVal pNombrePerfil As String) As Boolean Implements DAO.PresenciaDAO.Insertar_Perfil

        Dim mConsulta As String

        Try

            mConsulta = "insert into perfil_tarjeta(cod_perfil,nombre_perfil) values(" & pCodigo & ",'" & pNombrePerfil & "')"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

        Catch ex As Exception
            Trata_Error("Error en Insertar_Perfil", ex, mConsulta)
        End Try

    End Function


    '######################################
    Public Function Lista_Empleados(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Object Implements PresenciaDAO.Lista_Empleados
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            'mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, Admin from empleados"
            mConsulta = "SELECT * from empleados"
            If FiltroDNI <> "" Then
                If InStr(FiltroDNI, ",") > 0 Then
                    mWhere = " dni in (" & FiltroDNI & ")"
                Else
                    If UCase(Left(FiltroDNI, 7)) = "SELECT " Then
                        mWhere = " dni in (" & FiltroDNI & ")"
                    Else
                        mWhere = " dni ='" & FiltroDNI & "'"
                    End If

                End If

            End If
            If FiltroEmail <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " and upper(email) ='" & UCase(FiltroEmail) & "'"
                Else
                    mWhere = " upper(email) ='" & UCase(FiltroEmail) & "'"
                End If
            End If
            If pUsuariosSinResponsable = True Then
                'solo usuarios sin responsable
                If mWhere <> "" Then
                    mWhere = mWhere & " and "
                End If
                mWhere = mWhere & " dni not in (select id_usuario from asignacion_responsable)"
                mWhere = mWhere & " and calcula_saldo = 'S'"
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere
            End If
            If Orden <> "" Then
                mConsulta = mConsulta & " ORDER BY " & Orden
            End If
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Empleados2(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Data.DataSet Implements PresenciaDAO.Lista_Empleados2
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            'mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, Admin from empleados"
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, telefono, centro, cargo, Empresa, calcula_saldo, permite_visita, clave_emp from empleados"
            If FiltroDNI <> "" Then
                If InStr(FiltroDNI, ",") > 0 Then
                    mWhere = " dni in (" & FiltroDNI & ")"
                Else
                    If UCase(Left(FiltroDNI, 7)) = "SELECT " Then
                        mWhere = " dni in (" & FiltroDNI & ")"
                    Else
                        mWhere = " dni ='" & FiltroDNI & "'"
                    End If

                End If

            End If
            If FiltroEmail <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " and upper(email) ='" & UCase(FiltroEmail) & "'"
                Else
                    mWhere = " upper(email) ='" & UCase(FiltroEmail) & "'"
                End If
            End If
            If pUsuariosSinResponsable = True Then
                'solo usuarios sin responsable
                If mWhere <> "" Then
                    mWhere = mWhere & " and "
                End If
                mWhere = mWhere & " dni not in (select id_usuario from asignacion_responsable)"
                mWhere = mWhere & " and calcula_saldo = 'S'"
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere
            End If
            If Orden <> "" Then
                mConsulta = mConsulta & " ORDER BY " & Orden
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function ListaEmpleadosDNI(Optional ByVal pDNI As String = "") As Data.DataSet Implements PresenciaDAO.ListaEmpleadosDNI
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            'mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, Admin from empleados"
            mConsulta = "SELECT CLAVE,EMPRESA from empleados WHERE DNI=" + pDNI

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function ListaProveedoresDNI(Optional ByVal pDNI As String = "") As Data.DataSet Implements PresenciaDAO.ListaProveedoresDNI
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            'mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, Admin from empleados"
            mConsulta = "SELECT DNI,EMPRESA from empleados WHERE DNI=" + pDNI

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Lista_Proveedores", ex, mConsulta)
        End Try

    End Function

    Public Function Listado_Empleados(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Data.DataSet Implements PresenciaDAO.Listado_Empleados
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            'mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, Admin from empleados"
            mConsulta = "SELECT * FROM EMPLEADOS ORDER BY APE1, APE2, NOMBRE"

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Listado_Empleados", ex, mConsulta)
        End Try

    End Function

    '######################################
    Public Function Registra_LOG_Visita(ByVal pUsuario As String, ByVal pObserva As String) As Boolean Implements PresenciaDAO.Registra_LOG_Visita
        'inserta un registro en la tabla de log de visitas, usa la fecha hora de la BD

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand

        Try

            mCommand.Connection = mConexion
            'pone los datos del dia seleccionado
            mConsulta = "INSERT INTO LOG_VISITAS(usuario,fecha, observaciones)"
            mConsulta = mConsulta & " values('" & pUsuario & "',SYSDATE,'" & Left(pObserva, 200) & "')"
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    '######################################
    Public Function Lista_Eventos(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object Implements PresenciaDAO.Lista_Eventos
        'devuelve la lista de eventos

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mCommand.Connection = mConexion
            mConsulta = "select e_s sentido, hora, eventos.cod_recurso, recursos.desc_recurso, eventos.cod_evento, eventos.cod_incidencia,  eventos.dni_empl "
            mConsulta = mConsulta & " from recursos, eventos"
            mConsulta = mConsulta & " where recursos.cod_recurso (+) = eventos.cod_recurso"
            If pUsuario <> "" Then
                mConsulta = mConsulta & " and dni_empl = '" & pUsuario & "'"
            End If
            mConsulta = mConsulta & " and fecha >= to_date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " and fecha <= to_date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " order by hora, e_s"
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Eventos", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_EventosGrupos(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object Implements PresenciaDAO.Lista_EventosGrupos
        'devuelve la lista de eventos

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mCommand.Connection = mConexion
            mConsulta = "select e_s sentido, hora, eventos.cod_recurso, recursos.desc_recurso, eventos.cod_evento, eventos.cod_incidencia, eventos.dni_empl "
            mConsulta = mConsulta & " from recursos, eventos"
            mConsulta = mConsulta & " where recursos.cod_recurso (+) = eventos.cod_recurso"
            If pUsuario <> "" Then
                mConsulta = mConsulta & " and dni_empl in ( select dni_empl from pertenecena where cod_grupo in (" & pUsuario & "))"
            End If
            mConsulta = mConsulta & " and fecha >= to_date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " and fecha <= to_date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " order by hora, e_s"
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Eventos", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Primer_Evento(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object Implements PresenciaDAO.Lista_Primer_Evento
        'devuelve el primer evento

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mCommand.Connection = mConexion
            mConsulta = "select e_s sentido, hora, eventos.cod_recurso, recursos.desc_recurso, eventos.cod_evento, eventos.cod_incidencia "
            mConsulta = mConsulta & " from recursos, eventos"
            mConsulta = mConsulta & " where recursos.cod_recurso = eventos.cod_recurso"
            If pUsuario <> "" Then
                mConsulta = mConsulta & " and dni_empl = '" & pUsuario & "'"
            End If
            mConsulta = mConsulta & " and fecha >= to_date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " and fecha <= to_date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " order by hora, e_s"
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Eventos", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Ultimo_Evento(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object Implements PresenciaDAO.Lista_Ultimo_Evento
        'devuelve el ultimo evento

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mCommand.Connection = mConexion
            mConsulta = "select e_s sentido, hora, eventos.cod_recurso, recursos.desc_recurso, eventos.cod_evento, eventos.cod_incidencia "
            mConsulta = mConsulta & " from recursos, eventos"
            mConsulta = mConsulta & " where recursos.cod_recurso = eventos.cod_recurso"
            If pUsuario <> "" Then
                mConsulta = mConsulta & " and dni_empl = '" & pUsuario & "'"
            End If
            mConsulta = mConsulta & " and fecha >= to_date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " and fecha <= to_date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " order by hora desc, e_s"
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Eventos", ex, mConsulta)
        End Try

    End Function

    '######################################
    Public Function Lista_Horarios(Optional ByVal pCod_Horario As String = "") As Object Implements PresenciaDAO.Lista_Horarios
        'da la lista de horarios

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mCommand.Connection = mConexion
            mConsulta = "SELECT horarios.cod_horario, desc_horario, nvl(sum(total_minutos),0) duracion, min(intervalos.hora_inicio) inicio, max(intervalos.hora_inicio + intervalos.TOTAL_MINUTOS) fin from horarios, intervalos"
            mConsulta = mConsulta & " where horarios.cod_horario = intervalos.cod_horario(+)"
            If pCod_Horario <> "" Then
                mConsulta = mConsulta & " and horarios.cod_horario = " & pCod_Horario
            End If
            mConsulta = mConsulta & " group by horarios.cod_horario, horarios.desc_horario "
            mConsulta = mConsulta & " order by horarios.desc_horario "

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Horarios", ex, mConsulta)
        End Try

    End Function

    '######################################
    Public Function Datos_Diario(ByVal pDNI As String, ByVal pFechaDesde As Date, ByVal pFechaHasta As Date) As Object Implements PresenciaDAO.Datos_Diario
        'da los datos del diario

        'comprueba el diario
        Dim mfecha As Date
        Dim i As Integer
        Dim mConsulta As String
        Try
            mfecha = pFechaDesde
            For i = 0 To DateDiff(DateInterval.Day, pFechaDesde, pFechaHasta)
                mfecha = DateAdd(DateInterval.Day, i, pFechaDesde)
                Comprueba_Diario(pDNI, mfecha)
            Next


            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object

            mCommand.Connection = mConexion
            mConsulta = "SELECT Presencia, Justificaciones, Saldo, Cod_Horario from Diario"
            mConsulta = mConsulta & " WHERE dni = '" & pDNI & "'"
            mConsulta = mConsulta & "  and fecha >= to_Date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & "  and fecha <= to_Date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & "  order by fecha"

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Datos_Diario", ex, mConsulta)
        End Try

    End Function


    Public Sub Calcula_Diario(ByVal pDNI As String, ByVal pFecha As Date) Implements PresenciaDAO.Calcula_Diario
        Comprueba_Diario(pDNI, pFecha)
    End Sub

    Private Sub Comprueba_Diario(ByVal pDNI As String, ByVal pFecha As Date)

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        mCommand.Connection = mConexion

        'consulta los datos del usuario y dia
        mSQL = "SELECT Presencia,Justificaciones,Saldo,Cod_Horario from Diario"
        mSQL = mSQL & " WHERE dni = '" & pDNI & "' and fecha = '" & Format(pFecha, "dd/MM/yyyy") & "'"
        mCommand.CommandText = mSQL
        mReader = mCommand.ExecuteReader()
        If mReader.Read Then
            mReader.Close()
            'comprobar que no hay picadas sin procesar
            mSQL = "SELECT count(*) FROM Eventos"
            mSQL = mSQL & " WHERE dni_empl = '" & pDNI & "' and fecha = '" & Format(pFecha, "dd/MM/yyyy") & "'"
            mSQL = mSQL & " AND (marcado is null or marcado = 'N')"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read() Then
                If mReader(0) > 0 Then
                    mReader.Close()
                    Actualiza_Saldo(pDNI, pFecha)
                    Exit Sub
                End If
            End If
            mReader.Close()
            'comprobar que no hay picadas sin procesar
            mSQL = "SELECT count(*) FROM Justificaciones"
            mSQL = mSQL & " WHERE dni_empl = '" & pDNI & "' and fecha_justificada = '" & Format(pFecha, "dd/MM/yyyy") & "'"
            mSQL = mSQL & " AND (CONTABILIZADA is null or CONTABILIZADA = 'N')"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read() Then
                If mReader(0) > 0 Then
                    mReader.Close()
                    Actualiza_Saldo(pDNI, pFecha)
                    Exit Sub
                End If
            End If
            mReader.Close()
        Else
            mReader.Close()
            Actualiza_Saldo(pDNI, pFecha)
        End If

    End Sub

    Public Function Actualiza_Saldo(ByVal pDNI As String, ByVal pFECHA As Date) As Boolean Implements PresenciaDAO.Actualiza_Saldo
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object


        mCommand.Connection = mConexion

        mConsulta = "{CALL saldo('" & pDNI & "',to_Date('" & Format(pFECHA, "dd/MM/yyyy") & "','DD/MM/YYYY')) }"
        mCommand.CommandText = mConsulta
        mCommand.ExecuteNonQuery()

    End Function

    '######################################
    Public Function Lista_Incidencias(Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As Boolean = False) As Object Implements PresenciaDAO.Lista_Incidencias
        'da la lista de incidencias

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try

            mCommand.Connection = mConexion
            mConsulta = "SELECT cod_incidencia, "
            mConsulta &= " nvl(desc_incidencia, 'Sin Desc.') desc_incidencia,"
            mConsulta &= " nvl(maximo, 0) maximo,"
            mConsulta &= " nvl(fecha_base, '01/01') fecha_base,"
            mConsulta &= " tipo, "
            mConsulta &= " nvl(tiempo_maximo, 0) tiempo_maximo,"
            mConsulta &= " nvl(fecha_termino, '31/12') fecha_termino,"
            mConsulta &= " nvl(orden, 0) orden,"
            mConsulta &= " nvl(tipofijo, 'N') tipofijo,"
            mConsulta &= " seleccionable, nvl(maximo_horas,0) maximo_horas,"
            'mConsulta &= " PL_MAXIMO, PL_SOLICITAR, PL_DEL_SOL, PL_JUSTIFICAR, PL_DEL_JUST,PL_MAXIMO_HORAS"
            mConsulta &= " PL_MAXIMO, PL_SOLICITAR, PL_MAXIMO_HORAS"
            mConsulta &= " from incidencias "
            If pCod_Incidencia >= 0 Then
                mConsulta = mConsulta & " WHERE Cod_incidencia = " & pCod_Incidencia
            End If
            If pSeleccionable Then
                mConsulta = mConsulta & " WHERE Seleccionable = 'S'"
            End If
            If pOrden <> "" Then
                mConsulta = mConsulta & " order by " & pOrden
            Else
                mConsulta = mConsulta & " order by cod_incidencia"
            End If

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Incidencias", ex, mConsulta)
        End Try

    End Function


    Public Function Cadena_de_Picadas(ByVal pDNI As String, ByVal pFecha As Date) As String Implements PresenciaDAO.Cadena_de_Picadas
        'da la lista de picadas

        Dim mLista_Presencia As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try

            mCommand.Connection = mConexion
            mSQL = "SELECT picadas_usuario('" & pDNI & "',to_date('" & Format(pFecha, "dd/MM/yyyy") & "','DD/MM/YYYY')) "
            mSQL = mSQL & " FROM dual "
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                mLista_Presencia = Convierte_Cadena_Picadas(NVL(mReader(0), ""))
            End If
            mReader.Close()

            Return mLista_Presencia

        Catch ex As Exception
            Trata_Error("Error en Cadena_de_Picadas", ex, mSQL)
        End Try

    End Function

    Public Function Cadena_de_Justificaciones(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pTipo As String = "") As String Implements PresenciaDAO.Cadena_de_Justificaciones
        'da la lista de picadas

        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try
            If pTipo = "" Then pTipo = " "
            mCommand.Connection = mConexion
            mSQL = "SELECT intervalos_justificacionesTipo('" & pDNI & "',to_date('" & Format(pFecha, "dd/MM/yyyy") & "','DD/MM/YYYY'),'" & pTipo & "') "
            mSQL = mSQL & " FROM dual "
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                mLista = Convierte_Cadena_Justificaciones(NVL(mReader(0), ""))
            End If
            mReader.Close()

            Return mLista

        Catch ex As Exception
            Trata_Error("Error Cadena_de_Justificaciones", ex, mSQL)
        End Try

    End Function

    Public Function Cadena_de_Solicitudes(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pTipo As String = "") As String Implements PresenciaDAO.Cadena_de_Solicitudes
        'da la lista de picadas

        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try
            If pTipo = "" Then pTipo = " "
            mCommand.Connection = mConexion
            mSQL = "SELECT intervalos_solicitudesTipo('" & pDNI & "',to_date('" & Format(pFecha, "dd/MM/yyyy") & "','DD/MM/YYYY'),'" & pTipo & "') "
            mSQL = mSQL & " FROM dual "
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                mLista = Convierte_Cadena_Justificaciones(NVL(mReader(0), ""))
            End If
            mReader.Close()

            Return mLista

        Catch ex As Exception
            Trata_Error("Error Cadena_de_Solicitudes", ex, mSQL)
        End Try

    End Function

    Public Function Cadena_de_Intervalos(ByVal pcod_horario As Integer, Optional ByVal pFormato As String = "N") As String Implements PresenciaDAO.Cadena_de_Intervalos
        ' pFormato: H:Horas ("09:34"), ó N:Numérico (546)

        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try
            mCommand.Connection = mConexion
            If pFormato = "N" Then
                mSQL = "SELECT hora_inicio, hora_inicio + total_minutos FROM Intervalos"
            Else
                mSQL = "SELECT formatea_hora(hora_inicio), formatea_hora(hora_inicio + total_minutos) FROM Intervalos"
            End If
            'mSQL = "SELECT hora_inicio, hora_inicio + total_minutos FROM Intervalos"
            mSQL = mSQL & " WHERE cod_horario = " & pcod_horario
            mSQL = mSQL & " ORDER BY hora_inicio"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            mLista = ""
            While mReader.Read
                If mLista <> "" Then
                    mLista = mLista & ";"
                End If
                If pFormato = "N" Then
                    mLista = mLista & Format(mReader(0), "0000") & "-" & Format(mReader(1), "0000")
                Else
                    mLista = mLista & mReader(0) & "-" & mReader(1)
                End If

            End While
            mReader.Close()

            Return mLista

        Catch ex As Exception
            Trata_Error("Error en Cadena_de_Intervalos", ex, mSQL)
        End Try

    End Function

    Public Function Cadena_de_Intervalos_Recuperacion(ByVal pcod_horario As Integer) As String Implements PresenciaDAO.Cadena_de_Intervalos_Recuperacion

        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try

            mCommand.Connection = mConexion
            mSQL = "SELECT inicio_intervalo, fin_intervalo FROM IntervaloRecuperacion "
            mSQL = mSQL & " WHERE cod_horario = " & pcod_horario
            mSQL = mSQL & " ORDER BY inicio_intervalo"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            mLista = ""
            While mReader.Read
                If mLista <> "" Then
                    mLista = mLista & ";"
                End If
                mLista = mLista & Format(mReader(0), "0000") & "-" & Format(mReader(1), "0000")
            End While
            mReader.Close()

            Return mLista

        Catch ex As Exception
            Trata_Error("Error en Cadena_de_Intervalos_Recuperacion", ex, mSQL)
        End Try

    End Function
    Public Function Cadena_de_Eventos(ByVal pDNI As String, ByVal pFecha As Date) As String Implements PresenciaDAO.Cadena_de_Eventos
        'da la lista de picadas
        'en formato E 0560;S 0590;

        Dim mLista_Eventos As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try

            mCommand.Connection = mConexion
            mSQL = "SELECT E_S, formatea_minuto(HORA)"
            mSQL = mSQL & " FROM EVENTOS "
            mSQL = mSQL & " WHERE DNI_EMPL = '" & pDNI & "'"
            mSQL = mSQL & " AND FECHA = '" & pFecha & "'"
            mSQL = mSQL & " ORDER BY HORA"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            While mReader.Read
                If mLista_Eventos <> "" Then
                    mLista_Eventos = mLista_Eventos & ";"
                End If
                mLista_Eventos = mLista_Eventos & mReader(0) & " " & Format(mReader(1), "0000")
            End While
            mReader.Close()


            Return mLista_Eventos

        Catch ex As Exception
            Trata_Error("Error en Cadena_de_Picadas", ex, mSQL)
        End Try

    End Function
    Public Function Inserta_Solicitud(ByVal pCODIGO As Long, ByVal pEstado As String, ByVal pDNI As String, ByVal pFecha As Date, ByVal pCod_Incidencia As Integer, ByVal pDesde As String, ByVal pHasta As String, ByVal pObservaciones As String, ByVal pSiguiente_Responsable As String, Optional ByVal pTipoEfecto As Integer = 1, Optional ByVal pCambioGrupo As Boolean = False, Optional ByVal pCod_solicitud_base As Long = 0) As Boolean Implements PresenciaDAO.Inserta_Solicitud
        'inserta una solicitud

        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            mSQL = "INSERT INTO Solicitud(CODIGO,estado, DNI, Fecha, Cod_Incidencia, Desde, Hasta, Observaciones"
            mSQL = mSQL & ",incidencia_original, Desde_original, Hasta_original, observaciones_original, id_siguiente_responsable, tipo, fecha_sol, CAMBIO_GRUPO, COD_SOLICITUD_BASE) "
            mSQL = mSQL & " VALUES(1,'" & pEstado & "','" & pDNI & "',to_date('" & pFecha.ToString("dd/MM/yyyy") & "','DD/MM/YYYY'),"
            If pCod_Incidencia >= 0 Then
                mSQL = mSQL & pCod_Incidencia
            Else
                mSQL = mSQL & "null"
            End If
            mSQL = mSQL & " ,'" & pDesde & "','" & pHasta & "','" & pObservaciones & "'"
            mSQL = mSQL & ", "
            If pCod_Incidencia >= 0 Then
                mSQL = mSQL & pCod_Incidencia
            Else
                mSQL = mSQL & "null"
            End If
            mSQL = mSQL & " ,'" & pDesde & "','" & pHasta & "','" & pObservaciones & "'"
            If pSiguiente_Responsable <> "" Then
                mSQL = mSQL & ",'" & pSiguiente_Responsable & "'"
            Else
                mSQL = mSQL & ",null"
            End If
            'tipo de solicitud
            mSQL = mSQL & "," & pTipoEfecto
            'fecha_sol
            mSQL = mSQL & ",sysdate"
            If pCambioGrupo Then
                mSQL = mSQL & ",'S'"
            Else
                mSQL = mSQL & ",'N'"
            End If
            If pCod_solicitud_base = 0 Then
                mSQL = mSQL & ", ''"
            Else
                mSQL = mSQL & "," & pCod_solicitud_base
            End If
            mSQL = mSQL & ")"

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Inserta_Solicitud", ex, mSQL)
        End Try

    End Function

    Public Function Actualiza_Campo_Empleado(ByVal pDNI As String, ByVal pCampo As String, ByVal pValor As String) As Boolean Implements PresenciaDAO.Actualiza_Campo_Empleado
        'actualiza un campo de un empleado

        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "UPDATE EMPLEADOS SET " & pCampo & " = " & pValor
            mSQL = mSQL & " WHERE DNI = '" & pDNI & "'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Actualiza_Campo_Empleado", ex, mSQL)
        End Try

    End Function

    Function Lista_Justificaciones_Dataset(ByRef pDatos As DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Justificaciones_Dataset

        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        mConsulta = "Select * "
        'mConsulta = "Select fecha "
        mConsulta &= " from justificaciones "
        mConsulta &= " where dni_empl = '" & pID_Usuario & "' "
        'If pEstado = "P" Then
        'mConsulta &= " and estado in  ('E','P') "
        'Else
        '    mConsulta &= " and estado in  ('" & pEstado & "') "
        'End If
        mConsulta &= " and  fecha_justificada >= to_date('" & pFecha_Desde & "','DD/MM/YYYY') "
        mConsulta &= " AND fecha_justificada <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY') "
        'mConsulta &= " group by fecha "
        'mConsulta &= " group by estado, fecha "
        If pOrden <> "" Then
            mConsulta &= " ORDER BY " & pOrden
        End If

        Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
        mDataAdapter.Fill(pDatos)
        Return True
    End Function
    Public Function Lista_Justificaciones(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pDNI As String = "", Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCod_Justificacion As Long = -1) As Object Implements PresenciaDAO.Lista_Justificaciones
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT fecha_justificada fecha, "
            mConsulta = mConsulta & " justificaciones.cod_incidencia, "
            mConsulta = mConsulta & " incidencias.desc_incidencia, "
            mConsulta = mConsulta & " desde_minutos, "
            mConsulta = mConsulta & " hasta_minutos, "
            mConsulta = mConsulta & " formatea_hora(desde_minutos) desde, "
            mConsulta = mConsulta & " formatea_hora(hasta_minutos) hasta, "
            mConsulta = mConsulta & " formatea_hora(duracion_minutos) duración, "
            mConsulta = mConsulta & " observaciones, cod_solic ,dni_empl, dni_empl_operador,maximo,fecha_base, incidencias.tipo,tiempo_maximo,fecha_termino,TipoFijo "
            mConsulta = mConsulta & " FROM justificaciones_t justificaciones, incidencias"
            mConsulta = mConsulta & " WHERE justificaciones.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND fecha_justificada >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " AND fecha_justificada <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pDNI <> "" Then
                mConsulta = mConsulta & " AND dni_empl = '" & pDNI & "'"
            End If
            If pCod_Incidencia >= 0 Then
                mConsulta = mConsulta & " AND justificaciones.cod_incidencia = " & pCod_Incidencia
            End If
            If pCod_Justificacion >= 0 Then
                mConsulta = mConsulta & " AND justificaciones.cod_justificacion = " & pCod_Justificacion
            End If
            mConsulta = mConsulta & " ORDER BY fecha_justificada"

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Justificaciones", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Solicitudes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False, Optional ByVal Cambio_Grupo As String = Nothing) As Object Implements PresenciaDAO.Lista_Solicitudes
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND ID_SIGUIENTE_RESPONSABLE IN ('" & pID_Lista_Responsables & "')"
            End If
            If pLista_Ultimo_Responsable <> "" Then
                mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            End If
            If pCodigoIncidencia >= 0 Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia in (" & pCodigoIncidencia & ")"
            End If
            If Not Cambio_Grupo Is Nothing Then
                mConsulta = mConsulta & " AND cambio_grupo =  '" & Cambio_Grupo & "'"
            End If
            If pOrdenUsuario Then
                mConsulta = mConsulta & " order by DNI, fecha, desde, solicitud.cod_incidencia"
            ElseIf pOrdenFechasol Then
                mConsulta = mConsulta & " order by fecha_sol, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            End If


            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_SolicitudesCuadrantes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As String = "", Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False, Optional ByVal Cambio_Grupo As String = Nothing) As Object Implements PresenciaDAO.Lista_SolicitudesCuadrantes
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND ID_SIGUIENTE_RESPONSABLE IN ('" & pID_Lista_Responsables & "')"
            End If
            If pLista_Ultimo_Responsable <> "" Then
                mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            End If
            If pCodigoIncidencia <> "" Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia in (" & pCodigoIncidencia & ")"
            End If
            If Not Cambio_Grupo Is Nothing Then
                mConsulta = mConsulta & " AND cambio_grupo =  '" & Cambio_Grupo & "'"
            End If
            If pOrdenUsuario Then
                mConsulta = mConsulta & " order by DNI, fecha, desde, solicitud.cod_incidencia"
            ElseIf pOrdenFechasol Then
                mConsulta = mConsulta & " order by fecha_sol, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            End If


            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Solicitudes_Movimiento(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean Implements PresenciaDAO.Lista_Solicitudes_Movimiento
        Dim mConsulta As String

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND ID_SIGUIENTE_RESPONSABLE IN ('" & pID_Lista_Responsables & "')"
            End If
            If pLista_Ultimo_Responsable <> "" Then
                mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            End If
            If pCodigoIncidencia >= 0 Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia =  " & pCodigoIncidencia
            End If
            If pOrdenUsuario Then
                mConsulta = mConsulta & " order by DNI, fecha, desde, solicitud.cod_incidencia"
            ElseIf pOrdenFechasol Then
                mConsulta = mConsulta & " order by fecha_sol, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Solicitudes_Movimiento_Grupos(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean Implements PresenciaDAO.Lista_Solicitudes_Movimiento_Grupos
        Dim mConsulta As String

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND exists (select * from siguientes_solicitud where cod_solicitud = solicitud.codigo and dni in ('" & pID_Lista_Responsables & "'))"
            End If
            If pLista_Ultimo_Responsable <> "" Then
                mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            End If
            If pCodigoIncidencia >= 0 Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia =  " & pCodigoIncidencia
            End If
            If pOrdenUsuario Then
                mConsulta = mConsulta & " order by DNI, fecha, desde, solicitud.cod_incidencia"
            ElseIf pOrdenFechasol Then
                mConsulta = mConsulta & " order by fecha_sol, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Asignacion_Responsable(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Lista_Responsables As String = "") As Object Implements PresenciaDAO.Lista_Asignacion_Responsable
        'da lista de asignaciones de responsables
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT ID_USUARIO, ID_RESPONSABLE "
            mConsulta = mConsulta & " FROM asignacion_responsable"
            mConsulta = mConsulta & " , empleados "
            mWhere = " WHERE asignacion_responsable.ID_USUARIO = empleados.dni "
            If pID_Usuario <> "" Then
                mWhere = mWhere & " AND ID_USUARIO = '" & pID_Usuario & "'"
            End If
            If pID_Lista_Responsables <> "" Then
                mWhere = mWhere & " AND ID_RESPONSABLE IN ( '" & pID_Lista_Responsables & "')"
            End If
            'mConsulta = mConsulta & mWhere & " order by ID_USUARIO, ID_RESPONSABLE "
            mConsulta = mConsulta & mWhere & " order by empleados.ape1, empleados.ape2, empleados.nombre"

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Asignacion_Responsable", ex, mConsulta)
        End Try

    End Function



    Public Function Lista_Aprobaciones(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "") As Object Implements PresenciaDAO.Lista_Aprobaciones
        'da lista de aprobaciones
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT ID_SOLICITUD, ID_RESPONSABLE, FECHA , OPERACION, ID_DELEGADO, CAUSA_DENEGACION "
            mConsulta = mConsulta & " FROM aprobacion"
            If pCod_Solicitud >= 0 Then
                mWhere = "WHERE ID_SOLICITUD = " & pCod_Solicitud
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & " " & mWhere & " order by ID_SOLICITUD, FECHA "

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Aprobaciones", ex, mConsulta)
        End Try
    End Function

    Public Function Lista_Operaciones_Pendientes(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "") As Object Implements PresenciaDAO.Lista_OperacionesPendientes
        'da lista de aprobaciones
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT siguientes_solicitud.COD_SOLICITUD, empleados.nombre || ' ' || empleados.ape1 || ' ' || empleados.ape2 RESPONSABLE,  '' FECHA , decode(solicitud.ESTADO,'E','En Curso', 'P', 'Pendiente', 'A', 'Autorizada') ESTADO "
            mConsulta = mConsulta & " FROM solicitud,siguientes_solicitud,empleados"
            If pCod_Solicitud >= 0 Then
                mWhere = "WHERE CODIGO = " & pCod_Solicitud
            End If
            mWhere = mWhere & " and solicitud.codigo = siguientes_solicitud.cod_solicitud and solicitud.estado <> 'D' and empleados.dni = siguientes_solicitud.dni"
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & " " & mWhere & " order by CODIGO, FECHA "

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Operaciones_Pendientes", ex, mConsulta)
        End Try
    End Function

    Public Function Actualiza_Solicitud(ByVal pCODIGO As Long, Optional ByVal pEstado As String = Nothing, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDesde As String = Nothing, Optional ByVal pHasta As String = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pID_Sig_Responsable As String = Nothing, Optional ByVal pCod_Justificacion As Long = -1, Optional ByVal pUltimo_Responsable As String = Nothing, Optional ByVal pTipo As String = Nothing, Optional ByVal pCambioGrupo As String = Nothing, Optional ByVal pCod_solicitud_base As Long = 0) As Boolean Implements PresenciaDAO.Actualiza_Solicitud
        'actualiza la solicitud

        Dim mConsulta As String

        Try
            mConsulta = " UPDATE SOLICITUD SET "
            If Not IsNothing(pEstado) Then
                mConsulta = mConsulta & " ESTADO = '" & pEstado & "' "
            Else
                mConsulta = mConsulta & " ESTADO = ESTADO "
            End If
            'mConsulta = mConsulta & " ,DNI = '" & pDNI & "'"
            'mConsulta = mConsulta & " ,FECHA = to_date('" & pFecha.ToString("DD/MM/YYYY") & "','DD/MM/YYYY')"
            If pCod_Incidencia >= 0 Then
                mConsulta = mConsulta & " ,COD_INCIDENCIA = " & pCod_Incidencia
            End If
            If Not IsNothing(pDesde) Then
                mConsulta = mConsulta & " ,DESDE = '" & pDesde & "'"
            End If
            If Not IsNothing(pHasta) Then
                mConsulta = mConsulta & " ,HASTA = '" & pHasta & "'"
            End If
            If Not IsNothing(pObservaciones) Then
                If pObservaciones <> "" Then
                    mConsulta = mConsulta & " ,Observaciones = '" & QuitaComilla(pObservaciones) & "'"
                Else
                    mConsulta = mConsulta & " ,Observaciones = null"
                End If
            End If
            If Not IsNothing(pID_Sig_Responsable) Then
                If pID_Sig_Responsable = "" Then
                    mConsulta = mConsulta & " ,ID_SIGUIENTE_RESPONSABLE = null"
                Else
                    mConsulta = mConsulta & " ,ID_SIGUIENTE_RESPONSABLE = '" & pID_Sig_Responsable & "'"
                End If
            End If
            If pCod_Justificacion > 0 Then
                mConsulta = mConsulta & " ,COD_JUSTIFICACION = " & pCod_Justificacion
            ElseIf pCod_Justificacion = 0 Then
                mConsulta = mConsulta & " ,COD_JUSTIFICACION = NULL "
            End If
            If Not IsNothing(pUltimo_Responsable) Then
                If pUltimo_Responsable <> "" Then
                    mConsulta = mConsulta & ",ultimo_responsable = '" & pUltimo_Responsable & "'"
                Else
                    mConsulta = mConsulta & ",ultimo_responsable = null"
                End If
            End If
            If Not IsNothing(pTipo) Then
                If IsNumeric(pTipo) Then
                    mConsulta = mConsulta & " ,tipo = " & pTipo
                End If
            End If
            If Not IsNothing(pCambioGrupo) Then
                If CBool(pCambioGrupo) Then
                    mConsulta = mConsulta & " ,CAMBIO_GRUPO = 'S'"
                Else
                    mConsulta = mConsulta & " ,CAMBIO_GRUPO = 'N'"
                End If
            End If
            If pCod_solicitud_base > 0 Then
                mConsulta = mConsulta & " ,COD_SOLICITUD_BASE = " & pCod_solicitud_base
            ElseIf pCod_solicitud_base = 0 Then
                'mConsulta = mConsulta & " ,COD_SOLICITUD_BASE = NULL "
            End If
            mConsulta = mConsulta & " WHERE CODIGO = " & pCODIGO
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Solicitud", ex, mConsulta)
        End Try
    End Function

    Public Function Inserta_Aprobacion(ByVal pCod_Solicitud As Long, ByVal pID_Responsable As String, Optional ByVal pOperacion As String = "A", Optional ByVal pDelegado As String = "") As Boolean Implements PresenciaDAO.Inserta_Aprobacion
        'inserta una aprobacion
        Dim mConsulta As String
        Dim mIDResponsable As String
        Dim mIDDelegado As String
        Try
            mIDResponsable = pID_Responsable
            mIDDelegado = pDelegado
            If mIDResponsable = "" And mIDDelegado <> "" Then
                mIDResponsable = mIDDelegado
                mIDDelegado = ""
            End If

            mConsulta = "INSERT INTO APROBACION(ID_SOLICITUD, ID_RESPONSABLE, FECHA, OPERACION,ID_DELEGADO)"
            mConsulta = mConsulta & " VALUES(" & pCod_Solicitud & ",'" & mIDResponsable & "',SYSDATE,'" & pOperacion & "'"
            If mIDDelegado <> "" Then
                mConsulta = mConsulta & ",'" & mIDDelegado & "')"
            Else
                mConsulta = mConsulta & ",NULL)"
            End If
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Aprobacion", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Inserta_Justificacion(ByVal pFecha_Justificada As Date, ByVal pDesde_minutos As Integer, ByVal pHasta_minutos As Integer, ByVal pDni As String, ByVal pOperador As String, ByVal pCod_Incidencia As Integer, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing, Optional ByVal pEfecto As Integer = 1) As Long Implements PresenciaDAO.Inserta_Justificacion

        Dim mConsulta As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim auxcodigo As Long = 0
            Dim mDesde_minutos As Integer
            Dim mHasta_minutos As Integer
            mDesde_minutos = pDesde_minutos
            mHasta_minutos = pHasta_minutos
            If mHasta_minutos < mDesde_minutos Then
                mHasta_minutos = mHasta_minutos + 1440
            End If

            'esto ya lo hace un triger
            'mConsulta = "SELECT MAX(Cod_Justificacion) FROM JUSTIFICACIONES"
            'mCommand.Connection = mConexion
            'mCommand.CommandText = mConsulta
            'mReader = mCommand.ExecuteReader
            'If mReader.Read Then
            '    auxcodigo = NVL(mReader(0), 0)
            'End If
            'mReader.Close()
            'auxcodigo = auxcodigo + 1

            auxcodigo = 1

            mConsulta = "INSERT INTO JUSTIFICACIONES(COD_JUSTIFICACION,FECHA_HORA,FECHA_JUSTIFICADA,DNI_EMPL,DNI_EMPL_OPERADOR,"
            mConsulta = mConsulta & " DESDE_MINUTOS, HASTA_MINUTOS,DURACION_MINUTOS,COD_INCIDENCIA,OBSERVACIONES, "
            mConsulta = mConsulta & " CONTABILIZADA,COD_SOLIC) "
            mConsulta = mConsulta & " VALUES(" & auxcodigo & ",SYSDATE,TO_DATE('" & pFecha_Justificada.ToString("dd/MM/yyyy") & "','DD/MM/YYYY'),'" & pDni & "','" & pOperador & "',"
            mConsulta = mConsulta & mDesde_minutos & "," & mHasta_minutos & "," & (mHasta_minutos - mDesde_minutos) * pEfecto & "," & pCod_Incidencia & ","
            If IsNothing(pObservaciones) Then
                mConsulta = mConsulta & "null,"
            Else
                If pObservaciones <> "" Then
                    mConsulta = mConsulta & "'" & QuitaComilla(pObservaciones) & "',"
                Else
                    mConsulta = mConsulta & "null,"
                End If
            End If
            mConsulta = mConsulta & "'N',"
            If IsNothing(pCod_Solicitud) Then
                mConsulta = mConsulta & "null"
            Else
                If pCod_Solicitud > 0 Then
                    mConsulta = mConsulta & pCod_Solicitud
                Else
                    mConsulta = mConsulta & "null"
                End If
            End If
            mConsulta = mConsulta & ")"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            mConsulta = "SELECT MAX(Cod_Justificacion) FROM JUSTIFICACIONES WHERE DNI_EMPL = '" & pDni & "' AND FECHA_JUSTIFICADA = TO_DATE('" & pFecha_Justificada.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader
            If mReader.Read Then
                auxcodigo = NVL(mReader(0), 0)
            End If
            mReader.Close()


            Return auxcodigo
        Catch ex As Exception
            Trata_Error("Error en Inserta_Justificaciones", ex, mConsulta)
            Return -1
        End Try

    End Function

    Public Function Elimina_Aprobacion(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pOperacion As String = "") As Boolean Implements PresenciaDAO.Elimina_Aprobacion
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE Aprobacion "
            If pCod_Solicitud >= 0 Then
                mWhere = " WHERE ID_SOLICITUD = " & pCod_Solicitud
            End If
            If pID_Responsable <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE"
                Else
                    mWhere &= " AND"
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            If pOperacion <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE"
                Else
                    mWhere &= " AND"
                End If
                mWhere = mWhere & " OPERACION = '" & pOperacion & "'"
            End If
            mConsulta = mConsulta & mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Aprobacion", ex, mConsulta)
        End Try
    End Function

    'Public Function Lista_Asignacion_Responsable_Usuario(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Responsable As String = "") As Object Implements PresenciaDAO.Lista_Asignacion_Responsable_Usuario
    '    'da lista de asignaciones de responsables
    '    Dim mConsulta As String
    '    Dim mWhere As String
    '    Dim mCommand As New OleDb.OleDbCommand()
    '    Dim mReader As Object

    '    Try
    '        mConsulta = "SELECT ID_USUARIO, ID_RESPONSABLE "
    '        mConsulta = mConsulta & " FROM asignacion_responsable_usuario"
    '        If pID_Usuario <> "" Then
    '            mWhere = " WHERE ID_USUARIO = '" & pID_Usuario & "'"
    '        End If
    '        If pID_Responsable <> "" Then
    '            If mWhere <> "" Then
    '                mWhere = mWhere & " AND "
    '            Else
    '                mWhere = mWhere & " WHERE "
    '            End If
    '            mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
    '        End If
    '        mConsulta = mConsulta & mWhere & " order by ID_USUARIO, ID_RESPONSABLE "

    '        mCommand.CommandText = mConsulta
    '        mCommand.Connection = mConexion
    '        mReader = mCommand.ExecuteReader()
    '        Return mReader

    '    Catch ex As Exception
    '        Trata_Error("Error en Lista_Asignacion_Responsable_usuario", ex)
    '    End Try

    'End Function

    Public Function Inserta_Asignacion_Responsable(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Inserta_Asignacion_Responsable
        Dim mConsulta As String
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "INSERT INTO asignacion_responsable(ID_USUARIO,ID_RESPONSABLE)"
            mConsulta = mConsulta & " VALUES('" & pID_Usuario & "','" & pID_Responsable & "')"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asignacion_Responsable", ex, mConsulta)
        End Try
    End Function

    'Public Function Inserta_Asignacion_Responsable_Usuario(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Inserta_Asignacion_Responsable_Usuario
    '    Dim mConsulta As String
    '    Dim mCommand As New OleDb.OleDbCommand()
    '    mConsulta = "INSERT INTO asignacion_responsable_usuario(ID_USUARIO,ID_RESPONSABLE)"
    '    mConsulta = mConsulta & " VALUES('" & pID_Usuario & "','" & pID_Responsable & "')"
    '    mCommand.Connection = mConexion
    '    mCommand.CommandText = mConsulta
    '    mCommand.ExecuteNonQuery()
    'End Function

    Public Function Elimina_Asignacion_Responsable(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Responsable As String = "") As Boolean Implements PresenciaDAO.Elimina_Asignacion_Responsable
        Dim mConsulta As String
        Dim mWhere As String
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "DELETE asignacion_responsable "
            If pID_Usuario <> "" Then
                mWhere = " WHERE ID_USUARIO = '" & pID_Usuario & "'"
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & mWhere
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asignacion_Responsable", ex, mConsulta)
        End Try
    End Function

    'Public Function Elimina_Asignacion_Responsable_Usuario(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Elimina_Asignacion_Responsable_Usuario
    '    Dim mConsulta As String
    '    Dim mCommand As New OleDb.OleDbCommand()
    '    mConsulta = "DELETE asignacion_responsable_usuario "
    '    mConsulta = mConsulta & " WHERE ID_USUARIO = '" & pID_Usuario & "'"
    '    mCommand.Connection = mConexion
    '    mCommand.CommandText = mConsulta
    '    mCommand.ExecuteNonQuery()
    'End Function



    Public Function Lista_Intervalos_Presencia(Optional ByVal pDNI As String = "", Optional ByVal pFecha As String = "") As Object Implements PresenciaDAO.Lista_Intervalos_Presencia
        Dim mConsulta As String
        Dim mWhere As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            'obtengo los intervalos de presencia
            mConsulta = "select formatea_hora(entrada),formatea_hora(salida) from aux_emparejamiento "
            'mConsulta = mConsulta & "where fecha = to_date('" & Format(pFecha, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFecha <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " fecha = to_date('" & pFecha & "','DD/MM/YYYY')"
            End If

            If pDNI <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " dni = '" & pDNI & "'"
            End If
            mConsulta = mConsulta & mWhere
            mConsulta = mConsulta & " order by to_number(entrada) "
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            Return mCommand.ExecuteReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Intervalos_Presencia", ex, mConsulta)
        End Try
    End Function

    Public Function Busqueda_Empleados(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pOrden As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pCalcula_Saldo As String = "") As Object Implements PresenciaDAO.Busqueda_Empleados
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE SUPR_ACCENT(UPPER(DNI)) LIKE SUPR_ACCENT(UPPER('" & pDNI & "%'))"
            End If
            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If
            If pCalcula_Saldo <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " calcula_saldo = '" & pCalcula_Saldo & "'"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            If pOrden <> "" Then
                mConsulta = mConsulta & " ORDER BY " & pOrden
            Else
                'mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"
                mConsulta = mConsulta & " ORDER BY SUPR_ACCENT(UPPER(APE1)),SUPR_ACCENT(UPPER(APE2)),SUPR_ACCENT(UPPER(NOMBRE))"
            End If


            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Personal(Optional ByVal pDNI As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pTarjeta As String = "") As Data.DataSet Implements DAO.PresenciaDAO.Busqueda_Personal
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try


            mConsulta = "SELECT DNI,APE1 as Apellido1,APE2 as Apellido2,NOMBRE,TIPO FROM EMPLEADOS_PROVEEDORES"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE UPPER('" & pDNI & "%')"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " UPPER(APE1) LIKE UPPER('" & pApe1 & "%')"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " UPPER(APE2)) LIKE UPPER('" & pApe2 & "%')"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE UPPER('" & pClave_Empleado & "%')"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " UPPER(ape1 || ' ' || ape2) LIKE UPPER('" & pApellidos & "%')"
            End If
            'If pTarjeta <> "" Then
            'If mWhere <> "" Then
            'mWhere = mWhere & " AND "
            'Else
            '    mWhere = " WHERE "
            'End If
            'mWhere = mWhere & "PAN_TARJETA LIKE '" & pTarjeta & "' ORDER BY DNI"
            'End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function Ultima_Tarjeta(Optional ByVal pDNI As String = "", Optional ByVal Tipo As String = "") As Data.DataSet Implements DAO.PresenciaDAO.Ultima_Tarjeta
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try



            mConsulta = "SELECT pan_tarjeta from tarjetasasociadas where dni"
            If Tipo = "V" Then
                mConsulta = mConsulta & "_vis"
            ElseIf Tipo = "P" Then
                mConsulta = mConsulta & "_prov"
            Else
                mConsulta = mConsulta & "_empl"
            End If
            'strsql = strsql & " = '" & dni & "' and fecha_hora_baja is null order by fecha_hora_alta desc"
            mConsulta = mConsulta & " = '" & pDNI & "' order by fecha_hora_alta desc"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Ultima_Tarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Personal_Tarjeta(Optional ByVal pDNI As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pTarjeta As String = "", Optional ByVal pTipo As Integer = 0) As Data.DataSet Implements DAO.PresenciaDAO.Busqueda_Personal_Tarjeta
        'devuelve la tabla de empleados
        Dim mConsulta As String

        Try


            If pTipo = 1 Then
                mConsulta = "SELECT DNI,APE1,APE2,NOMBRE,PAN_TARJETA FROM EMPLEADOS_PROVEEDORES WHERE "
                mConsulta = mConsulta & "PAN_TARJETA LIKE '" & pTarjeta & "' ORDER BY DNI"
            ElseIf pTipo = 2 Then
                mConsulta = "SELECT DNI,APE1,APE2,NOMBRE,CLAVE_EMP FROM EMPLEADOS WHERE "
                mConsulta = mConsulta & "CLAVE_EMP ='" & pClave_Empleado & "' ORDER BY CLAVE_EMP"
            ElseIf pTipo = 3 Then
                mConsulta = "SELECT DNI,APE1,APE2,NOMBRE,CLAVE_EMP FROM EMPLEADOS WHERE "
                mConsulta = mConsulta & "DNI LIKE('" & pDNI & "%') ORDER BY DNI"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Busca_Personal_tarjeta", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Visita_Tarjeta_Asignada(Optional ByVal pTarjeta As String = "") As Data.DataSet Implements DAO.PresenciaDAO.Busqueda_Visita_Tarjeta_Asignada
        'devuelve la tabla de empleados
        Dim mConsulta As String

        Try

            mConsulta = "SELECT DNI,APE1,APE2,NOMBRE FROM VISITANTES,TARJETASASOCIADAS WHERE TARJETASASOCIADAS.PAN_TARJETA='" & pTarjeta & "' and visitantes.dni = tarjetasasociadas.dni_vis AND (fecha_hora_baja is NULL or fecha_hora_baja >= sysdate)"



            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet
        Catch ex As Exception
            Trata_Error("Error en Busqueda_Visita_Tarjeta_Asignada", ex, mConsulta)
        End Try

    End Function

    Public Function Actualiza_Empleado(ByVal pDNI As String, Optional ByVal pNombre As String = Nothing, Optional ByVal pApe1 As String = Nothing, Optional ByVal pApe2 As String = Nothing, Optional ByVal pClave_Empleado As String = Nothing, Optional ByVal pCentro As String = Nothing, Optional ByVal pCargo As String = Nothing, Optional ByVal pEmail As String = Nothing, Optional ByVal pTelefono As String = Nothing, Optional ByVal pCalcula_Saldo As Integer = -10, Optional ByVal pAdmin As Integer = -10, Optional ByVal pEmpresa As String = "", Optional ByVal pFecha_Antiguedad As String = "", Optional ByVal pUsuarioLdap As String = "") As Boolean Implements PresenciaDAO.Actualiza_Empleado
        Dim mConsulta As String
        Dim mSet As String
        Try
            'Consulta = "UPDATE EMPLEADOS set NOMBRE='" & pNombre & "'" _
            '            & ", APE1=" & IIf(pApe1 = "", "null", "'" & pApe1 & "'") _
            '            & ", APE2=" & IIf(pApe2 = "", "null", "'" & pApe2 & "'") _
            '            & ", CLAVE_EMP=" & pClave_Empleado _
            '            & ", CENTRO=" & IIf(pCentro = "", "null", "'" & pCentro & "'") _
            '            & ", CARGO=" & IIf(pCargo = "", "null", "'" & pCargo & "'") _
            '            & ", EMAIL=" & IIf(pEmail = "", "null", "'" & pEmail & "'") _
            '            & ", TELEFONO=" & IIf(pTelefono = "", "null", "'" & pTelefono & "'") _
            '            & ", CALCULA_SALDO='" & IIf(pCalcula_Saldo, "S", "N") & "'" _
            '            & ", ADMIN=" & IIf(pAdmin, "'1'", "Null") _
            '            & " WHERE DNI='" & pDNI & "'"

            mConsulta = "UPDATE EMPLEADOS set "
            If Not IsNothing(pNombre) Then
                If pNombre = "" Then
                    pNombre = "null"
                Else
                    pNombre = "'" & pNombre & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " NOMBRE = " & pNombre
            End If
            If Not IsNothing(pApe1) Then
                If pApe1 = "" Then
                    pApe1 = "null"
                Else
                    pApe1 = "'" & pApe1 & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " Ape1 = " & pApe1
            End If
            If Not IsNothing(pApe2) Then
                If pApe2 = "" Then
                    pApe2 = "null"
                Else
                    pApe2 = "'" & pApe2 & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " Ape2 = " & pApe2
            End If
            If Not IsNothing(pClave_Empleado) Then
                If pApe2 = "" Then
                    pClave_Empleado = "null"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " CLAVE_EMP = " & pClave_Empleado
            End If
            If Not IsNothing(pCentro) Then
                If pCentro = "" Then
                    pCentro = "null"
                Else
                    pCentro = "'" & pCentro & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " CENTRO = " & pCentro
            End If
            If Not IsNothing(pCargo) Then
                If pCargo = "" Then
                    pCargo = "null"
                Else
                    pCargo = "'" & pCargo & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " CARGO = " & pCargo
            End If
            If Not IsNothing(pEmail) Then
                If pEmail = "" Then
                    pEmail = "null"
                Else
                    pEmail = "'" & pEmail & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " EMAIL = " & pEmail
            End If
            If Not IsNothing(pTelefono) Then
                If pTelefono = "" Then
                    pTelefono = "null"
                Else
                    pTelefono = "'" & pTelefono & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " TELEFONO = " & pTelefono
            End If
            If Not pCalcula_Saldo = -10 Then
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                If CBool(pCalcula_Saldo) Then
                    mSet &= " CALCULA_SALDO = 'S'"
                Else
                    mSet &= " CALCULA_SALDO = 'N'"
                End If
            End If
            If Not pAdmin = -10 Then
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                If CBool(pAdmin) Then
                    mSet &= " ADMIN = 1"
                Else
                    mSet &= " ADMIN = 0"
                End If
            End If
            If Not IsNothing(pEmpresa) Then
                If pEmpresa = "" Then
                    pEmpresa = "null"
                Else
                    pEmpresa = "'" & pEmpresa & "'"
                End If
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " EMPRESA = " & pEmpresa
            End If
            If pFecha_Antiguedad = "" Then
            Else
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " FECHA_ANTIGUEDAD = '" & pFecha_Antiguedad & "'"
            End If
            If pUsuarioLdap = "" Then
            Else
                If mSet <> "" Then
                    mSet = mSet & ", "
                End If
                mSet &= " USUARIO = '" & pUsuarioLdap & "'"
            End If
            mConsulta &= " " & mSet & " WHERE DNI='" & pDNI & "'"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            Trata_Error("Error en Actualiza_Empleado", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Empleado(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_Empleado
        Dim mConsulta As String
        Try
            mConsulta = "DELETE FROM EMPLEADOS WHERE DNI='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Empleado", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Elimina_TarjetasAsociadas(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_TarjetasAsociadas
        Dim mConsulta As String
        Try
            mConsulta = "DELETE tarjetasasociadas where dni_empl='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_TarjetasAsociadas", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Elimina_PerteneceA(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_PerteneceA
        Dim mConsulta As String
        Try
            mConsulta = "DELETE pertenecena where dni_empl='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_PerteneceA", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Elimina_AsociaUsuarioGrupoTrabajo(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_AsociaUsuarioGrupoTrabajo
        Dim mConsulta As String
        Try
            mConsulta = "DELETE asociausuariogrupotrabajo where dni='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_AsociaUsuarioGrupoTrabajo", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Elimina_VisitasRealizadas(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_VisitasRealizadas
        Dim mConsulta As String
        Try
            mConsulta = "DELETE visitas where dni_empl_visitado='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_VisitasRealizadas", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Elimina_VisitasRegistradas(ByVal pDNI As String) As Boolean Implements PresenciaDAO.Elimina_VisitasRegistradas
        Dim mConsulta As String
        Try
            mConsulta = "DELETE visitas where dni_empl_operario='" & pDNI & "'"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_VisitasRealizadas", ex, mConsulta)
            Return -1
        End Try
    End Function

    Public Function Inserta_Empleado(ByVal pDNI As String, ByVal pNombre As String, ByVal pApe1 As String, ByVal pApe2 As String, ByVal pClave_Empleado As String, ByVal pCentro As String, ByVal pCargo As String, ByVal pEmail As String, ByVal pTelefono As String, ByVal pCalcula_Saldo As Boolean, Optional ByVal pPuedeRecib As Boolean = False, Optional ByVal pEmpresa As String = "", Optional ByVal pFecha_Antiguedad As String = "", Optional ByVal pUsuarioLdap As String = "") As Boolean Implements PresenciaDAO.Inserta_Empleado
        Dim mConsulta As String
        Dim PermiteVisita As String
        Dim calcula_Sal As String
        Try
            'mConsulta = "INSERT INTO EMPLEADOS (DNI, NOMBRE, APE1, APE2, CLAVE_EMP, CENTRO, CARGO, EMAIL, TELEFONO, CALCULA_SALDO, ADMIN, PERMITE_VISITA) " _
            '           & " VALUES ('" & UCase(pDNI) & "'" _
            '          & ",'" & pNombre & "'" _
            '         & ",'" & pApe1 & "'" _
            '        & ",'" & pApe2 & "'" _
            '       & "," & pClave_Empleado _
            '      & "," & IIf(pCentro = "", "null", "'" & pCentro & "'") _
            '     & "," & IIf(pCargo = "", "null", "'" & pCargo & "'") _
            '    & "," & IIf(pEmail = "", "null", "'" & pEmail & "'") _
            '   & "," & IIf(pTelefono = "", "null", "'" & pTelefono & "'") _
            '  & "," & IIf(pCalcula_Saldo, "'S'", "'N'") _
            ' & "," & IIf(pAdmin, "'1'", "null") & ", 'N') _"
            '& ", 'N'"
            If pPuedeRecib = True Then
                PermiteVisita = "S"
            Else
                PermiteVisita = "N"
            End If

            If pCalcula_Saldo = True Then
                calcula_Sal = "S"
            Else
                calcula_Sal = "N"
            End If

            If pFecha_Antiguedad = "" Then
                mConsulta = "INSERT INTO EMPLEADOS (DNI,CLAVE_EMP,NOMBRE,APE1,APE2,PERMITE_VISITA,CALCULA_SALDO,TELEFONO,EMAIL,CENTRO,CARGO,EMPRESA,USUARIO) " & "VALUES ('" & pDNI & "','" & pClave_Empleado & "','" & pNombre & "','" & pApe1 & "','" & pApe2 & "','" & PermiteVisita & "','" & calcula_Sal & "','" & pTelefono & "','" & pEmail & "','" & pCentro & "','" & pCargo & "','" & pEmpresa & "', '" & pUsuarioLdap & "')"
            Else
                mConsulta = "INSERT INTO EMPLEADOS (DNI,CLAVE_EMP,NOMBRE,APE1,APE2,PERMITE_VISITA,CALCULA_SALDO,TELEFONO,EMAIL,CENTRO,CARGO,EMPRESA, FECHA_ANTIGUEDAD,USUARIO) " & "VALUES ('" & pDNI & "','" & pClave_Empleado & "','" & pNombre & "','" & pApe1 & "','" & pApe2 & "','" & PermiteVisita & "','" & calcula_Sal & "','" & pTelefono & "','" & pEmail & "','" & pCentro & "','" & pCargo & "','" & pEmpresa & "', '" & pFecha_Antiguedad & "', '" & pUsuarioLdap & "')"
            End If


            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Empleado", ex, mConsulta)
            Return -1
        End Try
    End Function



    Public Function Actualiza_Justificacion(ByVal pCodigo As Long, ByVal pFecha_Justificada As Date, ByVal pDesde_minutos As Integer, ByVal pHasta_minutos As Integer, Optional ByVal pDni As String = Nothing, Optional ByVal pOperador As String = Nothing, Optional ByVal pCod_Incidencia As Integer = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing, Optional ByVal pEfecto As Integer = 1) As Boolean Implements PresenciaDAO.Actualiza_Justificacion
        Dim mConsulta As String
        Try
            mConsulta = "UPDATE JUSTIFICACIONES SET FECHA_HORA = SYSDATE "
            mConsulta = mConsulta & " ,FECHA_JUSTIFICADA = TO_DATE('" & pFecha_Justificada.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            If Not IsNothing(pDni) Then
                mConsulta = mConsulta & " ,DNI_EMPL = '" & pDni & "'"
            End If
            mConsulta = mConsulta & " ,DNI_EMPL_OPERADOR = '" & pOperador & "'"
            mConsulta = mConsulta & " ,DESDE_MINUTOS = " & pDesde_minutos
            mConsulta = mConsulta & " ,HASTA_MINUTOS = " & pHasta_minutos
            mConsulta = mConsulta & " ,DURACION_MINUTOS = " & (pHasta_minutos - pDesde_minutos) * pEfecto
            mConsulta = mConsulta & " ,COD_INCIDENCIA = " & pCod_Incidencia
            If Not IsNothing(pObservaciones) Then
                If pObservaciones <> "" Then
                    mConsulta = mConsulta & " ,OBSERVACIONES = '" & QuitaComilla(pObservaciones) & "'"
                End If
            End If
            mConsulta = mConsulta & " ,CONTABILIZADA = 'N'"
            If IsNothing(pCod_Solicitud) Then
                If pCod_Solicitud > 0 Then
                    mConsulta = mConsulta & " ,COD_SOLIC = " & pCod_Solicitud
                End If
            End If
            mConsulta = mConsulta & " WHERE COD_JUSTIFICACION = " & pCodigo

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Justificaciones", ex, mConsulta)
            Return -1
        End Try

    End Function

    Public Function Actualiza_Justificacion_Cod_Base(ByVal pCodigoJustf As Long, ByVal pCodigoBase As Long) As Boolean Implements PresenciaDAO.Actualiza_Justificacion_Cod_Base
        Dim mConsulta As String
        Try
            If pCodigoBase > 0 Then
                mConsulta = "UPDATE JUSTIFICACIONES SET COD_JUSTIFICACION_BASE=" & pCodigoBase
                mConsulta = mConsulta & " WHERE COD_JUSTIFICACION = " & pCodigoJustf

                Dim mCommand As New OleDb.OleDbCommand
                mCommand.Connection = mConexion
                mCommand.CommandText = mConsulta
                mCommand.ExecuteNonQuery()
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("Error en Actualiza_Justificacion_Cod_Base", ex, mConsulta)
            Return -1
        End Try

    End Function

    Public Function Elimina_Justificacion(Optional ByVal pCodigo As Long = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing) As Boolean Implements PresenciaDAO.Elimina_Justificacion
        Dim mConsulta As String
        Try
            Dim mWhere As String
            mConsulta = "DELETE JUSTIFICACIONES WHERE "
            If pCodigo > 0 Then
                mWhere = " COD_JUSTIFICACION = " & pCodigo
            End If
            If pCod_Solicitud > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = " COD_SOLIC = " & pCod_Solicitud
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & mWhere
                Dim mCommand As New OleDb.OleDbCommand
                mCommand.Connection = mConexion
                mCommand.CommandText = mConsulta
                mCommand.ExecuteNonQuery()
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Trata_Error("Error en Elimina_Justificacion", ex, mConsulta)
            Return -1
        End Try

    End Function
    Public Function Elimina_Solicitud(ByVal pCodigo As Long) As Boolean Implements PresenciaDAO.Elimina_Solicitud
        Dim mConsulta As String
        Try
            Dim mWhere As String
            mConsulta = "DELETE SOLICITUD WHERE "
            If pCodigo > 0 Then
                mWhere = " CODIGO = " & pCodigo
            End If

            mConsulta = mConsulta & mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True


        Catch ex As Exception
            Trata_Error("Error en Elimina_Solicitud", ex, mConsulta)
            Return False
        End Try

    End Function
    Public Function Elimina_Solicitud(ByVal pListaCodigos As String) As Boolean Implements PresenciaDAO.Elimina_Solicitud
        Dim mConsulta As String
        Try
            Dim mWhere As String
            mConsulta = "DELETE SOLICITUD WHERE "
            mWhere = " CODIGO in (" & pListaCodigos & ")"
            mConsulta = mConsulta & mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True


        Catch ex As Exception
            Trata_Error("Error en Elimina_Solicitud", ex, mConsulta)
            Return False
        End Try

    End Function
    Public Function Resumen_Justificaciones(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pListaDNI As String = "", Optional ByVal pListaIncidencia As String = "") As Object Implements PresenciaDAO.Resumen_Justificaciones
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT "
            mConsulta = mConsulta & " justificaciones.cod_incidencia, "
            mConsulta = mConsulta & " count(1) Cuenta, "
            mConsulta = mConsulta & " count(distinct fecha_justificada) Dias, "
            mConsulta = mConsulta & " formatea_hora(sum(duracion_minutos)) duración "
            mConsulta = mConsulta & " FROM justificaciones_t justificaciones"
            mConsulta = mConsulta & " WHERE fecha_justificada >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " AND fecha_justificada <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pListaDNI <> "" Then
                mConsulta = mConsulta & " AND dni_empl in  (" & pListaDNI & ")"
            End If
            If pListaIncidencia <> "" Then
                mConsulta = mConsulta & " AND justificaciones.cod_incidencia in (" & pListaIncidencia & ")"
            End If
            mConsulta = mConsulta & " GROUP BY justificaciones.cod_incidencia"
            mConsulta = mConsulta & " ORDER BY sum(duracion_minutos) desc "


            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Resumen_Justificaciones", ex, mConsulta)
        End Try

    End Function

    Public Function Acumuladores_Usuario(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pSoloFavoritos As Boolean = False) As Object Implements PresenciaDAO.Acumuladores_Usuario
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, DESCRIPCION, "
            mConsulta = mConsulta & " valor_acumulador_todosf('" & pDNI & "',codigo,TO_DATE('" & pFecha.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')) valor"
            mConsulta = mConsulta & " FROM acumuladores_todos"
            If pSoloFavoritos Then
                mConsulta = mConsulta & " WHERE Favoritos > 0"
                mConsulta = mConsulta & " ORDER BY Favoritos"
            Else
                mConsulta = mConsulta & " ORDER BY DESCRIPCION"
            End If

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Acumuladores_Usuario", ex, mConsulta)
        End Try

    End Function




    Public Function Lista_AutorizadoJustificar(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = Nothing, Optional ByVal pFechaHasta As String = Nothing) As Object Implements PresenciaDAO.Lista_AutorizadoJustificar
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object


        Try
            mConsulta = "SELECT *"
            mConsulta = mConsulta & " FROM AutorizadoJustificar "
            If pDNI <> "" Then
                mWhere = mWhere & " WHERE DNI = '" & pDNI & "'"
            End If
            If pFechaDesde Is Nothing Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " FECHA_DESDE <= sysdate"
            Else
                If pFechaDesde <> "" Then
                    If mWhere <> "" Then
                        mWhere = mWhere & " AND "
                    Else
                        mWhere = mWhere & " WHERE "
                    End If
                    mWhere = mWhere & " FECHA_DESDE <= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
                End If
            End If

            If pFechaHasta Is Nothing Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " (FECHA_HASTA >= sysdate OR FECHA_HASTA IS NULL)"
            Else
                If pFechaHasta <> "" Then
                    If mWhere <> "" Then
                        mWhere = mWhere & " AND "
                    Else
                        mWhere = mWhere & " WHERE "
                    End If
                    mWhere = mWhere & " FECHA_HASTA >= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
                End If
            End If


            mCommand.CommandText = mConsulta & mWhere
            mCommand.Connection = mConexion
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_AutorizadoConsultar(Optional ByVal pDNI As String = "") As Object Implements PresenciaDAO.Lista_AutorizadoConsultar
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object


        Try
            mConsulta = "SELECT count(*) total"
            mConsulta = mConsulta & " FROM AutorizadoConsultar "
            If pDNI <> "" Then
                mWhere = mWhere & " WHERE ID_RESPONSABLE = '" & pDNI & "'"
            End If
            

            mCommand.CommandText = mConsulta & mWhere
            mCommand.Connection = mConexion
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_AutorizadoAprobar(Optional ByVal pDNI As String = "") As Object Implements PresenciaDAO.Lista_AutorizadoAprobar
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object


        Try
            mConsulta = "SELECT count(*) total"
            mConsulta = mConsulta & " FROM Asignacion_responsable "
            If pDNI <> "" Then
                mWhere = mWhere & " WHERE ID_RESPONSABLE = '" & pDNI & "'"
            End If


            mCommand.CommandText = mConsulta & mWhere
            mCommand.Connection = mConexion
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar", ex, mConsulta)
        End Try

    End Function

    Public Function Usuario_ResponsableSolicitud(Optional ByVal pUsuario As String = "", Optional ByVal pSolicitud As String = "") As Object Implements PresenciaDAO.Usuario_ResponsableSolicitud
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object


        Try
            mConsulta = "select count(*) total from siguientes_solicitud where cod_solicitud = " & pSolicitud & " and dni ='" & pUsuario & "'"

            


            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar", ex, mConsulta)
        End Try

    End Function


    Public Overloads Function Lista_Solicitudes_de_Grupos(Optional ByVal pDni As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodIncidencia As String = "", Optional ByVal pOrden As String = "Usuario") As Object Implements PresenciaDAO.Lista_Solicitudes_de_Grupos
        Dim mConsulta As String
        Dim mOrden As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object
        Dim contador As Integer
        Try

            Dim mGruposAutorizados As String

            'Dim mBD As PresenciaDAO
            'mBD = DAOFactory.GetFactory(CTE_Tipo_BD).getPresenciaDAO(CTE_Cadena_conexion)
            'mBD.Conecta()

            contador = NumeroDeGrupos(pListaGrupos)


            mConsulta = "SELECT codigo, fecha, estado, solicitud.DNI , solicitud.DNI || ' ' ||  ape1 || ' ' || ape2 || ' ' || nombre Nombre,  "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM solicitud, incidencias,empleados"
            If pCodIncidencia <> "" Then
                mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia"
                mConsulta = mConsulta & " AND incidencias.cod_incidencia =  " & pCodIncidencia
            Else
                mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            End If
            mConsulta = mConsulta & " AND SOLICITUD.dni = EMPLEADOS.dni"
            If pDni <> "" Then
                If Left(pDni, 1) = "(" Then
                    mConsulta = mConsulta & " AND SOLICITUD.dni in " & pDni
                Else
                    mConsulta = mConsulta & " AND SOLICITUD.dni = '" & pDni & "'"
                End If
            End If
            If pListaGrupos <> "" And pListaGrupos <> "Todos" Then
                mConsulta = mConsulta & " AND solicitud.dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni) " 'having count(distinct cod_grupo) = " & contador & ")"
            End If

            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If

            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            'If pOrden <> "" Then
            '    mConsulta = mConsulta & " order by " & pOrden & ",fecha, desde, solicitud.cod_incidencia"
            'Else
            '    mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            'End If
            If pOrden = "Usuario" Then
                mOrden = " ORDER BY ape1,ape2, nombre, fecha, desde, codigo"
            ElseIf pOrden = "Fecha" Then
                mOrden = " ORDER BY fecha, desde, codigo"
            ElseIf pOrden = "Incidencia" Then
                mOrden = " ORDER BY cod_incidencia,ape1,ape2, nombre, fecha, desde"
            ElseIf pOrden = "Desde" Then
                mOrden = " ORDER BY desde "
            ElseIf pOrden = "Hasta" Then
                mOrden = " ORDER BY hasta "
            End If

            mCommand.CommandText = mConsulta & mOrden
            mCommand.Connection = mConexion
            mreader = mCommand.ExecuteReader()
            Return mreader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_de_Grupos", ex, mConsulta & mOrden)
        End Try

    End Function

    Public Function Lista_Solicitudes_de_Grupos_Rownum(ByRef pDatos As DataSet, ByVal pListaGrupos As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pListaEstados As String, Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Solicitudes_de_Grupos_Rownum
        Dim mConsulta As String
        Dim mWhere As String
        Dim contador As Integer
        Try

            Dim mGruposAutorizados As String

            'Dim mBD As PresenciaDAO
            'mBD = DAOFactory.GetFactory(CTE_Tipo_BD).getPresenciaDAO(CTE_Cadena_conexion)
            'mBD.Conecta()

            contador = NumeroDeGrupos(pListaGrupos)

            mConsulta &= "SELECT codigo, fecha, estado, solicitud.DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            If pListaGrupos <> "" And pListaGrupos <> "Todos" Then
                mConsulta = mConsulta & " AND solicitud.dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni)"
            End If

            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If

            If pOrden <> "" Then
                mConsulta = mConsulta & " order by " & pOrden & ",fecha, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, solicitud.cod_incidencia"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_de_Grupos", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Empleados_en_Grupos(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal pApellidos As String = "") As Object Implements PresenciaDAO.Busqueda_Empleados_en_Grupos
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim contador As Integer

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pListaGrupos <> "" Then
                If pListaGrupos <> "Todos" Then
                    'porque si son Todos, no tiene que seleccionar por grupos
                    If mWhere <> "" Then
                        mWhere = mWhere & " AND "
                    Else
                        mWhere = " WHERE "
                    End If
                    contador = NumeroDeGrupos(pListaGrupos)
                    'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                    'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                    mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni) "
                End If
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Empleados_en_Asignaciones(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Object Implements PresenciaDAO.Busqueda_Empleados_en_Asignaciones
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim contador As Integer

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo from empleados,asignacion_responsable"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " asignacion_responsable.id_responsable = '" & pResponsable & "' and asignacion_responsable.id_usuario = empleados.dni"
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
        End Try

    End Function

    Private Function NumeroDeGrupos(ByVal pListaGrupos As String)
        Dim lista As String
        Dim listaAux As String
        Dim contador As Integer
        listaAux = Val(Left(pListaGrupos, 4)) & ","
        If Len(pListaGrupos) < 5 Then
            lista = ""
        Else
            lista = Right(pListaGrupos, Len(pListaGrupos) - 5)
        End If
        contador = 1
        While Len(lista) > 0
            listaAux = listaAux & Val(Left(lista, 4)) & ","
            contador = contador + 1
            If Len(lista) > 4 Then
                lista = Right(lista, Len(lista) - 5)
            Else
                lista = ""
            End If
        End While
        If Right(listaAux, 1) = "," Then
            listaAux = Left(listaAux, Len(listaAux) - 1)
        End If
        Return contador
    End Function



    Public Function Lista_Delegados(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "", Optional ByVal pOrden As String = "Ape1,Ape2,Nombre") As Object Implements PresenciaDAO.Lista_Delegados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT Delegados.ID_Responsable, Delegados.ID_Delegado, Empleados.Ape1, Empleados.Ape2, Empleados.Nombre "
            mConsulta = mConsulta & " FROM Delegados, Empleados"
            mWhere = mWhere & " WHERE Delegados.ID_Delegado = Empleados.Dni"
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            If pID_Delegado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = mWhere & " WHERE "
                End If
                mWhere = mWhere & " ID_DELEGADO = '" & pID_Delegado & "'"
            End If
            mCommand.Connection = mConexion
            If pOrden <> "" Then
                mCommand.CommandText = mConsulta & mWhere & " Order by " & pOrden
            Else
                mCommand.CommandText = mConsulta & mWhere
            End If

            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Delegados", ex, mConsulta)
        End Try
    End Function


    Public Function Inserta_Delegados(ByVal pID_Responsable As String, ByVal pID_Delegado As String) As Boolean Implements PresenciaDAO.Inserta_Delegados
        'inserta un Delegado

        Dim mConsulta As String
        Try
            mConsulta = "INSERT INTO Delegados(ID_RESPONSABLE, ID_DELEGADO)"
            mConsulta = mConsulta & " VALUES('" & pID_Responsable & "','" & pID_Delegado & "')"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Delegados", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Delegados(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "") As Object Implements PresenciaDAO.Elimina_Delegados
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE Delegados "
            If pID_Responsable <> "" Then
                mWhere = " WHERE ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            If pID_Delegado <> "" Then
                If mWhere <> "" Then
                    mWhere = " WHERE"
                Else
                    mWhere = " AND"
                End If
                mWhere = mWhere & " ID_Delegado = '" & pID_Delegado & "'"
            End If

            mConsulta = mConsulta & mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
        Catch ex As Exception
            Trata_Error("Error en Elimina_Delegados", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Grupos_Consulta(Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Object Implements PresenciaDAO.Lista_Grupos_Consulta
        Dim mSQL As String
        Dim mWhere As String
        Try
            mSQL = "SELECT cod_grupo, desc_grupo, grupo_padre"
            mSQL = mSQL & " FROM gruposconsulta "
            If pCodigo > 0 Then
                mWhere = " WHERE cod_grupo = " & pCodigo
            End If
            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "Desc_grupo LIKE '" & pNombre & "'"
            End If
            If pPadre > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "Grupo_Padre = " & pPadre
            End If

            If mWhere <> "" Then
                mSQL = mSQL & mWhere
            End If
            mSQL = mSQL & " order by desc_grupo"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader
            Me.Haz_Log("Ejecutando SQL:" & mSQL, 3)
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta", ex, mSQL)
        End Try
    End Function

    '##########################################
    Public Function Inserta_Evento(ByVal pDNI As String, ByVal pFecha As DateTime, ByVal pSentido As String, ByVal pCod_recurso As Long, Optional ByVal pTarjeta As String = "", Optional ByVal pCod_incidencia As Long = 0, Optional ByVal pPermitido As Boolean = False, Optional ByVal pIP As String = "") As Boolean Implements PresenciaDAO.Inserta_Evento
        'inserta un evento
        Dim mSQL As String
        Dim mCodigo As Long

        Try
            'mSQL = "SELECT seqeventos.nextval FROM dual"
            Dim mCommand As New OleDb.OleDbCommand
            'Dim mReader As Object
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            'mReader = mCommand.ExecuteReader
            'If mReader.Read Then
            'mCodigo = NVL(mReader(0), 0)
            'End If
            'mReader.Close()
            'mCodigo = mCodigo + 1

            mSQL = "INSERT INTO EVENTOS(cod_evento,dni_empl,fecha,hora,e_s,tipoevento,estadoterminal,permitido,cod_incidencia,cod_recurso, ip)"
            mSQL = mSQL & " VALUES(" & 1 & ",'" & pDNI & "','" & pFecha.ToString("dd/MM/yyyy") & "','" & pFecha.ToString("HH:mm") & "',"
            If pSentido = "S" Then
                mSQL = mSQL & "'S'"
            Else
                mSQL = mSQL & "'E'"
            End If
            mSQL = mSQL & ",2,'ON',"
            If pPermitido Then
                mSQL = mSQL & "'S',"
            Else
                mSQL = mSQL & "'N',"
            End If
            If pCod_incidencia > 0 Then
                mSQL = mSQL & pCod_incidencia
            Else
                mSQL = mSQL & "NULL"
            End If
            mSQL = mSQL & "," & pCod_recurso & ","
            If pIP <> "" Then
                mSQL = mSQL & "'" & Left(pIP, 15) & "')"
            Else
                mSQL = mSQL & "NULL)"
            End If

            mCommand.CommandText = mSQL
            Me.Haz_Log("Ejecutando SQL:" & mSQL, 3)
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Inserta_Evento", ex, mSQL)
        End Try

    End Function

    Public Function Elimina_Evento(ByVal pCod_Evento As String) As Boolean Implements PresenciaDAO.Elimina_Evento
        Dim mConsulta As String
        Dim mWhere As String

        Dim mDataset As New DataSet
        Dim oAdapter As OleDb.OleDbDataAdapter
        Dim mDia As String = ""
        Dim mDNI As String

        Try

            oAdapter = New OleDb.OleDbDataAdapter("select * from eventos where COD_EVENTO in (" & pCod_Evento & ")", mconexion)
            oAdapter.Fill(mDataset)

            If mDataset.Tables(0).Rows.Count <> 0 Then
                'aunque vayan varios eventos, el dia y el dni sera el mismo.
                'asi que cogemos el de la primera fila
                mDia = NVL(mDataset.Tables(0).Rows(0)("fecha"), Today.ToString("dd-MM-yyyy"))
                mDNI = mDataset.Tables(0).Rows(0)("dni_empl")
                mDataset.Clear()
            End If
            mDataset = Nothing


            'borramos el diario del dia de la fecha
            If mDia <> "" Then
                Dim mCommand_Borrar_Diario As New OleDb.OleDbCommand
                mCommand_Borrar_Diario.Connection = mConexion
                mCommand_Borrar_Diario.CommandText = "delete diario where dni ='" & mDNI & "' and " & " fecha ='" & mDia & "'"
                mCommand_Borrar_Diario.ExecuteNonQuery()
            End If


            mConsulta = "DELETE Eventos "
            mWhere = " WHERE COD_EVENTO IN (" & pCod_Evento & ")"

            mConsulta = mConsulta & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Evento", ex, mConsulta)
        End Try
    End Function

    Public Function Lista_SolicitudesAprobadas(ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Object Implements PresenciaDAO.Lista_SolicitudesAprobadas
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrden As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM Solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mWhere = mWhere & " and ULTIMO_RESPONSABLE IN ('" & pLista_Responsables & "')"
            mWhere = mWhere & " and (ULTIMO_RESPONSABLE <>  ID_SIGUIENTE_RESPONSABLE or ID_SIGUIENTE_RESPONSABLE is null)"
            mWhere = mWhere & " and Estado IN ('A','E')"
            If pOrdenUsuario Then
                mOrden = " ORDER BY DNI, fecha, desde, codigo"
            ElseIf pOrdenFechasol Then
                mOrden = " ORDER BY fecha_sol, fecha, desde, codigo"
            Else
                mOrden = " ORDER BY fecha, desde, codigo"
            End If
            mCommand.CommandText = mConsulta & mWhere & mOrden
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()

            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_SolicitudesAprobadas", ex, mConsulta & mWhere & mOrden)
        End Try

    End Function

    Public Function Lista_SolicitudesAprobadas_Grupos(ByVal pLista_Responsables As String, Optional ByVal pOrden As String = "Usuario", Optional ByVal pEstado As String = "A") As Object Implements PresenciaDAO.Lista_SolicitudesAprobadas_Grupos
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrden As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, solicitud.DNI , solicitud.DNI || ' ' ||  ape1 || ' ' || ape2 || ' ' || nombre Nombre,  "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM Solicitud, incidencias,empleados"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND SOLICITUD.dni = EMPLEADOS.dni"
            mWhere = mWhere & " and ULTIMO_RESPONSABLE IN ('" & pLista_Responsables & "')"
            'mWhere = mWhere & " and ((ULTIMO_RESPONSABLE not in (select dni from siguientes_solicitud where cod_solicitud = solicitud.codigo)) or not exists (select * from siguientes_solicitud where cod_Solicitud = solicitud.codigo))"
            mWhere = mWhere & " and Estado ='" & pEstado & "'"
            If pOrden = "Usuario" Then
                mOrden = " ORDER BY ape1,ape2, nombre, fecha, desde, codigo"
            ElseIf pOrden = "Fecha" Then
                mOrden = " ORDER BY fecha, desde, codigo"
            ElseIf pOrden = "Incidencia" Then
                mOrden = " ORDER BY cod_incidencia,ape1,ape2, nombre, fecha, desde"
            ElseIf pOrden = "Desde" Then
                mOrden = " ORDER BY desde "
            ElseIf pOrden = "Hasta" Then
                mOrden = " ORDER BY hasta "
            End If
            mCommand.CommandText = mConsulta & mWhere & mOrden
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()

            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_SolicitudesAprobadas", ex, mConsulta & mWhere & mOrden)
        End Try

    End Function

    Public Function Lista_SolicitudesAprobadas_Movimiento(ByRef pDatos As DataSet, ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean Implements PresenciaDAO.Lista_SolicitudesAprobadas_Movimiento
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrden As String

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM Solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mWhere = mWhere & " and ULTIMO_RESPONSABLE IN ('" & pLista_Responsables & "')"
            mWhere = mWhere & " and (ULTIMO_RESPONSABLE <>  ID_SIGUIENTE_RESPONSABLE or ID_SIGUIENTE_RESPONSABLE is null)"
            mWhere = mWhere & " and Estado IN ('A','E')"
            If pOrdenUsuario Then
                mOrden = " ORDER BY DNI, fecha, desde, codigo"
            ElseIf pOrdenFechasol Then
                mOrden = " ORDER BY fecha_sol, fecha, desde, codigo"
            Else
                mOrden = " ORDER BY fecha, desde, codigo"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta & mWhere & mOrden, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_SolicitudesAprobadas", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_SolicitudesAprobadas_Movimiento_Grupos(ByRef pDatos As DataSet, ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean Implements PresenciaDAO.Lista_SolicitudesAprobadas_Movimiento_Grupos
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrden As String

        Try
            mConsulta = "SELECT codigo, fecha, estado, DNI, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo"
            mConsulta = mConsulta & " FROM Solicitud, incidencias"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mWhere = mWhere & " and ULTIMO_RESPONSABLE IN ('" & pLista_Responsables & "')"
            'mWhere = mWhere & " and (ULTIMO_RESPONSABLE <>  ID_SIGUIENTE_RESPONSABLE or ID_SIGUIENTE_RESPONSABLE is null)"
            'mWhere = mWhere & " and ((ULTIMO_RESPONSABLE not in (select dni from siguientes_solicitud where cod_solicitud = solicitud.codigo)) or not exists (select * from siguientes_solicitud where cod_Solicitud = solicitud.codigo))"
            mWhere = mWhere & " and Estado ='A'"
            If pOrdenUsuario Then
                mOrden = " ORDER BY DNI, fecha, desde, codigo"
            ElseIf pOrdenFechasol Then
                mOrden = " ORDER BY fecha_sol, fecha, desde, codigo"
            Else
                mOrden = " ORDER BY fecha, desde, codigo"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta & mWhere & mOrden, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_SolicitudesAprobadas", ex, mConsulta)
        End Try

    End Function

    Public Function CambiaSentido_Evento(ByVal pCod_Evento As String, ByVal pSentido As String) As Boolean Implements PresenciaDAO.CambiaSentido_Evento
        Dim mSQL As String
        Dim mWhere As String

        Try
            mSQL = "UPDATE Eventos set E_S='" & pSentido & "', MARCADO='N'"
            mWhere = " WHERE COD_EVENTO IN (" & pCod_Evento & ")"

            mSQL = mSQL & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en CambiaSentido_Evento", ex, mSQL)
        End Try

    End Function

    '#######################
    Public Function Valor_Acumulador(ByVal pDNI As String, ByVal pFecha As Date, ByVal pAcumulador As String) As String Implements PresenciaDAO.Valor_Acumulador
        Dim mSQL As String

        Try
            mSQL = "SELECT Valor_acumulador_todosf('" & pDNI & "','" & pAcumulador & "',to_date('" & pFecha.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')) from dual"

            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim mSalida As String
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                mSalida = NVL(mReader(0), "")
            End If
            mReader.Close()

            If mSalida = "NO DISP" Then mSalida = "0"
            Return mSalida
        Catch ex As Exception
            Trata_Error("Error en Valor_acumulador", ex, mSQL)
        End Try
    End Function

    Public Function Ahora() As Date Implements PresenciaDAO.Ahora
        'coge la fecha de la base de datos
        'a traves de la funcion mysysdate (definida para cada usuario)

        Dim mConsulta As String
        Dim mSalida As DateTime

        mConsulta = "SELECT mysysdate FROM dual"

        Dim mCommand As New OleDb.OleDbCommand
        Dim mreader As Object

        mCommand.Connection = mConexion
        mCommand.CommandText = mConsulta
        mreader = mCommand.ExecuteReader()
        mreader.Read()
        mSalida = NVL(mreader(0), Now)
        mreader.Close()

        Return mSalida

    End Function

    '##################################
    Public Function Lista_Grupo_Trabajo(ByRef pDatos As DataSet, Optional ByVal pCodigo As Integer = 0, Optional ByVal pDescGrupo As String = "", Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Grupo_Trabajo
        Dim mSQL As String

        Try
            mSQL = "SELECT distinct(COD_GRUPOTRABAJO), DESC_GRUPOTRABAJO, '' total  from GRUPOTRABAJO " ',(select distinct(cod_grupo),count(*) total from pertenecena group by cod_grupo) p"
            If pCodigo > 0 Then
                mSQL = mSQL & " WHERE COD_GRUPOTRABAJO = " & pCodigo
                If pDescGrupo <> "" Then
                    mSQL = mSQL & " AND SUPR_ACCENT(UPPER(DESC_GRUPOTRABAJO)) LIKE  SUPR_ACCENT(UPPER('" & pDescGrupo & "%'))"
                End If
            ElseIf pDescGrupo <> "" Then
                mSQL = mSQL & " WHERE SUPR_ACCENT(UPPER(DESC_GRUPOTRABAJO)) LIKE SUPR_ACCENT(UPPER('" & pDescGrupo & "%'))"
                'mSQL = mSQL & " AND COD_GRUPOTRABAJO = cod_grupo"
            ElseIf pDescGrupo = "" Then
                'mSQL = mSQL & " WHERE COD_GRUPOTRABAJO = cod_grupo"
            End If

            If pOrden <> "" Then
                mSQL = mSQL & " ORDER BY " & pOrden
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Grupo_Trabajo", ex, mSQL)
        End Try
    End Function


    '################################
    Public Function Lista_Asocia_Usuario_Grupo_Trabajo(Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0) As Object Implements PresenciaDAO.Lista_Asocia_Usuario_Grupo_Trabajo
        Dim mSQL As String
        Dim mWHERE As String
        Try
            mSQL = "SELECT COD_ASOC, DNI_EMPL, COD_GRUPOTRABAJO, FECHA_DESDE, FECHA_HASTA from ASOCIAUSUARIOGRUPOTRABAJO"
            If pCod_Asoc > 0 Then
                mWHERE = " COD_ASOC = " & pCod_Asoc
            End If
            If pDNI <> "" Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " DNI_EMPL = '" & pDNI & "'"
            End If
            If pCodigo > 0 Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " COD_GRUPOTRABAJO = " & pCodigo
            End If
            If mWHERE <> "" Then
                mSQL = mSQL & " WHERE " & mWHERE
            End If
            mSQL = mSQL & " ORDER BY dni_empl, fecha_desde"

            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Asocia_Usuario_Grupo_Trabajo", ex, mSQL)
        End Try
    End Function

    '##############################
    Public Function Elimina_Asocia_Usuario_Grupo_Trabajo(ByVal pCod_Asoc As Integer) As Boolean Implements PresenciaDAO.Elimina_Asocia_Usuario_Grupo_Trabajo
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE ASOCIAUSUARIOGRUPOTRABAJO "
            mWhere = " WHERE COD_ASOC = " & pCod_Asoc

            mConsulta = mConsulta & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asocia_Usuario_Grupo_Trabajo", ex, mConsulta)
        End Try
    End Function

    '#############################
    Public Function Actualiza_Asocia_Usuario_Grupo_Trabajo(ByVal pCod_Asoc As Integer, ByVal pDNI As String, ByVal pCodigo As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean Implements PresenciaDAO.Actualiza_Asocia_Usuario_Grupo_Trabajo
        Dim mConsulta As String
        Dim mSET As String
        Dim mWhere As String
        Try
            mConsulta = "UPDATE ASOCIAUSUARIOGRUPOTRABAJO "
            mSET = " DNI_EMPL = '" & pDNI & "'"
            mSET = mSET & ", COD_GRUPOTRABAJO = " & pCodigo
            mSET = mSET & ", FECHA_DESDE = TO_DATE('" & CDate(pFecha_Desde).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFecha_Hasta <> "" Then
                mSET = mSET & ", FECHA_HASTA = TO_DATE('" & CDate(pFecha_Hasta).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            Else
                mSET = mSET & ", FECHA_HASTA = NULL"
            End If
            mWhere = " COD_ASOC = " & pCod_Asoc

            mConsulta = mConsulta & " SET " & mSET & " WHERE " & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asocia_Usuario_Grupo_Trabajo", ex, mConsulta)
        End Try
    End Function

    '###################################
    Public Function Inserta_Asocia_Usuario_Grupo_Trabajo(ByVal pDNI As String, ByVal pCodigo As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean Implements PresenciaDAO.Inserta_Asocia_Usuario_Grupo_Trabajo
        Dim mConsulta As String
        Dim mVALUES As String
        Try
            mConsulta = "INSERT INTO ASOCIAUSUARIOGRUPOTRABAJO(COD_ASOC,DNI_EMPL,COD_GRUPOTRABAJO,FECHA_DESDE,FECHA_HASTA) "
            mVALUES = "0"
            mVALUES = mVALUES & "," & "'" & pDNI & "'"
            mVALUES = mVALUES & "," & pCodigo
            mVALUES = mVALUES & ",TO_DATE('" & CDate(pFecha_Desde).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFecha_Hasta <> "" Then
                mVALUES = mVALUES & ",TO_DATE('" & CDate(pFecha_Hasta).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            Else
                mVALUES = mVALUES & ",NULL"
            End If

            mConsulta = mConsulta & " VALUES(" & mVALUES & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asocia_Usuario_Grupo_Trabajo", ex, mConsulta)
        End Try
    End Function

    '##########################################
    Public Function Datos_Diario_DataSet(ByVal pDNI As String, ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, ByRef pSalida As System.Data.DataSet, Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Datos_Diario_DataSet

        Dim mfecha As Date
        Dim i As Integer
        Dim mConsulta As String

        Try

            'comprueba el diario
            mfecha = pFechaDesde
            For i = 0 To DateDiff(DateInterval.Day, pFechaDesde, pFechaHasta)
                mfecha = DateAdd(DateInterval.Day, i, pFechaDesde)
                Comprueba_Diario(pDNI, mfecha)
            Next

            mConsulta = "SELECT Dni, Fecha, Presencia, Justificaciones, Saldo, Cod_Horario from Diario"
            mConsulta = mConsulta & " WHERE dni = '" & pDNI & "'"
            mConsulta = mConsulta & "  and fecha >= to_Date('" & pFechaDesde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & "  and fecha <= to_Date('" & pFechaHasta.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & "  order by dni, fecha"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)
            pSalida = mDataSet
            Return True

        Catch ex As Exception
            Trata_Error("Error en Datos_Diario_Dataset", ex, mConsulta)
            pError = ex.Message
            Return False
        End Try

    End Function


    Public Function Lista_Intervalos_Horario(ByVal pcod_horario As String) As DataSet Implements PresenciaDAO.Lista_Intervalos_Horario

        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mSQL As String

        Try
            mCommand.Connection = mConexion
            mSQL = "SELECT formatea_hora(hora_inicio) Desde ,  formatea_hora(hora_inicio + total_minutos) Hasta FROM Intervalos"
            mSQL = mSQL & " WHERE cod_horario = " & pcod_horario
            mSQL = mSQL & " ORDER BY hora_inicio"
            mCommand.CommandText = mSQL
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            Dim mDataSet As New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mDataSet)

            Return mDataSet




            'mReader.Close()
        Catch ex As Exception
            Trata_Error("Error en Lista_Intervalos_Horario", ex, mSQL)
        End Try

    End Function

    Public Function Horario_Dia(ByVal pDni As String, ByVal pFecha_Desde As Date) As String Implements DAO.PresenciaDAO.Horario_Dia
        'Calcula el Horario asignado a una persona en un dia.
        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String
        Dim mCodigo As String

        Try
            mCommand.Connection = mConexion
            mSQL = "select scap_saldos.horario_dia_multiple('" & pDni & "',to_date('" & pFecha_Desde.ToString("dd/MM/yyyy") & "','DD/MM/YYYY')) from dual"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            mReader.Read()
            mCodigo = mReader(0)
            mReader.Close()
            Return mCodigo

        Catch ex As Exception
            Trata_Error("Error en Horario_Dia", ex, mSQL)
        End Try
    End Function

    Public Function Inserta_TVR_IP(ByVal pIP As String, Optional ByVal pDescripcion As String = "") As Boolean Implements DAO.PresenciaDAO.Inserta_TVR_IP
        Dim mConsulta As String
        Dim mVALUES As String
        Try
            mConsulta = "INSERT INTO TVR_IP(IP,DESCRIPCION) "
            mConsulta = mConsulta & " VALUES('" & pIP & "',"
            If pDescripcion <> "" Then
                mConsulta = mConsulta & "'" & pDescripcion & "'"
            Else
                mConsulta = mConsulta & "NULL"
            End If
            mConsulta = mConsulta & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Inserta_TVR_IP", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_TVR_IP(ByVal pIP As String) As Boolean Implements DAO.PresenciaDAO.Elimina_TVR_IP
        Dim mConsulta As String
        Dim mVALUES As String
        Try
            mConsulta = "DELETE TVR_IP "
            mConsulta = mConsulta & " WHERE IP = '" & pIP & "'"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_TVR_IP", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_TVR_IP(Optional ByVal pIP As String = "", Optional ByVal pOrdenDescripcion As Boolean = False) As System.Data.DataSet Implements DAO.PresenciaDAO.Lista_TVR_IP
        Dim mConsulta As String

        mConsulta = "SELECT IP,Descripcion FROM TVR_IP"
        If pIP <> "" Then
            mConsulta = mConsulta & " WHERE IP = '" & pIP & "'"
        End If
        If pOrdenDescripcion Then
            mConsulta = mConsulta & " ORDER BY Descripcion"
        Else
            mConsulta = mConsulta & " ORDER BY IP"
        End If

        Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
        Dim mDataSet As New DataSet
        'Conectarse, buscar datos y desconectarse de la base de datos 
        mDataAdapter.Fill(mDataSet)

        Return mDataSet

    End Function

    Public Function Modifica_TVR_IP(ByVal pIP As String, Optional ByVal pDescripcion As String = "") As Boolean Implements DAO.PresenciaDAO.Modifica_TVR_IP
        Dim mConsulta As String
        Dim mVALUES As String
        Try
            mConsulta = "UPDATE TVR_IP SET "
            mConsulta = mConsulta & " DESCRIPCION = '" & pDescripcion & "'"
            mConsulta = mConsulta & " WHERE IP = '" & pIP & "'"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Modifica_TVR_IP", ex, mConsulta)
            Return False
        End Try
    End Function


    '   Public Function seguridad(ByVal pIP As String) As Boolean Implements DAO.PresenciaDAO.seg
    '
    '       Dim mConsulta1 As String
    '      Dim mConsulta2 As String
    '     Dim mConsulta3 As String
    '    Dim xconsul As String
    '
    '       Dim mVALUES As String
    '      Try
    '         mConsulta1 = "select substr(ip,1,(instr(dni,'.',1,1)) -1) from TVR_IP where ip = '" & pIP & "'"
    '        mConsulta2 = "select substr(ip,(instr(dni,'.',1,1))+1,ABS(length(instr(ip,'.',1,1)-1)-(instr(ip,'.',1,2)-1 from TVR_IP where ip = '" & pIP & "'"""
    '       mConsulta3 = "select substr(ip,(instr(dni,'.',1,2))+1,ABS(length(instr(ip,'.',1,2)-1)-(instr(ip,'.',1,3)-1 from TVR_IP where ip = '" & pIP & "'"""
    '
    '           xconsul = "select concat(mconsulta1,'.',mconsulta2,'.',mconsulta3,'.') from dual"
    '
    '           Dim mCommand As New OleDb.OleDbCommand()
    '          mCommand.Connection = xconsul
    '         mCommand.CommandText = mConsulta1
    '        mCommand.ExecuteReader()
    '
    '           Return True
    '      Catch ex As Exception
    '         Trata_Error("Error en lista_control", ex, mConsulta)
    '        Return False
    '   End Try
    'End Function


    '******************************************************
    ' Public Function seguridad(ByVal IP As String, ByRef segmento1 As String, ByRef segmento2 As String, ByRef segmento3 As String, ByRef segmento4 As String) As Boolean Implements DAO.PresenciaDAO.seguridad
    '    Dim xconsul As String
    '
    '       Dim punto1 As Integer = InStr(ip, ".")
    '      Dim punto2 As Integer = InStr(punto1 + 1, ip, ".")
    '     Dim punto3 As Integer = InStr(punto2 + 1, ip, ".")
    '    Dim punto4 As Integer = InStr(punto3 + 1, ip, ".")
    '
    '       segmento1 = Left(ip, punto1 - 1)
    '      segmento2 = Mid(ip, punto1 + 1, punto2 - 1)
    '     segmento3 = Mid(ip, punto2 + 1, punto3 - 1)
    '    segmento4 = Right(ip, punto3 - 1)
    '
    '       Try
    '          ' mConsulta1 = "select substr(ip,1,(instr(dni,'.',1,1)) -1) from TVR_IP where ip = '" & pIP & "'"
    '         'mConsulta2 = "select substr(ip,(instr(dni,'.',1,1))+1,ABS(length(instr(ip,'.',1,1)-1)-(instr(ip,'.',1,2)-1 from TVR_IP where ip = '" & pIP & "'"""
    '        ' mConsulta3 = "select substr(ip,(instr(dni,'.',1,2))+1,ABS(length(instr(ip,'.',1,2)-1)-(instr(ip,'.',1,3)-1 from TVR_IP where ip = '" & pIP & "'"""
    '
    '           xconsul = "select concat(mconsulta1,'.',mconsulta2,'.',mconsulta3,'.') from tvrip"
    '
    '           Dim mCommand As New OleDb.OleDbCommand()
    '          mCommand.Connection = mConexion
    '         mCommand.CommandText = xconsul
    '        mCommand.ExecuteReader()
    '
    '           Return True
    ''      Catch ex As Exception
    '        Trata_Error("Error en lista_control", ex, xconsul)
    '       Return False
    '  End Try
    '
    '   End Function



    Public Function Lista_Justificaciones_Solicitud(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, ByVal pDNI As String, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pOrden As String = "", Optional ByVal pFechaOmitir As String = "") As Object Implements PresenciaDAO.Lista_Justificaciones_Solicitud
        'da el sumatorio de solicitudes + justificaciones por usuario e incidencia
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'mConsulta = "select a.dni_empl, sum(cuenta) numero"
            'mConsulta = mConsulta & " from ("
            'mConsulta = mConsulta & " select dni_empl, fecha_justificada, 1 cuenta"
            'mConsulta = mConsulta & " from justificaciones"
            'mConsulta = mConsulta & " where dni_empl = '" & pDNI & "'"
            'If pCod_Incidencia >= 0 Then
            '    mConsulta = mConsulta & " and cod_incidencia = 5"
            'End If
            'mConsulta = mConsulta & " AND fecha_justificada >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            'mConsulta = mConsulta & " AND fecha_justificada <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            'If pFechaOmitir <> "" Then
            '    If IsDate(pFechaOmitir) Then
            '        mConsulta = mConsulta & " AND fecha_justificada <> to_date('" & Format(CDate(pFechaOmitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            '    End If
            'End If
            'mConsulta = mConsulta & " group by dni_empl, fecha_justificada"
            'mConsulta = mConsulta & " union "
            'mConsulta = mConsulta & " select dni, fecha, 1"
            'mConsulta = mConsulta & " from solicitud"
            'mConsulta = mConsulta & " where dni = '" & pDNI & "'"
            'If pCod_Incidencia >= 0 Then
            '    mConsulta = mConsulta & " and cod_incidencia = 5"
            'End If
            ''solicitudes en curso o pendientes
            'mConsulta = mConsulta & " and estado in  ('E','P')"
            'mConsulta = mConsulta & " AND fecha >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            'mConsulta = mConsulta & " AND fecha <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            'If pFechaOmitir <> "" Then
            '    If IsDate(pFechaOmitir) Then
            '        mConsulta = mConsulta & " AND fecha <> to_date('" & Format(CDate(pFechaOmitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            '    End If
            'End If
            'mConsulta = mConsulta & " group by dni, fecha ) a"
            'mConsulta = mConsulta & " group by a.dni_empl"

            mConsulta = "select a.dni_empl, a.fecha_justificada, sum(cuenta) numero"
            mConsulta = mConsulta & " from ("


            mConsulta = mConsulta & " select dni_empl, fecha_justificada, 1 cuenta"
            mConsulta = mConsulta & " from justificaciones_t justificaciones"
            mConsulta = mConsulta & " where dni_empl = '" & pDNI & "'"
            If pCod_Incidencia >= 0 Then
                mConsulta = mConsulta & " and cod_incidencia = 5"
            End If
            mConsulta = mConsulta & " AND fecha_justificada >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " AND fecha_justificada <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFechaOmitir <> "" Then
                If IsDate(pFechaOmitir) Then
                    mConsulta = mConsulta & " AND fecha_justificada <> to_date('" & Format(CDate(pFechaOmitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
                End If
            End If
            mConsulta = mConsulta & " group by dni_empl, fecha_justificada"
            mConsulta = mConsulta & " union "
            mConsulta = mConsulta & " select dni, fecha, 1"
            mConsulta = mConsulta & " from solicitud"
            mConsulta = mConsulta & " where dni = '" & pDNI & "'"
            If pCod_Incidencia >= 0 Then
                mConsulta = mConsulta & " and cod_incidencia = 5"
            End If
            'solicitudes en curso o pendientes
            mConsulta = mConsulta & " and estado in  ('A','E','P')"
            mConsulta = mConsulta & " AND fecha >= to_date('" & Format(pFechaDesde, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            mConsulta = mConsulta & " AND fecha <= to_date('" & Format(pFechaHasta, "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFechaOmitir <> "" Then
                If IsDate(pFechaOmitir) Then
                    mConsulta = mConsulta & " AND fecha <> to_date('" & Format(CDate(pFechaOmitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
                End If
            End If
            mConsulta = mConsulta & " group by dni, fecha ) a"
            mConsulta = mConsulta & " group by a.dni_empl, a.fecha_justificada"


            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Justificaciones_Solicitud", ex, mConsulta)
        End Try

    End Function

    Function Numero_Solicitudes_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Integer Implements PresenciaDAO.Numero_Solicitudes_Dataset
        Dim mConsulta As String
        Dim mSolicitudes As New DataSet
        Dim numSolicitudes As Integer

        Try
            mConsulta = "select count(fecha) dias from ("
            mConsulta &= " select count(*) cuenta, fecha "
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "'"
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia
            mConsulta &= " and estado in  ('E','P','A')"
            mConsulta &= " and  fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY')"
            mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY')"
            mConsulta &= "  and fecha not in (select fecha_justificada  "
            mConsulta &= " from justificaciones_t J "
            mConsulta &= " where dni_empl = '" & pID_Usuario & "' "
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFecha_Desde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFecha_Hasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia & ") "
            mConsulta &= " group by  fecha)"



            Dim mDataAdapter2 As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter2.Fill(mSolicitudes)

            If mSolicitudes.Tables(0).Rows.Count > 0 Then
                If mSolicitudes.Tables(0).Rows(0)("dias") Is DBNull.Value Then
                    numSolicitudes = 0
                Else
                    numSolicitudes = mSolicitudes.Tables(0).Rows(0)("dias")
                End If
            Else
                numSolicitudes = 0
            End If

            mSolicitudes.Clear()
            mSolicitudes = Nothing

            Return numSolicitudes

        Catch ex As Exception
            Trata_Error("Error en Numero_Solicitudes_Dataset", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Justificaciones_Solicitud_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pFecha_Omitir As String = "", Optional ByVal pCodigo_Justificacion As String = "") As Integer Implements PresenciaDAO.Lista_Justificaciones_Solicitud_Dataset
        'Devuelve el numero de Justificaciones y solicitudes del usuario y del tipo de incidencia para el año especificado
        Dim mConsulta As String
        Dim mJustificaciones As New DataSet
        Dim mSolicitudes As New DataSet
        Dim numJustificaciones As Integer
        Dim numSolicitudes As Integer

        Try

            mConsulta = " select count(dia) dias  from "
            mConsulta &= " (select j.cod_incidencia codigo, fecha_justificada dia  "
            mConsulta &= " from justificaciones_t J "
            mConsulta &= " where dni_empl = '" & pID_Usuario & "' "
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFecha_Desde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFecha_Hasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia
            If pFecha_Omitir <> "" Then
                mConsulta &= " AND fecha_justificada <> to_date('" & Format(CDate(pFecha_Omitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            End If
            If pCodigo_Justificacion <> "" Then
                mConsulta &= " and cod_justificacion <> " & pCodigo_Justificacion
            End If
            mConsulta &= " group by  J.cod_incidencia, fecha_justificada "
            mConsulta &= " order by j.cod_incidencia  "
            mConsulta &= " ) group by codigo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mJustificaciones)

            If mJustificaciones.Tables(0).Rows.Count > 0 Then
                If mJustificaciones.Tables(0).Rows(0)("dias") Is DBNull.Value Then
                    numJustificaciones = 0
                Else
                    numJustificaciones = mJustificaciones.Tables(0).Rows(0)("dias")
                End If
            Else
                numJustificaciones = 0
            End If
            mJustificaciones.Clear()
            mJustificaciones = Nothing


            mConsulta = "select count(fecha) dias from ("
            mConsulta &= " select count(*) cuenta, fecha "
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "'"
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia
            mConsulta &= " and estado in  ('E','P','A')"
            mConsulta &= " and  fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY')"
            mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY')"
            mConsulta &= "  and fecha not in (select fecha_justificada  "
            mConsulta &= " from justificaciones_t J "
            mConsulta &= " where dni_empl = '" & pID_Usuario & "' "
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFecha_Desde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFecha_Hasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia & ") "
            If pFecha_Omitir <> "" Then
                mConsulta &= " AND fecha <> to_date('" & Format(CDate(pFecha_Omitir), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            End If
            mConsulta &= " group by  fecha)"



            Dim mDataAdapter2 As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter2.Fill(mSolicitudes)

            If mSolicitudes.Tables(0).Rows.Count > 0 Then
                If mSolicitudes.Tables(0).Rows(0)("dias") Is DBNull.Value Then
                    numSolicitudes = 0
                Else
                    numSolicitudes = mSolicitudes.Tables(0).Rows(0)("dias")
                End If
            Else
                numSolicitudes = 0
            End If

            mSolicitudes.Clear()
            mSolicitudes = Nothing

            Return numJustificaciones + numSolicitudes

        Catch ex As Exception
            Trata_Error("Error en Lista_Justificaciones_Solicitud_Dataset", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Justificaciones_Solicitud_Horas_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pCodigo_Solicitud As String = "", Optional ByVal pCodigo_Justificacion As String = "") As Integer Implements PresenciaDAO.Lista_Justificaciones_Solicitud_Horas_Dataset
        'Devuelve el numero de horas en Justificaciones y solicitudes del usuario y del tipo de incidencia para el año especificado
        Dim mConsulta As String
        Dim mJustificaciones As New DataSet
        Dim mSolicitudes As New DataSet
        Dim numJustificaciones As Integer
        Dim numSolicitudes As Integer

        Try

            mConsulta = "select sum(hasta_minutos - desde_minutos) minutos "
            mConsulta &= " from justificaciones_t J "
            mConsulta &= " where dni_empl = '" & pID_Usuario & "' "
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFecha_Desde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFecha_Hasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia
            If pCodigo_Justificacion <> "" Then
                mConsulta &= " and cod_justificacion <> " & pCodigo_Justificacion
            End If
            mConsulta &= " group by  J.cod_incidencia"



            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mJustificaciones)

            If mJustificaciones.Tables(0).Rows.Count > 0 Then
                If mJustificaciones.Tables(0).Rows(0)("minutos") Is DBNull.Value Then
                    numJustificaciones = 0
                Else
                    numJustificaciones = mJustificaciones.Tables(0).Rows(0)("minutos")
                End If
            Else
                numJustificaciones = 0
            End If
            mJustificaciones.Clear()
            mJustificaciones = Nothing



            mConsulta = " select sum(formatea_minuto(hasta) - formatea_minuto(desde)) minutos"
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "'"
            mConsulta &= " and cod_incidencia = " & pCodigoIncidencia
            mConsulta &= " and estado in  ('E','P','A')"
            mConsulta &= " and  fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY')"
            mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY')"
            'aqu si se cuenta el mismo dia, porque son horas solicitadas, y no dias.
            If pCodigo_Solicitud <> "" Then
                mConsulta &= " AND codigo <> " & pCodigo_Solicitud
            End If




            Dim mDataAdapter2 As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter2.Fill(mSolicitudes)

            If mSolicitudes.Tables(0).Rows.Count > 0 Then
                If mSolicitudes.Tables(0).Rows(0)("minutos") Is DBNull.Value Then
                    numSolicitudes = 0
                Else
                    numSolicitudes = mSolicitudes.Tables(0).Rows(0)("minutos")
                End If
            Else
                numSolicitudes = 0
            End If

            mSolicitudes.Clear()
            mSolicitudes = Nothing

            Return numJustificaciones + numSolicitudes

        Catch ex As Exception
            Trata_Error("Error en Lista_Justificaciones_Solicitud_Dataset", ex, mConsulta)
        End Try

    End Function


    Public Function Lista_Grupos_Privilegios(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Grupos_Privilegios
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM GRUPOSPRIVILEGIOS"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " WHERE COD_GRUPO = " & pCodigo
                If Not IsNothing(pNombre) And pNombre <> "" Then
                    mConsulta = mConsulta & " AND SUPR_ACCENT(UPPER(DESC_GRUPO)) like SUPR_ACCENT(UPPER('" & pNombre & "%'))"
                End If
            ElseIf Not IsNothing(pNombre) And pNombre <> "" Then
                mConsulta = mConsulta & " WHERE SUPR_ACCENT(UPPER(DESC_GRUPO)) like SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Privilegios", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Pertenecena(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pDNI As String = "", Optional ByVal pTipoGrp As Long = 0, Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Lista_Pertenecena
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM PERTENECENA"
            Dim mWhere As String
            If pCodigoGrupo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPO = " & pCodigoGrupo
            End If
            If pDNI <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " DNI_EMPL = '" & pDNI & "'"
            End If
            If pTipoGrp > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " TIPO_GRP = " & pTipoGrp
            End If
            mConsulta = mConsulta & " WHERE " & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Pertenecena", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_AccesosCalendarioGrupos(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pCodigoCalendario As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Lista_AccesosCalendarioGrupos
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM ACCESOSCALENDARIOGRUPOS"
            Dim mWhere As String
            If pCodigoGrupo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPO = " & pCodigoGrupo
            End If
            If pCodigoCalendario > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_CALENDARIO = " & pCodigoCalendario
            End If
            If pCodigoGrupoRecursos > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPORECURSOS = " & pCodigoGrupoRecursos
            End If
            mConsulta = mConsulta & " WHERE " & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AccesosCalendarioGrupos", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Calendarios(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pTipoCalendario As Long = 0, Optional ByVal pClaseCalendario As Long = 0, Optional ByVal pDescCalendario As String = "", Optional ByVal pAnio As String = "", Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Lista_Calendarios
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM CALENDARIOS"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_CALENDARIO = " & pCodigo
            End If
            If pTipoCalendario > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " TIPO_CALENDARIO = " & pTipoCalendario
            End If
            If pClaseCalendario > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " CLASE_CALENDARIO = " & pClaseCalendario
            End If

            If pDescCalendario <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " DESC_CALENDARIO LIKE '" & pDescCalendario & "'"
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Calendarios", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_CalendariosLaborables(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Lista_CalendariosLaborables
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM CALENDARIOSLABORABLES"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_CALENDARIO = " & pCodigo
            End If
            If mWhere <> "" Then
                mConsulta = mConsulta & " WHERE " & mWhere
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_CalendariosLaborables", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_DiasCalendario(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pFecha As String = "", Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_DiasCalendario
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM DIASCALENDARIO"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_CALENDARIO = " & pCodigo
            End If
            If IsDate(pFecha) Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " FECHA_CALENDARIO = TO_DATE('" & CDate(pFecha).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            End If
            mConsulta = mConsulta & " WHERE " & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_DiasCalendario", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Grupos_Recursos(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pCodigoPadre As Long = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Grupos_Recursos
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM GRUPORECURSOS"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPORECURSOS = " & pCodigo
            End If
            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " DESC_GRUPORECURSOS like '" & pNombre & "'"
            End If
            If pCodigoPadre > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " GRUPORECURSOS_PADRE = " & pCodigoPadre
            End If
            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Recursos", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Recursos(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Recursos
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM RECURSOS"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_RECURSO = " & pCodigo
            End If
            If pCodigoGrupoRecursos > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPORECURSOS = " & pCodigoGrupoRecursos
            End If
            mConsulta = mConsulta & " WHERE " & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Recursos", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_RecursosDeGrupo(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_RecursosDeGrupo
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM RECURSOS_DE_GRUPO"
            Dim mWhere As String
            If pCodigo > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_RECURSO = " & pCodigo
            End If
            If pCodigoGrupoRecursos > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPORECURSOS = " & pCodigoGrupoRecursos
            End If
            mConsulta = mConsulta & " WHERE " & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_RecursosDeGrupo", ex, mConsulta)
            Return False
        End Try
    End Function
    Function Lista_GrupoTrabajo_Calendarios(ByRef pDatos As DataSet, Optional ByVal pCodigo_Grupo As Integer = 0, Optional ByVal pTipo_Cal As Integer = 1) As Boolean Implements PresenciaDAO.Lista_GrupoTrabajo_Calendarios
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM ASOCIAGRUPOTRABAJOCALENDARIO A, CALENDARIOS C where a.COD_CALENDARIO = C.COD_CALENDARIO  "
            Dim mWhere As String
            If pCodigo_Grupo <> 0 Then
                mWhere = mWhere & " and A.COD_GRUPOTRABAJO = " & pCodigo_Grupo
            End If
            mWhere = mWhere & " and C.TIPO_CALENDARIO =" & pTipo_Cal
            mConsulta = mConsulta & mWhere
            mConsulta = mConsulta & " ORDER BY A.COD_CALENDARIO, A.FECHA_DESDE "
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AsociacionesGruposTrabajo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_AsociacionesGruposTrabajo(ByRef pDatos As System.Data.DataSet, Optional ByVal pID_Usuario As String = Nothing, Optional ByVal pCodGrupo As String = Nothing, Optional ByVal pAño As String = Nothing) As Boolean Implements PresenciaDAO.Lista_AsociacionesGruposTrabajo
        Dim mConsulta As String
        Try
            mConsulta = "SELECT A.COD_ASOC, DNI_EMPL, DNI_PROV, A.COD_GRUPOTRABAJO, FECHA_HASTA, FECHA_DESDE, B.DESC_GRUPOTRABAJO " & _
            " FROM ASOCIAUSUARIOGRUPOTRABAJO A, GRUPOTRABAJO B"
            Dim mWhere As String
            mWhere = mWhere & " A.COD_GRUPOTRABAJO = B.COD_GRUPOTRABAJO"
            If Not pID_Usuario Is Nothing Then
                mWhere = mWhere & " AND DNI_EMPL = '" & pID_Usuario & "'"
            End If
            If Not pCodGrupo Is Nothing Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " COD_GRUPOTRABAJO = " & pCodGrupo
            End If
            'If Not pAño Is Nothing Then
            '    If mWhere <> "" Then
            '        mWhere = mWhere & " AND "
            '    End If
            '    mWhere = mWhere & "(TO_CHAR(FECHA_DESDE,'YYYY') = '" & pAño & "' OR (FECHA_HASTA IS NULL AND TO_CHAR(FECHA_DESDE,'YYYY') < '" & pAño & "' ))"
            'End If
            mConsulta = mConsulta & " WHERE " & mWhere
            mConsulta = mConsulta & " ORDER BY COD_GRUPOTRABAJO, FECHA_DESDE "
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AsociacionesGruposTrabajo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Actualiza_Aprobacion(ByVal pCod_Solicitud As Long, ByVal pID_Responsable As String, ByVal pOperacion As String, Optional ByVal pDelegado As String = Nothing, Optional ByVal pCausaDenegacion As String = Nothing) As Boolean Implements PresenciaDAO.Actualiza_Aprobacion
        Dim mConsulta As String
        Dim mSET As String
        Dim mWhere As String
        Try
            mConsulta = "UPDATE APROBACION "

            If Not IsNothing(pDelegado) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                If pDelegado <> "" Then
                    mSET = mSET & " ID_DELEGADO = '" & pDelegado & "'"
                Else
                    mSET = mSET & " ID_DELEGADO = NULL"
                End If
            End If
            If Not IsNothing(pCausaDenegacion) Then
                If mSET = "" Then
                    mSET = "SET "
                Else
                    mSET = mSET & " , "
                End If
                If pCausaDenegacion <> "" Then
                    mSET = mSET & " Causa_Denegacion = '" & pCausaDenegacion & "'"
                Else
                    mSET = mSET & " Causa_Denegacion = NULL"
                End If
            End If
            mWhere = " ID_SOLICITUD = " & pCod_Solicitud
            mWhere = mWhere & " AND ID_RESPONSABLE = '" & pID_Responsable & "'"
            mWhere = mWhere & " AND OPERACION = '" & pOperacion & "'"

            If mSET <> "" Then
                mConsulta = mConsulta & mSET & " WHERE " & mWhere
                Dim mCommand As New OleDb.OleDbCommand
                mCommand.Connection = mConexion
                mCommand.CommandText = mConsulta
                mCommand.ExecuteNonQuery()
            End If

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Aprobacion", ex, mConsulta)
        End Try
    End Function

    Public Function Lista_Asocia_GrupoTrabajo_Calendario(ByRef pDatos As DataSet, Optional ByVal pCodGrupo As Integer = 0, Optional ByVal pCodCalendario As Integer = 0, Optional ByVal pDesde As String = "", Optional ByVal pHasta As String = "", Optional ByVal pFestivo As Integer = 0, Optional ByVal pAnyoFestivo As Integer = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Asocia_GrupoTrabajo_Calendario
        Dim mSQL As String
        Dim mWHERE As String
        Try
            mSQL = "select COD_GRUPOTRABAJO, asociagrupotrabajocalendario.COD_CALENDARIO,  FECHA_DESDE, FECHA_HASTA, calendarios.DESC_CALENDARIO, calendarios.ANIO"
            mSQL = mSQL & " from asociagrupotrabajocalendario ,calendarios"
            mWHERE = " asociagrupotrabajocalendario.cod_calendario = calendarios.cod_calendario"

            If pFestivo > 0 Then
                'If pFestivo = False Then
                mWHERE = mWHERE & " and tipo_calendario =" & pFestivo 'laborable
                'Else
                'mWHERE = mWHERE & " and tipo_calendario = 2 " 'festivo
                'If pAnyoFestivo > 0 Then
                'mWHERE = mWHERE & " and calendarios.ANIO = " & pAnyoFestivo
                'End If
                'End If
            End If
            'mWHERE = mWHERE & " and clase_calendario = 1 " 'de trabajo

            If pCodGrupo > 0 Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " COD_GRUPOTRABAJO = " & pCodGrupo
            End If
            If pCodCalendario > 0 Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " calendarios.COD_CALENDARIO =" & pCodCalendario
            End If
            If IsDate(pDesde) Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                If Not IsDate(pHasta) Then
                    mWHERE = mWHERE & " FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY')"
                Else
                    mWHERE = mWHERE & " ((FECHA_DESDE >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
                    & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null))"
                End If
                'mWHERE = mWHERE & " FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            End If
            'If IsDate(pHasta) Then
            'If mWHERE <> "" Then
            'mWHERE = mWHERE & " AND "
            'End If
            'mWHERE = mWHERE & " FECHA_HASTA >= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            'End If
            If mWHERE <> "" Then
                mSQL = mSQL & " WHERE " & mWHERE
            End If
            mSQL = mSQL & " ORDER BY Cod_grupotrabajo, fecha_desde"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Asocia_GrupoTrabajo_Calendario", ex, mSQL)
            Return False
        End Try

    End Function

    Public Overloads Function Lista_Asocia_Usuario_Grupo_Trabajo1(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Asocia_Usuario_Grupo_Trabajo
        Dim mSQL As String
        Dim mWHERE As String
        Try
            mSQL = "SELECT COD_ASOC, DNI_EMPL, COD_GRUPOTRABAJO, FECHA_DESDE, FECHA_HASTA from ASOCIAUSUARIOGRUPOTRABAJO"
            If pCod_Asoc > 0 Then
                mWHERE = " COD_ASOC = " & pCod_Asoc
            End If
            If pDNI <> "" Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " DNI_EMPL = '" & pDNI & "'"
            End If
            If pCodigo > 0 Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " COD_GRUPOTRABAJO = " & pCodigo
            End If
            If mWHERE <> "" Then
                mSQL = mSQL & " WHERE " & mWHERE
            End If
            mSQL = mSQL & " ORDER BY dni_empl, fecha_desde"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("Error en Lista_Asocia_Usuario_Grupo_Trabajo", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Lista_Asocia_Usuario_Grupo_Trabajo_Fechas(ByRef pDatos As DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Lista_Asocia_Usuario_Grupo_Trabajo_Fechas
        Dim mSQL As String
        Dim mWHERE As String
        Try
            mSQL = "SELECT distinct DNI_EMPL from ASOCIAUSUARIOGRUPOTRABAJO "
            mSQL &= " where ((FECHA_DESDE >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
                & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            If pCod_Asoc > 0 Then
                mWHERE = " COD_ASOC = " & pCod_Asoc
            End If
            If pDNI <> "" Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " DNI_EMPL = '" & pDNI & "'"
            End If
            If pCodigo > 0 Then
                If mWHERE <> "" Then
                    mWHERE = mWHERE & " AND "
                End If
                mWHERE = mWHERE & " COD_GRUPOTRABAJO = " & pCodigo
            End If
            If mWHERE <> "" Then
                mSQL = mSQL & "  " & mWHERE
            End If
            mSQL = mSQL & " ORDER BY dni_empl"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("Error en Lista_Asocia_Usuario_Grupo_Trabajo_Fechas", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Elimina_Diario(ByVal pDNI As String, ByVal pFechaDesde As String, Optional ByVal pFechaHasta As String = "", Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.Elimina_Diario
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE DIARIO "
            mWhere = " WHERE DNI = '" & pDNI & "'"
            mWhere = mWhere & " AND FECHA >= TO_DATE('" & Format(CDate(pFechaDesde), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            If IsDate(pFechaHasta) Then
                mWhere = mWhere & " AND FECHA <= TO_DATE('" & Format(CDate(pFechaHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')"
            End If

            mConsulta = mConsulta & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            pError = ""
            Return True

        Catch ex As Exception
            pError = ex.Message
            Trata_Error("Error en Elimina_Diario", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Actualiza_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_DesdeAnt As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean Implements PresenciaDAO.Actualiza_Asocia_Grupo_Trabajo_Calendario
        Dim mConsulta As String
        Dim mSET As String
        Dim mWhere As String
        Try
            mConsulta = "UPDATE ASOCIAGRUPOTRABAJOCALENDARIO "
            'mSET = " COD_CALENDARIO  = '" & pCodigoCalendario & "'"
            'mSET = mSET & ", COD_GRUPOTRABAJO = " & pCodigoGrupo
            mSET = " FECHA_DESDE = TO_DATE('" & CDate(pFecha_Desde).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFecha_Hasta <> "" Then
                mSET = mSET & ", FECHA_HASTA = TO_DATE('" & CDate(pFecha_Hasta).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            Else
                mSET = mSET & ", FECHA_HASTA = NULL"
            End If

            mWhere = mWhere & " COD_CALENDARIO  = " & pCodigoCalendario
            mWhere = mWhere & " AND COD_GRUPOTRABAJO = " & pCodigoGrupo
            mWhere = mWhere & " AND FECHA_DESDE = '" & pFecha_DesdeAnt & "'"

            mConsulta = mConsulta & " SET " & mSET & " WHERE " & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asocia_Grupo_Trabajo_Calendario", ex, mConsulta)
        End Try
    End Function

    Public Function Inserta_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean Implements PresenciaDAO.Inserta_Asocia_Grupo_Trabajo_Calendario
        Dim mConsulta As String
        Dim mVALUES As String
        Try
            mConsulta = "INSERT INTO ASOCIAGRUPOTRABAJOCALENDARIO(COD_GRUPOTRABAJO,COD_CALENDARIO,FECHA_DESDE,FECHA_HASTA) "
            mVALUES = pCodigoGrupo
            mVALUES = mVALUES & "," & pCodigoCalendario
            mVALUES = mVALUES & ",TO_DATE('" & CDate(pFecha_Desde).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            If pFecha_Hasta <> "" Then
                mVALUES = mVALUES & ",TO_DATE('" & CDate(pFecha_Hasta).ToString("dd/MM/yyyy") & "','DD/MM/YYYY')"
            Else
                mVALUES = mVALUES & ",NULL"
            End If

            mConsulta = mConsulta & " VALUES(" & mVALUES & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asocia_Grupo_Trabajo_Calendario", ex, mConsulta)
        End Try
    End Function

    Public Function Elimina_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_Desde As String) As Boolean Implements PresenciaDAO.Elimina_Asocia_Grupo_Trabajo_Calendario
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE ASOCIAGRUPOTRABAJOCALENDARIO "
            mWhere = mWhere & " COD_CALENDARIO  = " & pCodigoCalendario
            mWhere = mWhere & " AND COD_GRUPOTRABAJO = " & pCodigoGrupo
            mWhere = mWhere & " AND FECHA_DESDE = '" & pFecha_Desde & "'"


            mConsulta = mConsulta & " WHERE " & mWhere

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asocia_Grupo_Trabajo_Calendario", ex, mConsulta)
        End Try

    End Function

    Public Function Busqueda_Empleados_Grupos_Privilegios(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal Tipo_Grp As String = "", Optional ByVal pApellidos As String = "") As Object Implements PresenciaDAO.Busqueda_Empleados_Grupos_Privilegios
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim contador As Integer

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2 from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pListaGrupos <> "" Then

                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                contador = NumeroDeGrupos(pListaGrupos)
                mWhere = mWhere & " dni not IN (SELECT dni_empl FROM pertenecena WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) and tipo_grp=" & Tipo_Grp & ")"
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
        End Try

    End Function

    Public Function Inserta_Pertenecena(ByVal pCod_Grupo As Integer, ByVal pTipo_Grp As Integer, ByVal pDni_Empl As String) As Boolean Implements PresenciaDAO.Inserta_Pertenecena
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(codpertenece) FROM pertenecena"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into pertenecena (codpertenece, cod_grupo, tipo_grp, dni_empl) " _
                & " values (" & mauxdbl & "," & pCod_Grupo & "," & pTipo_Grp & ",'" & pDni_Empl & "')"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Pertenecena", ex, mConsulta)
        End Try

    End Function

    Public Function Elimina_Pertenecena(ByVal pCodpertenece As Integer) As Boolean Implements PresenciaDAO.Elimina_Pertenecena
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "delete from pertenecena where codpertenece = " & pCodpertenece
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()


            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Pertenecena", ex, mSQL)
        End Try
    End Function

    Public Function Actualiza_Grupos_Privilegios(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean Implements PresenciaDAO.Actualiza_Grupos_Privilegios
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "update gruposprivilegios set desc_grupo='" & pDesc_Grupo & "' where cod_grupo = " & pCod_Grupo
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Grupos_Privilegios", ex, mSQL)
        End Try
    End Function

    Public Function Inserta_Grupos_Privilegios(ByVal pDesc_Grupo As String) As Boolean Implements PresenciaDAO.Inserta_Grupos_Privilegios
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_grupo) FROM gruposprivilegios"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into gruposprivilegios (cod_grupo, desc_grupo) " _
                & " values (" & mauxdbl & ",'" & pDesc_Grupo & "')"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Grupos_Privilegios", ex, mConsulta)
        End Try


    End Function

    Public Function Elimina_Grupos_Privilegios(ByVal pCod_Grupo As Integer) As Boolean Implements PresenciaDAO.Elimina_Grupos_Privilegios
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT count(*) FROM pertenecena where cod_grupo=" & pCod_Grupo & " and tipo_grp=2"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM accesoscalendariogrupos where cod_grupo=" & pCod_Grupo
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            '***************************************************
            'borramos
            '***************************************************

            mSQL = "delete from gruposprivilegios where cod_grupo= " & pCod_Grupo
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Grupos_Privilegios", ex, mSQL)
        End Try

    End Function

    Public Function Elimina_Grupos_Consulta(ByVal pCod_Grupo As Integer) As Boolean Implements PresenciaDAO.Elimina_Grupos_Consulta
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT count(*) FROM pertenecena where cod_grupo=" & pCod_Grupo & " and tipo_grp=1"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM gruposconsulta where grupo_padre=" & pCod_Grupo
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            '***************************************************
            'borramos
            '***************************************************

            mSQL = "delete from gruposconsulta where cod_grupo= " & pCod_Grupo
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Grupos_Consultas", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Actualiza_Grupos_Consulta(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean Implements PresenciaDAO.Actualiza_Grupos_Consulta
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "update gruposconsulta set desc_grupo='" & pDesc_Grupo & "' where cod_grupo = " & pCod_Grupo
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Grupos_Consulta", ex, mSQL)
        End Try
    End Function

    Public Function Inserta_Grupos_Consulta(ByVal pDesc_Grupo As String, Optional ByVal pPadre As String = "") As Boolean Implements PresenciaDAO.Inserta_Grupos_Consulta
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_grupo) FROM gruposconsulta"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            Try
                If pPadre <> "" Then
                    If pPadre = mauxdbl Then pPadre = ""
                End If
            Catch ex As Exception
                pPadre = ""
            End Try

            mConsulta = "insert into gruposconsulta (cod_grupo, desc_grupo, grupo_padre) " _
                & " values (" & mauxdbl & ",'" & pDesc_Grupo & "'," & IIf(pPadre = "", "null", pPadre) & ")"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Grupos_Consulta", ex, mConsulta)
        End Try


    End Function

    Public Function Lista_Grupos_Recursos2(Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Object Implements PresenciaDAO.Lista_Grupos_Recursos2
        Dim mSQL As String
        Dim mWhere As String
        Try
            mSQL = "SELECT cod_gruporecursos, desc_gruporecursos, gruporecursos_padre"
            mSQL = mSQL & " FROM gruporecursos "
            If pCodigo > 0 Then
                mWhere = " WHERE cod_gruporecursos = " & pCodigo
            End If
            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "desc_gruporecursos LIKE '" & pNombre & "'"
            End If
            If pPadre > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "gruporecursos_padre = " & pPadre
            End If
            If mWhere <> "" Then
                mSQL = mSQL & mWhere
            End If
            mSQL = mSQL & " order by desc_gruporecursos"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Recursos2", ex, mSQL)
        End Try
    End Function

    Public Function Lista_Recursos_SinAsignar() As Object Implements PresenciaDAO.Lista_Recursos_SinAsignar
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM RECURSOS Where cod_gruporecursos is null"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Recursos_SinAsignar", ex, mConsulta)

        End Try

    End Function

    Public Function Elimina_Grupos_Recursos(ByVal pCod_Grupo As Integer) As Boolean Implements PresenciaDAO.Elimina_Grupos_Recursos
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT count(*) FROM recursos where cod_gruporecursos=" & pCod_Grupo
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM gruporecursos where gruporecursos_padre=" & pCod_Grupo
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            '***************************************************
            'borramos
            '***************************************************

            mSQL = "delete from gruporecursos where cod_gruporecursos= " & pCod_Grupo
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Grupos_Recursos", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_Grupos_Recursos(ByVal pDesc_Grupo As String, Optional ByVal pPadre As String = "") As Boolean Implements PresenciaDAO.Inserta_Grupos_Recursos
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_gruporecursos) FROM gruporecursos"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************"***********

            mConsulta = "insert into gruporecursos (cod_gruporecursos, desc_gruporecursos, gruporecursos_padre) " _
                & " values (" & mauxdbl & ",'" & pDesc_Grupo & "'," & IIf(pPadre = "", "null", pPadre) & ")"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Grupos_Recursos", ex, mConsulta)
        End Try

    End Function

    Public Function Actualiza_Grupos_Recursos(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean Implements PresenciaDAO.Actualiza_Grupos_Recursos
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "update gruporecursos set desc_gruporecursos='" & pDesc_Grupo & "' where cod_gruporecursos = " & pCod_Grupo
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Grupos_Recursos", ex, mSQL)
        End Try

    End Function

    Public Function Asigna_Recursos(ByVal pCod_Recurso As Integer, ByVal pPadre As String) As Boolean Implements PresenciaDAO.Asigna_Recursos
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand
        Try

            mSQL = "update recursos set cod_gruporecursos=" & pPadre & " where cod_recurso = " & pCod_Recurso
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Asigna_Recursos", ex, mSQL)
        End Try
    End Function




    Public Function Lee_Pertenecena(ByVal pCodigoGrupo As Long, ByVal pDNI As String, ByVal pTipoGrp As Long) As Boolean Implements PresenciaDAO.Lee_Pertenecena
        Dim mSQL As String
        Dim mWhere As String
        Try
            mSQL = "SELECT count(*) as valor"
            mSQL = mSQL & " FROM pertenecena p"
            mSQL = mSQL & " where p.cod_grupo = " & pCodigoGrupo & " And tipo_grp = " & pTipoGrp
            mSQL = mSQL & " and p.dni_empl = '" & pDNI & "'"


            Dim mReader As New DataSet
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)

            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(mReader)

            If mReader.Tables(0).Rows(0)("valor") > 0 Then
                mReader = Nothing
                Return True
            Else
                mReader = Nothing
                Return False
            End If
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Lista_Grupos_Consulta_Usuario(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As String = "", Optional ByVal pPadre As Long = 0, Optional ByVal pcoleccion As Collection = Nothing) As Boolean Implements PresenciaDAO.Lista_Grupos_Consulta_Usuario
        Dim mSQL As String
        Dim mWhere As String
        Try
            'mSQL = "SELECT g.cod_grupo, g.desc_grupo, g.grupo_padre, p.dni_empl"
            'mSQL = mSQL & " FROM gruposconsulta G, pertenecena p"
            'mSQL = mSQL & " where p.cod_grupo = g.cod_grupo and tipo_grp=1 "
            'mSQL = mSQL & " and p.dni_empl = '" & pCodigo & "'"

            mSQL = "select * from gruposconsulta"
            If pPadre > 0 Then
                mWhere = mWhere & " where Grupo_Padre = " & pPadre
            End If

            If Not IsNothing(pcoleccion) Then
                Dim cadena_Coleccion As String
                Dim i As Integer
                If pcoleccion.Count > 0 Then
                    For i = 1 To pcoleccion.Count
                        cadena_Coleccion = cadena_Coleccion & pcoleccion(i) & ","
                    Next
                    cadena_Coleccion = Left(cadena_Coleccion, Len(cadena_Coleccion) - 1)

                    If mWhere <> "" Then
                        mWhere = mWhere & " and "
                    Else
                        mWhere = mWhere & " where "
                    End If
                    mWhere = mWhere & " cod_grupo in (" & cadena_Coleccion & ")"
                End If
            End If
            If mWhere <> "" Then
                mSQL = mSQL & mWhere
            End If
            mSQL = mSQL & " order by desc_grupo"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)


            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta", ex, mSQL)
        End Try
    End Function

    Public Function Lista_AccesosCalendarioGrupos_Extendida(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pCodigoCalendario As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByVal pError As String = "") As Boolean Implements PresenciaDAO.Lista_AccesosCalendarioGrupos_Extendida
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM ACCESOSCALENDARIOGRUPOS A, calendarios C where c.tipo_calendario = 1 and A.cod_calendario = C.cod_calendario"
            Dim mWhere As String
            If pCodigoGrupo > 0 Then
                mWhere = mWhere & " AND COD_GRUPO = " & pCodigoGrupo
            End If
            If pCodigoCalendario > 0 Then
                mWhere = mWhere & " AND COD_CALENDARIO = " & pCodigoCalendario
            End If
            If pCodigoGrupoRecursos > 0 Then
                mWhere = mWhere & " AND COD_GRUPORECURSOS = " & pCodigoGrupoRecursos
            End If
            mConsulta = mConsulta & mWhere
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AccesosCalendarioGrupos_Extendida", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Elimina_AccesosCalendarioGrupos(ByVal pCod_Grupo As String, ByVal pCod_Calendario As Integer, ByVal pCod_Recurso As String) As Boolean Implements PresenciaDAO.Elimina_AccesosCalendarioGrupos
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "delete from AccesosCalendarioGrupos where cod_grupo=" & pCod_Grupo & " and cod_calendario = " & pCod_Calendario & " and cod_gruporecursos =" & pCod_Recurso
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_AccesosCalendarioGrupos", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_AccesosCalendarioGrupos(ByVal pCod_Grupo As String, ByVal pCod_Calendario As Integer, ByVal pCod_Recurso As String) As Boolean Implements PresenciaDAO.Inserta_AccesosCalendarioGrupos
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "insert into AccesosCalendarioGrupos (cod_grupo, cod_calendario, cod_gruporecursos) values (" & pCod_Grupo & "," & pCod_Calendario & "," & pCod_Recurso & ")"
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_AccesosCalendarioGrupos", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Lista_Calendarios_Asociacion(ByRef pDatos As System.Data.DataSet, ByVal pCod_Grupo As String, ByVal pCod_Recurso As String) As Boolean Implements PresenciaDAO.Lista_Calendarios_Asociacion
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM CALENDARIOS where tipo_calendario =1"
            mConsulta = mConsulta & " and cod_calendario not in "
            mConsulta = mConsulta & "(select cod_calendario from AccesosCalendarioGrupos where cod_grupo=" & pCod_Grupo & " and cod_gruporecursos=" & pCod_Recurso & ")"
            mConsulta = mConsulta & " order by cod_calendario"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Calendarios", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Incidencias_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDescripcion As String = "", Optional ByVal pCodigoIncidenciaDeCompensacion As Integer = -1) As Boolean Implements PresenciaDAO.Lista_Incidencias_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM incidencias"
            Dim mWhere As String
            If pCod_Incidencia <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_incidencia = " & pCod_Incidencia
            End If

            If pCodigoIncidenciaDeCompensacion <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_incidencia_compensacion = " & pCodigoIncidenciaDeCompensacion
            End If
            If pDescripcion <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " upper(desc_incidencia) like '" & UCase(pDescripcion) & "%'"
            End If
            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere
            mConsulta &= " order by cod_incidencia"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Incidencias_Dataset", ex, mConsulta)
            Return False
        End Try



    End Function

    Public Function Lista_Incidencias_Contrato_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDescripcion As String = "", Optional ByVal pCod_Contrato As Integer = -1) As Boolean Implements PresenciaDAO.Lista_Incidencias_Contrato_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT i.desc_incidencia, t.* FROM tipocontrato_incidencia t,incidencias i"
            Dim mWhere As String
            If pCod_Incidencia <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " t.cod_incidencia = " & pCod_Incidencia
            End If
            If pCod_Contrato <> -1 Then
                mWhere = mWhere & " and t.cod_tipocontrato = " & pCod_Contrato
            End If
            mWhere = mWhere & " And t.cod_incidencia=i.cod_incidencia"
            'If pDescripcion <> "" Then
            'If mWhere <> "" Then
            'mWhere = mWhere & " AND "
            'End If
            'mWhere = mWhere & " upper(desc_incidencia) like '" & UCase(pDescripcion) & "%'"
            'End If
            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere
            mConsulta &= " order by t.cod_incidencia"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Incidencias_Dataset", ex, mConsulta)
            Return False
        End Try



    End Function

    Public Function Actualiza_Incidencia(ByVal pCod_Incidencia As String, Optional ByVal pDesc_Incidencia As String = "", Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal pSeleccionable_TVR As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean Implements PresenciaDAO.Actualiza_Incidencia
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String

        Try

            mSQL = "update incidencias set "

            'if pdesc_incidencia
            'If mWhere <> "" Then mWhere &= ","

            'desc_gruporecursos='" & pDesc_Grupo & "' 

            mSQL &= "desc_incidencia = '" & pDesc_Incidencia & "'"
            If pTipo = "" Then
                mSQL &= ", tipo =null"
            Else
                mSQL &= ", tipo ='" & pTipo & "'"
            End If

            If pTipoFijo = "" Then
                mSQL &= ", tipofijo = null"
            Else
                mSQL &= ", tipofijo = '" & pTipoFijo & "'"
            End If

            If pMaximo = "" Then
                mSQL &= ", maximo = null"
            Else
                mSQL &= ", maximo = " & pMaximo
            End If

            If pFecha_Base = "" Then
                mSQL &= ", fecha_base = null"
            Else
                mSQL &= ", fecha_base = '" & pFecha_Base & "'"
            End If

            If pFecha_Termino = "" Then
                mSQL &= ", fecha_termino = null"
            Else
                mSQL &= ", fecha_termino = '" & pFecha_Termino & "'"
            End If

            If pTiempo_Maximo = "" Then
                mSQL &= ", tiempo_maximo = null"
            Else
                mSQL &= ", tiempo_maximo = " & pTiempo_Maximo
            End If

            If pOrden = "" Then
                mSQL &= ", orden = null"
            Else
                mSQL &= ", orden = " & pOrden
            End If

            If pGrupo = "" Then
                mSQL &= ", grupo = null"
            Else
                mSQL &= ", grupo = " & pGrupo
            End If

            If pSeleccionable = "" Then
                mSQL &= ", seleccionable = null"
            Else
                mSQL &= ", seleccionable = '" & pSeleccionable & "'"
            End If

            If pMaximo_Horas = "" Then
                mSQL &= ", maximo_horas = null"
            Else
                mSQL &= ", maximo_horas = " & pMaximo_Horas
            End If

            If pSeleccionable_TVR = "" Then
                mSQL &= ", seleccionable_tvr = null"
            Else
                mSQL &= ", seleccionable_tvr = '" & pSeleccionable_TVR & "'"
            End If

            If Maximo_Duracion <> 0 Then
                mSQL &= ", MAXIMO_DURACION = '" & Maximo_Duracion & "'"
            Else
                mSQL &= ", MAXIMO_DURACION = null"
            End If

            If Minimo_Duracion <> 0 Then
                mSQL &= ", MINIMO_DURACION = '" & Minimo_Duracion & "'"
            Else
                mSQL &= ", MINIMO_DURACION = null"
            End If

            If Tiempo_Minimo <> 0 Then
                mSQL &= ", TIEMPO_MINIMO = '" & Tiempo_Minimo & "'"
            Else
                mSQL &= ", TIEMPO_MINIMO = null"
            End If

            If Naturales = 1 Then
                mSQL &= ", NATURALES = 1"
            Else
                mSQL &= ", NATURALES = 0"
            End If


            mSQL &= " where cod_incidencia = " & pCod_Incidencia
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Incidencia", ex, mSQL)
        End Try


    End Function

    Public Function Elimina_Incidencia(ByVal pCod_Incidencia As String) As Boolean Implements PresenciaDAO.Elimina_Incidencia
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try


            Dim mCommand_valor As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim mauxdbl As Integer



            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT count(*) FROM eventos where cod_incidencia=" & pCod_Incidencia
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM esseleccionada where cod_incidencia=" & pCod_Incidencia
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM tipocontrato_incidencia where cod_incidencia=" & pCod_Incidencia
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "delete from incidencias "
            mSQL &= "where cod_incidencia = " & pCod_Incidencia

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Incidencia", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Inserta_Incidencia(ByVal pCod_Incidencia As String, ByVal pDesc_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal pSeleccionable_TVR As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Integer Implements PresenciaDAO.Inserta_Incidencia
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_incidencia) FROM incidencias"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into incidencias (cod_incidencia, desc_incidencia, tipo, tipofijo, maximo, fecha_base, fecha_termino, tiempo_maximo, orden, grupo, seleccionable, maximo_horas, seleccionable_tvr, MAXIMO_DURACION, MINIMO_DURACION, TIEMPO_MINIMO,NATURALES) " _
                & " values (" & mauxdbl & ","
            mConsulta &= "'" & pDesc_Incidencia & "',"
            mConsulta &= IIf(pTipo <> "", "'" & pTipo & "'", "null") & ","
            mConsulta &= IIf(pTipoFijo <> "", "'" & pTipoFijo & "'", "null") & ","
            mConsulta &= IIf(pMaximo <> "", pMaximo, "null") & ","
            mConsulta &= IIf(pFecha_Base <> "", "'" & pFecha_Base & "'", "null") & ","
            mConsulta &= IIf(pFecha_Termino <> "", "'" & pFecha_Termino & "'", "null") & ","
            mConsulta &= IIf(pTiempo_Maximo <> "", pTiempo_Maximo, "null") & ","
            mConsulta &= IIf(pOrden <> "", pOrden, "null") & ","
            mConsulta &= IIf(pGrupo <> "", pGrupo, "null") & ","
            mConsulta &= IIf(pSeleccionable <> "", "'" & pSeleccionable & "'", "null") & ","
            mConsulta &= IIf(pMaximo_Horas <> "", pMaximo_Horas, "null") & ","
            mConsulta &= IIf(pSeleccionable <> "", "'" & pSeleccionable & "'", "null") & ","
            mConsulta &= IIf(Maximo_Duracion <> 0, "" & Maximo_Duracion & "", "null") & ","
            mConsulta &= IIf(Minimo_Duracion <> 0, "" & Minimo_Duracion & "", "null") & ","
            mConsulta &= IIf(Tiempo_Minimo <> 0, "" & Tiempo_Minimo & "", "null") & ","
            mConsulta &= IIf(Tiempo_Minimo <> 0, "" & Naturales & "", "0") & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return mauxdbl
        Catch ex As Exception
            Trata_Error("Error en Inserta_Incidencia", ex, mConsulta)
        End Try


    End Function

    Public Function Inserta_Tipo_Contrato(ByVal pDesc_Tipo As String, Optional ByVal pObs_Tipo As String = "") As Integer Implements PresenciaDAO.Inserta_Tipo_Contrato
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_tipocontrato) FROM tipocontrato"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into tipocontrato (cod_tipocontrato, desc_tipocontrato, obs_tipocontrato) " _
                & " values (" & mauxdbl & ","
            mConsulta &= "'" & pDesc_Tipo & "',"
            mConsulta &= IIf(pObs_Tipo <> "", "'" & pObs_Tipo & "'", "null") & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return mauxdbl
        Catch ex As Exception
            Trata_Error("Error en Inserta_TipoContrato", ex, mConsulta)
        End Try


    End Function

    Public Function Lista_Tipo_Contrato_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Tipo As String = "", Optional ByVal pDescripcion As String = "") As Boolean Implements PresenciaDAO.Lista_Tipo_Contrato_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM tipocontrato"
            Dim mWhere As String
            If pCod_Tipo <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_tipocontrato = " & pCod_Tipo
            End If
            If pDescripcion <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " upper(desc_tipocontrato) like '" & UCase(pDescripcion) & "%'"
            End If
            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere
            mConsulta &= " order by cod_tipocontrato"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_TipoContrato_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Elimina_Tipo_Contrato(ByVal pCod_Tipo As String) As Boolean Implements PresenciaDAO.Elimina_Tipo_Contrato
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try


            Dim mCommand_valor As New OleDb.OleDbCommand
0:          Dim mReader As Object
            Dim mauxdbl As Integer



            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT count(*) FROM tipocontrato_incidencia where cod_incidencia=" & pCod_Tipo
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "SELECT count(*) FROM asocia_emp_tipocontrato where cod_tipocontrato=" & pCod_Tipo
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "delete from tipocontrato "
            mSQL &= "where cod_tipocontrato = " & pCod_Tipo

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Incidencia", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Actualiza_Tipo_Contrato(ByVal pCod_Tipo As String, Optional ByVal pDesc_Tipo As String = "", Optional ByVal pObs_Tipo As String = "") As Boolean Implements PresenciaDAO.Actualiza_Tipo_Contrato
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String

        Try

            mSQL = "update tipocontrato set "

            If pDesc_Tipo = "" Then
                mWhere &= " desc_tipocontrato =null"
            Else
                mWhere &= " desc_tipocontrato = '" & pDesc_Tipo & "'"
            End If

            If pObs_Tipo = "" Then
                If mWhere <> "" Then
                    mWhere &= ", obs_tipocontrato = null"
                Else
                    mWhere &= " obs_tipocontrato = null"
                End If
            Else
                If mWhere <> "" Then
                    mWhere &= ", obs_tipocontrato = '" & pObs_Tipo & "'"
                Else
                    mWhere &= " obs_tipocontrato = '" & pObs_Tipo & "'"
                End If
            End If

            mSQL &= mWhere
            mSQL &= " where cod_tipocontrato = " & pCod_Tipo
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Tipocontrato", ex, mSQL)
        End Try


    End Function

    Public Function Actualiza_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo_Antiguo As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta_Antigua As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean Implements PresenciaDAO.Actualiza_Asocia_Emp_Tipocontrato
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String
        Dim mFechas As String
        Try


            Dim mCommand_valor As New OleDb.OleDbCommand
            Dim mReader As Object

            If pFecha_Baja = "" Then pFecha_Baja = "31/12/2099"

            mFechas = "SELECT count(*) FROM Asocia_Emp_Tipocontrato where dni='" & pDNI & "'"
            mFechas &= " and not (dni = '" & pDNI & "' and cod_tipocontrato =" & pCod_Tipo_Antiguo & " and fecha_alta='" & pFecha_Alta_Antigua & "')"
            mFechas &= " and ((FECHA_ALTA >= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_BAJA >= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA <= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
              & "(FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA >= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA is null)) "

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mFechas
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()

            If pFecha_Baja = "31/12/2099" Then pFecha_Baja = ""

            mSQL = "update asocia_emp_tipocontrato set "

            mSQL &= " cod_tipocontrato =" & pCod_Tipo
            mSQL &= ", fecha_alta ='" & pFecha_Alta & "'"


            If pFecha_Baja = "" Then
                mWhere &= ", fecha_baja =null"
            Else
                mWhere &= ", fecha_baja = '" & pFecha_Baja & "'"
            End If

            mSQL &= mWhere
            mSQL &= " where dni = '" & pDNI & "' and cod_tipocontrato =" & pCod_Tipo_Antiguo & " and fecha_alta='" & pFecha_Alta_Antigua & "'"
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asocia_Emp_Tipocontrato", ex, mSQL)
        End Try


    End Function

    Public Function Elimina_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta As String) As Boolean Implements PresenciaDAO.Elimina_Asocia_Emp_Tipocontrato
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "delete from Asocia_Emp_Tipocontrato "
            mSQL &= " where dni = '" & pDNI & "' and cod_tipocontrato =" & pCod_Tipo & " and fecha_alta='" & pFecha_Alta & "'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asocia_Emp_Tipocontrato", ex, mSQL)
            Return False
        End Try

    End Function

    
    Public Function Inserta_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean Implements PresenciaDAO.Inserta_Asocia_Emp_Tipocontrato
        Dim mConsulta As String
        Dim mFechas As String

        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'antes comprobamos si existe otro asociacion la cual pise ésta:

            If pFecha_Baja = "" Then pFecha_Baja = "31/12/2099"
            mFechas = "SELECT count(*) FROM Asocia_Emp_Tipocontrato where dni='" & pDNI & "'"
            mFechas = mFechas & " and ((FECHA_ALTA >= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_BAJA >= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA <= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
                       & "(FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA >= TO_DATE('" & Format(CDate(pFecha_Baja), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_ALTA <= TO_DATE('" & Format(CDate(pFecha_Alta), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA is null))"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mFechas
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()

            If pFecha_Baja = "31/12/2099" Then pFecha_Baja = ""
            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into Asocia_Emp_Tipocontrato (dni, cod_tipocontrato, fecha_alta, fecha_baja) " _
                & " values ('" & pDNI & "'," & pCod_Tipo & ",'" & pFecha_Alta & "',"
            mConsulta &= IIf(pFecha_Baja <> "", "'" & pFecha_Baja & "'", "null") & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asocia_Emp_Tipocontrato", ex, mConsulta)
        End Try


    End Function

    Public Function Lista_Asocia_Emp_Tipocontrato_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pDNI As String) As Boolean Implements PresenciaDAO.Lista_Asocia_Emp_Tipocontrato_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM Asocia_Emp_Tipocontrato A, tipocontrato T "
            mConsulta &= " where  A.cod_tipocontrato= T.Cod_tipocontrato and A.dni ='" & pDNI & "'"
            mConsulta &= " order by A.fecha_alta"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Asocia_Emp_Tipocontrato_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Asocia_Emp_Tipocontrato_Fechas_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pDNI As String, ByVal pDesde As String, ByVal pHasta As String) As Boolean Implements PresenciaDAO.Lista_Asocia_Emp_Tipocontrato_Fechas_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM Asocia_Emp_Tipocontrato A, tipocontrato T "
            mConsulta &= " where  A.cod_tipocontrato= T.Cod_tipocontrato and A.dni ='" & pDNI & "'"
            mConsulta &= " and ((FECHA_ALTA >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_ALTA <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_BAJA >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
                & "(FECHA_ALTA <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA >= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_ALTA <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA is null))"
            mConsulta &= " order by A.fecha_alta"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Asocia_Emp_Tipocontrato_Fechas_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Asocia_Emp_Tipocontrato_Dia_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pDNI As String, ByVal pDia As String) As Boolean Implements PresenciaDAO.Lista_Asocia_Emp_Tipocontrato_Dia_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM Asocia_Emp_Tipocontrato A, tipocontrato T "
            mConsulta &= " where  A.cod_tipocontrato= T.Cod_tipocontrato and A.dni ='" & pDNI & "'"
            mConsulta &= " and ((FECHA_ALTA <= TO_DATE('" & Format(CDate(pDia), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA >= TO_DATE('" & Format(CDate(pDia), "dd/MM/yyyy") & "','DD/MM/YYYY')) " _
                & " or (FECHA_ALTA <= TO_DATE('" & Format(CDate(pDia), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_BAJA is null))"
            mConsulta &= " order by A.fecha_alta"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Asocia_Emp_Tipocontrato_Fechas_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Actualiza_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean Implements PresenciaDAO.Actualiza_TC_Incidencia
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String

        Try

            mSQL = "update tipocontrato_incidencia set "


            If pTipo = "" Then
                mSQL &= " tipo =null"
            Else
                mSQL &= " tipo ='" & pTipo & "'"
            End If

            If pTipoFijo = "" Then
                mSQL &= ", tipofijo = null"
            Else
                mSQL &= ", tipofijo = '" & pTipoFijo & "'"
            End If

            If pMaximo = "" Then
                mSQL &= ", maximo = null"
            Else
                mSQL &= ", maximo = " & pMaximo
            End If

            If pFecha_Base = "" Then
                mSQL &= ", fecha_base = null"
            Else
                mSQL &= ", fecha_base = '" & pFecha_Base & "'"
            End If

            If pFecha_Termino = "" Then
                mSQL &= ", fecha_termino = null"
            Else
                mSQL &= ", fecha_termino = '" & pFecha_Termino & "'"
            End If

            If pTiempo_Maximo = "" Then
                mSQL &= ", tiempo_maximo = null"
            Else
                mSQL &= ", tiempo_maximo = " & pTiempo_Maximo
            End If

            If pOrden = "" Then
                mSQL &= ", orden = null"
            Else
                mSQL &= ", orden = " & pOrden
            End If

            If pSeleccionable = "" Then
                mSQL &= ", seleccionable = null"
            Else
                mSQL &= ", seleccionable = '" & pSeleccionable & "'"
            End If

            If pMaximo_Horas = "" Then
                mSQL &= ", maximo_horas = null"
            Else
                mSQL &= ", maximo_horas = " & pMaximo_Horas
            End If

            If Maximo_Duracion <> 0 Then
                mSQL &= ", MAXIMO_DURACION = '" & Maximo_Duracion & "'"
            Else
                mSQL &= ", MAXIMO_DURACION = null"
            End If

            If Minimo_Duracion <> 0 Then
                mSQL &= ", MINIMO_DURACION = '" & Minimo_Duracion & "'"
            Else
                mSQL &= ", MINIMO_DURACION = null"
            End If

            If Tiempo_Minimo <> 0 Then
                mSQL &= ", TIEMPO_MINIMO = '" & Tiempo_Minimo & "'"
            Else
                mSQL &= ", TIEMPO_MINIMO = null"
            End If

            If Naturales = 1 Then
                mSQL &= ", NATURALES = 1"
            Else
                mSQL &= ", NATURALES = 0"
            End If



            mSQL &= " where cod_incidencia = " & pCod_Incidencia & " and cod_tipocontrato =" & pCod_TC
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_TC_Incidencia", ex, mSQL)
        End Try


    End Function

    Public Function Elimina_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String) As Boolean Implements PresenciaDAO.Elimina_TC_Incidencia
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "delete from tipocontrato_incidencia "
            mSQL &= "where cod_incidencia = " & pCod_Incidencia & " and cod_tipocontrato = " & pCod_TC

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_TCIncidencia", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean Implements PresenciaDAO.Inserta_TC_Incidencia
        Dim mConsulta As String
        Dim mCommand_valor As New OleDb.OleDbCommand

        Try

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into tipocontrato_incidencia (cod_tipocontrato, cod_incidencia, tipo, tipofijo, maximo, fecha_base, fecha_termino, tiempo_maximo, orden, seleccionable, maximo_horas, MAXIMO_DURACION, MINIMO_DURACION, TIEMPO_MINIMO,NATURALES) " _
                & " values (" & pCod_TC & "," & pCod_Incidencia & ","
            mConsulta &= IIf(pTipo <> "", "'" & pTipo & "'", "null") & ","
            mConsulta &= IIf(pTipoFijo <> "", "'" & pTipoFijo & "'", "null") & ","
            mConsulta &= IIf(pMaximo <> "", pMaximo, "null") & ","
            mConsulta &= IIf(pFecha_Base <> "", "'" & pFecha_Base & "'", "null") & ","
            mConsulta &= IIf(pFecha_Termino <> "", "'" & pFecha_Termino & "'", "null") & ","
            mConsulta &= IIf(pTiempo_Maximo <> "", pTiempo_Maximo, "null") & ","
            mConsulta &= IIf(pOrden <> "", pOrden, "null") & ","
            mConsulta &= IIf(pSeleccionable <> "", "'" & pSeleccionable & "'", "null") & ","
            mConsulta &= IIf(pMaximo_Horas <> "", pMaximo_Horas, "null") & ","
            mConsulta &= IIf(Maximo_Duracion <> 0, "" & Maximo_Duracion & "", "null") & ","
            mConsulta &= IIf(Minimo_Duracion <> 0, "" & Minimo_Duracion & "", "null") & ","
            mConsulta &= IIf(Tiempo_Minimo <> 0, "" & Tiempo_Minimo & "", "null") & ","
            mConsulta &= IIf(Naturales <> 0, "" & Naturales & "", "0") & ")"
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_TCIncidencia", ex, mConsulta)
            Return False
        End Try


    End Function

    Public Function Lista_TC_Incidencias_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_TC As Integer = -1, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pSeleccionable As String = "", Optional ByVal pCodigoIncidenciaDeCompensacion As Integer = -1) As Boolean Implements PresenciaDAO.Lista_TC_Incidencias_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT  "
            mConsulta &= " cod_tipocontrato,"
            mConsulta &= " cod_incidencia,"
            mConsulta &= " nvl(maximo, 0) maximo,"
            mConsulta &= " nvl(fecha_base, '01/01') fecha_base,"
            mConsulta &= " tipo, "
            mConsulta &= " nvl(tiempo_maximo, 0) tiempo_maximo,"
            mConsulta &= " nvl(fecha_termino, '31/12') fecha_termino,"
            mConsulta &= " nvl(orden, 0) orden,"
            mConsulta &= " tipofijo,"
            mConsulta &= " seleccionable, nvl(maximo_horas,0) maximo_horas,"
            mConsulta &= " COD_INCIDENCIA_COMPENSACION, CADUCIDAD_COMPENSACION,TIEMPO_MINIMO,MINIMO_DURACION,MAXIMO_DURACION,"
            mConsulta &= " PL_MAXIMO, PL_SOLICITAR, PL_MAXIMO_HORAS,NATURALES"

            mConsulta &= " FROM tipocontrato_incidencia "
            Dim mWhere As String
            If pCod_TC <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_tipocontrato = " & pCod_TC
            End If

            If pCod_Incidencia <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_incidencia = " & pCod_Incidencia
            End If

            If pCodigoIncidenciaDeCompensacion <> -1 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_incidencia_compensacion = " & pCodigoIncidenciaDeCompensacion
            End If

            If pSeleccionable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " seleccionable = '" & pSeleccionable & "'"
            End If

            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_TC_Incidencias_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_TC_Incidencias_Excepciones_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_TC As Integer = -1, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean Implements PresenciaDAO.Lista_TC_Incidencias_Excepciones_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM tipocontrato T, tipocontrato_incidencia TC where T.cod_tipocontrato= TC.cod_tipocontrato(+) "
            Dim mWhere As String
            'If pCod_TC <> -1 Then
            '    mWhere = mWhere & " AND TC.cod_tipocontrato = " & pCod_TC
            'End If

            If pCod_Incidencia <> -1 Then
                mWhere = mWhere & " AND TC.cod_incidencia(+) = " & pCod_Incidencia
            End If


            If mWhere <> "" Then mConsulta = mConsulta & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_TC_Incidencias_Excepciones_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Inserta_DiasCalendario(Optional ByVal pCodigo As String = "", Optional ByVal pFecha As String = "", Optional ByVal pDesc_Fecha As String = "", Optional ByVal pCod_Horario As String = "") As Boolean Implements PresenciaDAO.Inserta_DiasCalendario
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand


        Try

            mSQL = "insert into diascalendario (cod_calendario, fecha_calendario,desc_fecha, cod_horario) " _
            & " values (" & pCodigo & ",'" & pFecha & "','" & pDesc_Fecha & "'," & pCod_Horario & ")"
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_DiasCalendario", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Lista_Horarios_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Horario As String = "") As Boolean Implements PresenciaDAO.Lista_Horarios_Dataset
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * FROM horarios"
            Dim mWhere As String
            If pCod_Horario <> "" Then
                mWhere = mWhere & " cod_horario = " & pCod_Horario
            End If

            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Horarios_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Cambio_Calendario_Festivo(ByVal pFecha As String, ByVal pCod_Cal As String) As Boolean Implements PresenciaDAO.Cambio_Calendario_Festivo
        'tenemos que llamas a actualiza saldo por cada tio que tenga asignado rda calendario festivo en esa fecha.

        'vamos a ver: hemos modificado un dia festivo en un caledario.
        'PRIMERO: hay que buscar los grupos de trabajo que tengan asociado ese dia para ese calendario
        'SEGUNDO: hay que buscar los dni de los empleados que tengan asociado ese grupo de trabajo en esa fecha
        'tercero: hay que llamar a la funcion actualiza_saldo por cada dni y esa fecha que nos salga
        Dim mDatos As DataSet
        Dim mConsulta As String
        Try

            mConsulta = "SELECT dni_empl FROM asociausuariogrupotrabajo where "
            mConsulta &= " ((TO_DATE('" & pFecha & "', 'DD/MM/YYYY') >= FECHA_DESDE and TO_DATE('" & pFecha & "', 'DD/MM/YYYY') <= FECHA_HASTA) or (TO_DATE('" & pFecha & "', 'DD/MM/YYYY') >= FECHA_DESDE and FECHA_HASTA is null))"
            mConsulta &= " and cod_grupotrabajo in ("
            mConsulta &= "SELECT cod_grupotrabajo FROM asociagrupotrabajocalendario where cod_Calendario = " & pCod_Cal
            mConsulta &= " and ((TO_DATE('" & pFecha & "', 'DD/MM/YYYY') >= FECHA_DESDE and TO_DATE('" & pFecha & "', 'DD/MM/YYYY') <= FECHA_HASTA) or (TO_DATE('" & pFecha & "', 'DD/MM/YYYY') >= FECHA_DESDE and FECHA_HASTA is null))"
            mConsulta &= ")"

            mDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            Dim i As Integer
            mDataAdapter.Fill(mDatos)
            If mDatos.Tables(0).Rows.Count > 0 Then
                'actualizamos el saldo de todos los dni que salgan
                For i = 0 To mDatos.Tables(0).Rows.Count - 1
                    Me.Actualiza_Saldo(mDatos.Tables(0).Rows(i)("dni_empl"), pFecha)
                Next
            End If
            mDatos = Nothing
            Return True
        Catch ex As Exception
            Trata_Error("Error en Cambio_Calendario_Festivo", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Cambio_Calendario_Laborable(ByVal pCod_Cal As String, ByVal pDia As Integer) As Boolean Implements PresenciaDAO.Cambio_Calendario_Laborable

        Dim mDatos_Fechas As DataSet
        Dim mConsulta As String
        Dim mConsulta_Fechas As String
        Dim Minimo_1 As String
        Dim Minimo_2 As String
        Dim Maximo_1 As String
        Dim Maximo_2 As String
        Dim mGrupotrabajo As String
        Dim i As Integer
        Dim j As Integer
        Dim total As Integer = 0
        Try

            mConsulta_Fechas = "select min (fecha_desde) min, max(nvl(fecha_hasta, to_date(to_char(sysdate + 365, 'DD/MM/YYYY'), 'DD/MM/YYYY'))) max, cod_grupotrabajo  from asociagrupotrabajocalendario "
            mConsulta_Fechas &= " where cod_calendario = " & pCod_Cal & " "
            mConsulta_Fechas &= " and ((TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') >= FECHA_DESDE and TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') <= FECHA_HASTA) or (TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') >= FECHA_DESDE and FECHA_HASTA is null)) group by cod_grupotrabajo"
            mDatos_Fechas = New DataSet
            Dim mDataAdapter2 As New OleDb.OleDbDataAdapter(mConsulta_Fechas, mConexion)
            mDataAdapter2.Fill(mDatos_Fechas)

            If mDatos_Fechas.Tables(0).Rows.Count > 0 Then
                'ya tenemos por cada grupo, la fecha maxima hasta la que debemos llegar
                'ahora buscamos la fecha maxima de cada grupo en los que esten asociados a empleados
                'la menor de las dos, sera la fecha hasta donde hay que actualizar el diario, y el comienzo
                'el dia de hoy. Si ese grupo tiene fecha_hasta nula, solo actualizamos el diario hasta el año que viene
                For i = 0 To mDatos_Fechas.Tables(0).Rows.Count - 1
                    Minimo_1 = mDatos_Fechas.Tables(0).Rows(i)("min")
                    Maximo_1 = mDatos_Fechas.Tables(0).Rows(i)("max")
                    mGrupotrabajo = mDatos_Fechas.Tables(0).Rows(i)("cod_grupotrabajo")

                    Dim mDatos2 As DataSet
                    mConsulta_Fechas = "select min(fecha_desde) min, max(nvl(fecha_hasta, to_date(to_char(sysdate + 365, 'DD/MM/YYYY'), 'DD/MM/YYYY'))) max, dni_empl from asociausuariogrupotrabajo "
                    mConsulta_Fechas &= " where cod_grupotrabajo = " & mGrupotrabajo
                    mConsulta_Fechas &= " and ((TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') >= FECHA_DESDE and TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') <= FECHA_HASTA) or (TO_DATE(TO_CHAR(sysdate, 'DD/MM/YYYY'), 'DD/MM/YYYY') >= FECHA_DESDE and FECHA_HASTA is null)) group by dni_empl"
                    mDatos2 = New DataSet
                    Dim mDataAdapter3 As New OleDb.OleDbDataAdapter(mConsulta_Fechas, mConexion)
                    mDataAdapter3.Fill(mDatos2)

                    If mDatos2.Tables(0).Rows.Count > 0 Then
                        For j = 0 To mDatos2.Tables(0).Rows.Count - 1
                            Minimo_2 = mDatos2.Tables(0).Rows(j)("min")
                            Maximo_2 = mDatos2.Tables(0).Rows(j)("max")
                            '**************************************************************************
                            'ya tenemos las fechas para cada grupo de trabajo y dni
                            'tenemos que actualizar desde el dia de hoy hasta esa fecha minima
                            'para cada tio, Para ello, miramos si esa fecha y ese tio pertenecen a la
                            'tabla diario
                            '**************************************************************************

                            Dim mDESDE As String = Devuelve_Maxima_Fecha(Minimo_1, Minimo_2)
                            Dim mHASTA As String = Devuelve_Minima_Fecha(Maximo_1, Maximo_2)
                            Dim mDNI As String = mDatos2.Tables(0).Rows(j)("dni_empl")

                            If CDate(mDESDE) < Now Then
                                'antes no sha salido una fecha mínima, pero si esa fecha es menos que la de hoy, tenemos que
                                'cambiar los diarios a partir de hoy.
                                mDESDE = Format(Now, "dd/MM/yyyy")
                            End If

                            If CDate(mHASTA) > CDate(mDESDE) Then
                                'tenemos que ver si la fecha de hoy se corresponde al dia de la
                                'semana correcto, si no, avanzamos al siguiente dia que hayamos cambiado

                                'If Me.Lee_Hoy(mDESDE) <> pDia Then
                                'mDESDE = Me.Lee_Siguiente_Semana(mDESDE, pDia)
                                'End If
                                mDESDE = Cuadra_Dia(mDESDE, pDia)
                                While CDate(mDESDE) <= CDate(mHASTA)
                                    'actualizamos para cada tio esa fecha
                                    'If Me.Lee_De_Diario(mDESDE, mDNI) Then
                                    '
                                    Me.Actualiza_Saldo(mDNI, CDate(mDESDE))
                                    total += 1
                                    '
                                    'End If
                                    mDESDE = Me.Lee_Siguiente_Semana(mDESDE, pDia)
                                End While

                            End If

                            '**************************************************************************
                            '**************************************************************************
                            '**************************************************************************
                        Next
                    End If
                    mDatos2 = Nothing
                Next
            End If
            mDatos_Fechas = Nothing
            '***************************************************************************************
            'esta variable es para ver cuantas veces actualiza, para verlo en tiempo de depuracion
            'aparte de eso, no sirve para nada
            total = 0
            Return True
        Catch ex As Exception
            Trata_Error("Error en Cambio_Calendario_Laborable", ex, mConsulta)
            Return False
        End Try

    End Function

    Private Function Devuelve_Minima_Fecha(ByVal Minimo1 As String, ByVal Minimo2 As String) As String
        If Minimo1 = "" And Minimo2 = "" Then
            ' si los dos estan vacion, devolvemos como fecha el mismo dia del año que viene
            Return Day(Now) & "/" & Month(Now) & "/" & (Year(Now) + 1)
        ElseIf Minimo1 = "" And Minimo2 <> "" Then
            Return Minimo2
        ElseIf Minimo1 <> "" And Minimo2 = "" Then
            Return Minimo1
        Else
            'ninguno de los dos está vacio
            If CDate(Minimo1) > CDate(Minimo2) Then
                Return Minimo2
            Else
                Return Minimo1
            End If
        End If
    End Function

    Private Function Devuelve_Maxima_Fecha(ByVal Maximo1 As String, ByVal Maximo2 As String) As String

        If CDate(Maximo1) > CDate(Maximo2) Then
            Return Maximo1
        Else
            Return Maximo2
        End If

    End Function

    Private Function Lee_De_Diario(ByVal pFecha As String, ByVal pDNI As String) As Boolean
        Dim mConsulta As String
        Dim pDatos As DataSet
        Dim mCuenta As Integer
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'mConsulta = "SELECT count(*) FROM diario where dni='" & pDNI & "' and fecha='" & pFecha & "'"

            'Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'pDatos = New DataSet
            ''Conectarse, buscar datos y desconectarse de la base de datos 
            'mDataAdapter.Fill(pDatos)
            'pDatos.Clear()
            'mCuenta = pDatos.Tables(0).Rows.Count
            'pDatos = Nothing
            'If mCuenta > 0 Then
            '    Return True
            'Else
            '    Return False
            'End If
            mConsulta = "SELECT count(*) FROM diario where dni='" & pDNI & "' and fecha='" & pFecha & "'"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                If mReader(0) > 0 Then
                    mReader.Close()
                    Return True
                Else
                    mReader.Close()
                    Return False
                End If
            End If
        Catch ex As Exception
            Trata_Error("Error en Lee_De_Diario", ex, mConsulta)
            Return False
        End Try
    End Function

    Private Function Cuadra_Dia(ByVal pFecha As String, ByVal Dia As Integer) As String
        'tenemos que devolver la primera fecha corresponsiente al dia 
        ' donde dia=1 es lunes, 2, martes ....
        'a partir de la fecha
        Dim cual_hoy As Integer = Lee_Hoy(pFecha)
        If Dia = cual_hoy Then
            Return pFecha
        ElseIf Dia > cual_hoy Then
            Return Format(DateAdd(DateInterval.Day, Dia - cual_hoy, CDate(pFecha)), "dd/MM/yyyy")
        Else
            Return Format(DateAdd(DateInterval.Day, cual_hoy - Dia, CDate(pFecha)), "dd/MM/yyyy")
        End If

    End Function
    Private Function Lee_Siguiente_Semana(ByVal pFecha As String, ByVal Dia As Integer) As String
        Return Format(DateAdd(DateInterval.WeekOfYear, 1, CDate(pFecha)), "dd/MM/yyyy")
    End Function

    Private Function Lee_Hoy(ByVal pFecha As String) As Integer
        Return Weekday(pFecha, FirstDayOfWeek.Monday)
    End Function

    Public Function Elimina_del_Diario(ByVal pTipo As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pDNI As String = "", Optional ByVal pGrupo As String = "") As Boolean Implements PresenciaDAO.Elimina_del_Diario
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim Lista_Dni As String
        Dim Vectores_DNI() As String
        Dim mCadena As String
        Dim mDNI As New DataSet
        Dim i As Integer
        Dim mFecha As String = pFecha_Desde

        Try

            '**********************************************************************
            'primero, pillamos la lista de dni para despues recalcularles el saldo
            '**********************************************************************
            Dim mDataAdapter As OleDb.OleDbDataAdapter

            Select Case pTipo
                Case "TODOS"
                    'mCadena = "select distinct(dni) from diario where fecha >= '" & pFecha_Desde & "' and fecha <= '" & pFecha_Hasta & "'"
                    mCadena = "Select dni from empleados where calcula_saldo ='S'"
                    mDataAdapter = New OleDb.OleDbDataAdapter(mCadena, mConexion)
                    mDataAdapter.Fill(mDNI)

                    If mDNI.Tables(0).Rows.Count > 0 Then
                        For i = 0 To mDNI.Tables(0).Rows.Count - 1
                            Lista_Dni &= mDNI.Tables(0).Rows(i)(0)
                            Lista_Dni &= ","
                        Next
                    End If
                    If Lista_Dni <> "" Then Lista_Dni = Left(Lista_Dni, Len(Lista_Dni) - 1)
                Case "DNI"
                    Lista_Dni = pDNI
                Case "GRUPO"
                    'mCadena = "select distinct(dni) FROM diario WHERE diario.dni IN " _
                    mCadena = "Select dni from empleados where calcula_saldo ='S' and dni IN" _
                        & " (SELECT  dni_empl FROM asociausuariogrupotrabajo WHERE cod_grupotrabajo = " & pGrupo _
                        & " AND fecha_desde <= '" & pFecha_Desde & "' " _
                        & " AND (fecha_hasta >= '" & pFecha_Desde & "' OR fecha_hasta IS NULL)) " '_
                    '& " AND fecha = '" & pFecha_Desde & "'"

                    mDataAdapter = New OleDb.OleDbDataAdapter(mCadena, mConexion)
                    mDataAdapter.Fill(mDNI)

                    If mDNI.Tables(0).Rows.Count > 0 Then
                        For i = 0 To mDNI.Tables(0).Rows.Count - 1
                            Lista_Dni &= mDNI.Tables(0).Rows(i)(0)
                            Lista_Dni &= ","
                        Next
                    End If
                    If Lista_Dni <> "" Then Lista_Dni = Left(Lista_Dni, Len(Lista_Dni) - 1)
            End Select



            '****************************************************************
            'segundo, borramos del diario
            '****************************************************************

            'Select Case pTipo
            '    Case "TODOS"
            'mSQL = "delete from diario where fecha >= '" & pFecha_Desde & "' and fecha <= '" & pFecha_Hasta & "'"
            '    Case "DNI"
            'mSQL = "delete from diario where dni ='" & pDNI & "' and fecha >= '" & pFecha_Desde & "' and fecha <= '" & pFecha_Hasta & "'"
            '    Case "GRUPO"
            'mSQL = "DELETE FROM diario WHERE diario.dni IN "
            'mSQL &= " (SELECT  dni_empl FROM asociausuariogrupotrabajo WHERE cod_grupotrabajo = " & pGrupo
            'mSQL &= " AND fecha_desde <= '" & pFecha_Desde & "' "
            'mSQL &= " AND (fecha_hasta >= '" & pFecha_Desde & "' OR fecha_hasta IS NULL)) "
            'mSQL &= " AND fecha = '" & pFecha_Desde & "'"
            'End Select
            'mCommand.Connection = mConexion
            'mCommand.CommandText = mSQL
            'mCommand.ExecuteNonQuery()


            '****************************************************************
            'tercero, recalculamos todos los saldos
            '****************************************************************

            Select Case pTipo
                Case "TODOS"
                    If Lista_Dni <> "" Then
                        Vectores_DNI = Split(Lista_Dni, ",")
                        For i = 0 To Vectores_DNI.Length - 1
                            mFecha = pFecha_Desde
                            While CDate(mFecha) <= CDate(pFecha_Hasta)
                                Actualiza_Saldo(Vectores_DNI(i), CDate(mFecha))
                                mFecha = Format(DateAdd(DateInterval.Day, 1, CDate(mFecha)), "dd/MM/yyyy")
                            End While
                        Next
                    End If

                Case "GRUPO"
                    If Lista_Dni <> "" Then
                        Vectores_DNI = Split(Lista_Dni, ",")
                        For i = 0 To Vectores_DNI.Length - 1
                            mFecha = pFecha_Desde
                            While CDate(mFecha) <= CDate(pFecha_Hasta)
                                Actualiza_Saldo(Vectores_DNI(i), CDate(mFecha))
                                mFecha = Format(DateAdd(DateInterval.Day, 1, CDate(mFecha)), "dd/MM/yyyy")
                            End While
                        Next
                    End If
                Case "DNI"
                    mFecha = pFecha_Desde
                    While CDate(mFecha) <= CDate(pFecha_Hasta)
                        Actualiza_Saldo(Lista_Dni, CDate(mFecha))
                        mFecha = Format(DateAdd(DateInterval.Day, 1, CDate(mFecha)), "dd/MM/yyyy")
                    End While
                Case "GRUPO"
            End Select


            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_del_Diario", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Lista_Asignaciones_Responsables(ByRef pDatos As System.Data.DataSet, Optional ByVal pResponsable As String = "") As Boolean Implements PresenciaDAO.Lista_Asignaciones_Responsables
        Dim mConsulta As String
        Try
            If pResponsable = "" Then
                'buscamos los responsables que no tienen otro responsable a su vez.
                mConsulta = "SELECT distinct (A.ID_RESPONSABLE), A.siguiente, E.ape1, E.ape2, E.nombre FROM ASIGNACION_RESPONSABLE A, empleados E"
                mConsulta &= " where E.dni = A.ID_responsable and A.id_responsable not in (select id_usuario from asignacion_responsable) "
                mConsulta &= " and id_usuario not like 'G%' "
                mConsulta &= " order by ape1, ape2, nombre"
            Else
                'busca los empleados de un responsable y que ademas son responsables de otros.
                mConsulta = "SELECT A.ID_USUARIO, (select distinct(siguiente) from asignacion_responsable where id_Responsable = a.id_usuario and siguiente <> '1') SIGUIENTE, E.ape1, E.ape2, E.nombre FROM ASIGNACION_RESPONSABLE A, empleados E"
                mConsulta &= " where E.dni = A.ID_USUARIO and A.ID_RESPONSABLE ='" & pResponsable & "' and A.id_usuario IN (select id_responsable from asignacion_responsable where id_usuario not like 'G%')"
                mConsulta &= " order by E.ape1, E.ape2, E.nombre"
            End If

            'Dim mWhere As String
            'If pCod_Horario <> "" Then
            'mWhere = mWhere & " cod_horario = " & pCod_Horario
            'End If

            'If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Asignaciones_Responsables", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Lista_Usuarios_Responsable(ByRef pDatos As System.Data.DataSet, Optional ByVal pResponsable As String = "", Optional ByVal pSinAsignar As Boolean = False) As Boolean Implements PresenciaDAO.Lista_Usuarios_Responsable
        Dim mConsulta As String
        Try

            If pSinAsignar = True Then
                mConsulta = "select dni, nombre, ape1, ape2 from empleados where dni not in "
                mConsulta &= " (select distinct(id_usuario) from ASIGNACION_RESPONSABLE) order by ape1, ape2, nombre"
            Else
                If pResponsable <> "" Then
                    mConsulta = "SELECT A.ID_USUARIO dni, E.ape1 ape1, E.ape2 ape2, E.nombre nombre FROM ASIGNACION_RESPONSABLE A, empleados E"
                    mConsulta &= " where E.dni = A.ID_USUARIO "
                    mConsulta &= " and A.id_responsable='" & pResponsable & "' order by E.ape1, E.ape2, E.nombre"
                Else
                    mConsulta = "select dni, nombre, ape1, ape2 from empleados order by ape1, ape2, nombre"
                End If
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Usuarios_Responsable", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Empleados_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Lista_Empleados_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, clave_emp, nombre, ape1, ape2, permite_visita, calcula_saldo, telefono, email, clave_web, centro, cargo, admin from empleados "

            If pDNI <> "" Then
                mWhere = " where upper(dni) like '" & UCase(pDNI) & "%'"
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            'If pClave_Empleado <> "" Then
            '    If mWhere <> "" Then
            '        mWhere = mWhere & " AND "
            '    Else
            '        mWhere = " WHERE "
            '    End If
            '    'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
            '    mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            'End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If
            If mWhere <> "" Then mConsulta &= mWhere
            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Empleados_DNI_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Lista_Empleados_DNI_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre || ' ' || ape1 || ' ' || ape2 NOMBRE from empleados "

            If pDNI <> "" Then
                mWhere = " where upper(dni) like '" & UCase(pDNI) & "%'"
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            'If pClave_Empleado <> "" Then
            '    If mWhere <> "" Then
            '        mWhere = mWhere & " AND "
            '    Else
            '        mWhere = " WHERE "
            '    End If
            '    'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
            '    mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            'End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If
            If mWhere <> "" Then mConsulta &= mWhere
            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Empleado(ByRef pDatos As System.Data.DataSet, ByVal pDNI As String) As Boolean Implements PresenciaDAO.Lista_Empleado
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2, clave_web, calcula_saldo from empleados "

            If pDNI <> "" Then
                mWhere = " where upper(dni) =  '" & UCase(pDNI) & "'"
            End If

            If mWhere <> "" Then mConsulta &= mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleado", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Inserta_Asignacion_Responsable_Comprobaciones(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Inserta_Asignacion_Responsable_Comprobaciones
        Dim mConsulta As String
        Dim mComp As String

        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim Existe As Boolean = False
        Dim mTiene_Siguiente As String
        Dim mDatos As New DataSet

        Try
            'antes comprobamos si existe otro asociacion la cual pise ésta:


            mComp = "SELECT count(*) FROM Asignacion_Responsable "
            mComp &= " where id_usuario ='" & pID_Usuario & "'"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mComp
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Existe = True
                End If
            Else
                Return False
            End If
            mReader.Close()

            'bucamos si el responsable existe como siguiente = 1, para dejarlo como estaba:
            Dim mCadena As String
            mTiene_Siguiente = "0"

            mConsulta = "SELECT siguiente from asignacion_responsable "
            mConsulta &= " where id_responsable = '" & pID_Responsable & "' and siguiente = '1' and id_usuario not like 'G%'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            If mDatos.Tables(0).Rows.Count > 0 Then
                mCadena = " '1' "
            Else
                mCadena &= " null "
            End If
            mDatos.Clear()
            mDatos = Nothing


            If Existe Then
                'si existe el empleado con otro responsable, actualizamos el campo responsable
                mConsulta = "update Asignacion_Responsable set id_responsable='" & pID_Responsable & "' "
                mConsulta &= ", siguiente = " & mCadena
                mConsulta &= " where id_usuario ='" & pID_Usuario & "'"
            Else
                'si no existe, lo añadimos
                mConsulta = "insert into Asignacion_Responsable (id_usuario, id_responsable, siguiente) values "
                mConsulta &= "('" & pID_Usuario & "','" & pID_Responsable & "'," & mCadena & ")"
            End If

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asignacion_Responsable_Comprobaciones", ex, mConsulta)
        End Try


    End Function

    Public Function Lista_Empleados_Asignacion_Responsable(ByRef pDatos As DataSet, ByVal Lista_DNI As String) As Boolean Implements PresenciaDAO.Lista_Empleados_Asignacion_Responsable
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2 from empleados "
            mConsulta &= " where dni in (" & UCase(Lista_DNI) & ")"
            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Responsables(ByRef pDatos As System.Data.DataSet, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Lista_Responsables
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select E.dni dni, E.nombre nombre, E.ape1 ape1, E.ape2 ape2 from empleados E, delegados D"
            mConsulta &= " where E.dni = D.id_delegado and D.id_responsable = '" & pID_Responsable & "'"
            mConsulta &= " order by E.ape1, E.ape2, E.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Responsables", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Responsables_Libres(ByRef pDatos As System.Data.DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "") As Boolean Implements PresenciaDAO.Lista_Responsables_Libres
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2 from empleados where dni not in"
            mConsulta &= " (select id_delegado from delegados where id_responsable = '" & pID_Responsable & "')"
            mConsulta &= " and dni <> '" & pID_Responsable & "'"
            If pDNI <> "" Then
                mConsulta &= " and DNI LIKE '" & pDNI & "%'"
            End If
            If pNombre <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & UCase(pNombre) & "%'))"
            End If
            If pApe1 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe1) & "%'))"
            End If
            If pApe2 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe2) & "%'))"
            End If


            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Responsables_Libres", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Responsables_Libres_Justificadores(ByRef pDatos As System.Data.DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pGrupo As String = "") As Boolean Implements PresenciaDAO.Lista_Responsables_Libres_Justificadores
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2 from empleados where dni not in"
            mConsulta &= " (select id_delegado from delegados where id_responsable = '" & pID_Responsable & "')"
            mConsulta &= " and dni <> '" & pID_Responsable & "'"
            If pDNI <> "" Then
                mConsulta &= " and DNI LIKE '" & pDNI & "%'"
            End If
            If pNombre <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & UCase(pNombre) & "%'))"
            End If
            If pApe1 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe1) & "%'))"
            End If
            If pApe2 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe2) & "%'))"
            End If


            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Responsables_Libres_Justificadores", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Responsables_Libres_Aprobadores(ByRef pDatos As System.Data.DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pGrupo As String = "") As Boolean Implements PresenciaDAO.Lista_Responsables_Libres_Aprobadores
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2 from empleados where dni not in"
            mConsulta &= " (select id_delegado from delegados where id_responsable = '" & pID_Responsable & "')"
            mConsulta &= " and dni <> '" & pID_Responsable & "'"
            mConsulta &= " and (dni in (select id_usuario from asignacion_responsable where id_responsable='" & pID_Responsable & "')"
            mConsulta &= " or dni in (select dni_empl from pertenecena where cod_grupo in (select substr(id_usuario,6,2) from asignacion_responsable where id_responsable='" & pID_Responsable & "' and id_usuario like 'GRP%')))"
            If pDNI <> "" Then
                mConsulta &= " and DNI LIKE '" & pDNI & "%'"
            End If
            If pNombre <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & UCase(pNombre) & "%'))"
            End If
            If pApe1 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe1) & "%'))"
            End If
            If pApe2 <> "" Then
                mConsulta &= " and SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & UCase(pApe2) & "%'))"
            End If


            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Responsables_Libres_Aprobadores", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Delegados_Rodas(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "") As Boolean Implements PresenciaDAO.Elimina_Delegados_Rodas
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE Delegados "
            If pID_Responsable <> "" Then
                mWhere = " WHERE ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            If pID_Delegado <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE"
                Else
                    mWhere &= " AND"
                End If
                mWhere &= " ID_Delegado = '" & pID_Delegado & "'"
            End If

            mConsulta &= mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
        Catch ex As Exception
            Trata_Error("Error en Elimina_Delegados", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_AutorizadoJustificar_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pGrupo As String = "") As Boolean Implements PresenciaDAO.Lista_AutorizadoJustificar_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select E.dni, E.nombre, E.ape1, E.ape2, A.fecha_desde, A.fecha_hasta from empleados E, autorizadojustificar A where "
            mConsulta &= " E.dni = A.dni"
            If pGrupo <> "" Then
                mConsulta &= " and A.grupos = '" & pGrupo & "'"
            End If

            mConsulta &= " order by E.ape1, E.ape2, E.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Me.Haz_Log("Ejecutando SQL:" & mConsulta, 3)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_AutorizadoConsultar_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pGrupo As String = "") As Boolean Implements PresenciaDAO.Lista_AutorizadoConsultar_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select E.dni, E.nombre, E.ape1, E.ape2, A.fecha_desde, A.fecha_hasta from empleados E, autorizadoconsultar A where "
            mConsulta &= " E.dni = A.dni"
            If pGrupo <> "" Then
                mConsulta &= " and A.grupos = '" & pGrupo & "'"
            End If

            mConsulta &= " order by E.ape1, E.ape2, E.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Me.Haz_Log("Ejecutando SQL:" & mConsulta, 3)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Inserta_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, Optional ByVal pFecha_Hasta As String = "") As Boolean Implements PresenciaDAO.Inserta_AutorizadoJustificar
        Dim Existe As Boolean
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Try

            mConsulta = "insert into AutorizadoJustificar (dni, grupos, fecha_desde, fecha_hasta) values  "
            If pGrupo = "Todos" Then
                mConsulta &= "('" & pDni & "','Todos','" & pFecha_Desde & "',"
            Else
                mConsulta &= "('" & pDni & "','" & Format(CInt(pGrupo), "0000") & "','" & pFecha_Desde & "',"
            End If
            If pFecha_Hasta <> "" Then
                mConsulta &= "'" & pFecha_Hasta & "')"
            Else
                mConsulta &= "null)"
            End If

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Elimina_AutorizadoJustificar(Optional ByVal pDni As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pFecha_Desde As String = "") As Boolean Implements PresenciaDAO.Elimina_AutorizadoJustificar
        Dim mConsulta As String
        Dim mWhere As String
        Try
            mConsulta = "DELETE AutorizadoJustificar "
            If pDni <> "" Then
                mWhere = " WHERE dni = '" & pDni & "'"
            End If
            If pGrupo <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE"
                Else
                    mWhere &= " AND"
                End If
                If pGrupo = "Todos" Then
                    mWhere &= " grupos = 'Todos'"
                Else
                    mWhere &= " grupos = '" & Format(CInt(pGrupo), "0000") & "'"
                End If

            End If
            If pFecha_Desde <> "" Then
                If mWhere = "" Then
                    mWhere = " WHERE"
                Else
                    mWhere &= " AND"
                End If
                mWhere &= " to_Date(to_char(fecha_Desde, 'DD/MM/YYYY'), 'DD/MM/YYYY') = '" & pFecha_Desde & "'"
            End If

            mConsulta &= mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_AutorizadoJustificar", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, ByVal pFecha_Desde_Nueva As String, Optional ByVal pFecha_Hasta As String = "") As Boolean Implements PresenciaDAO.Actualiza_AutorizadoJustificar
        Dim Existe As Boolean
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Try
            mConsulta = "update AutorizadoJustificar set "
            mConsulta &= " fecha_Desde = '" & pFecha_Desde_Nueva & "'"
            If pFecha_Hasta <> "" Then
                mConsulta &= ", fecha_hasta = '" & pFecha_Hasta & "'"
            Else
                mConsulta &= ", fecha_hasta = null"
            End If
            If pGrupo = "Todos" Then
                mConsulta &= " where dni ='" & pDni & "' and grupos='Todos' and fecha_desde ='" & pFecha_Desde & "'"
            Else
                mConsulta &= " where dni ='" & pDni & "' and grupos='" & Format(CInt(pGrupo), "0000") & "' and fecha_desde ='" & pFecha_Desde & "'"
            End If


            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Existe_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, Optional ByVal pFecha_Hasta As String = "") As Boolean Implements PresenciaDAO.Existe_AutorizadoJustificar
        Dim mConsulta As String
        Dim mComp As String

        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim Existe As Boolean = False
        Try
            'antes comprobamos si existe otro asociacion la cual pise ésta:


            mComp = "SELECT count(*) FROM AutorizadoJustificar "
            mComp &= " where dni ='" & pDni & "' and grupos='" & Format(CInt(pGrupo), "0000") & "' and fecha_desde ='" & pFecha_Desde & "'"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mComp
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Existe = True
                End If
            End If
            mReader.Close()
            Return Existe
        Catch ex As Exception
            Return False
        End Try
    End Function




    Public Function Resumen_Justificaciones_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pDNI As String) As Boolean Implements PresenciaDAO.Resumen_Justificaciones_Dataset
        'select codigo incidencia, sum(numero) numero  , count(dia) dias  from (
        'select count(*) numero, j.cod_incidencia codigo, fecha_justificada dia  
        'from justificaciones J, incidencias I 
        'where I.cod_incidencia = J.cod_incidencia and  dni_empl = '31690444D' 
        'and fecha_justificada >= '12/07/2002' and fecha_justificada <= '30/07/2002' 
        'group by  J.cod_incidencia, fecha_justificada
        'order by j.cod_incidencia
        ')
        'group by codigo

        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select codigo incidencia, sum(numero) numero  , count(dia) dias  from ( "
            mConsulta &= "select count(*) numero, j.cod_incidencia codigo, fecha_justificada dia  "
            mConsulta &= "from justificaciones_t J, incidencias I "
            mConsulta &= "where I.cod_incidencia = J.cod_incidencia and  dni_empl = '" & pDNI & "' "
            mConsulta &= "and fecha_justificada >= TO_DATE ('" & pFechaDesde & "', 'DD/MM/YYYY') "
            mConsulta &= "and fecha_justificada <= TO_DATE ('" & pFechaHasta & "', 'DD/MM/YYYY') "
            mConsulta &= "group by  J.cod_incidencia, fecha_justificada "
            mConsulta &= "order by j.cod_incidencia "
            mConsulta &= ") group by codigo"


            mConsulta &= " union "
            mConsulta &= " select codigo incidencia, 0  , 0 " 'count(dia)  "
            mConsulta &= " from ( "
            mConsulta &= " select count(*) numero, cod_incidencia codigo, fecha dia  "
            mConsulta &= " from solicitud s where  dni = '" & pDNI & "' "
            mConsulta &= " and fecha >= TO_DATE ('" & pFechaDesde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha <= TO_DATE ('" & pFechaHasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and estado in ('A', 'P', 'E') "
            mConsulta &= " and cod_incidencia not in "
            mConsulta &= " (select codigo incidencia  "
            mConsulta &= " from ( "
            mConsulta &= " select count(*) numero, j.cod_incidencia codigo, fecha_justificada dia  "
            mConsulta &= " from justificaciones_t J, incidencias I where I.cod_incidencia = J.cod_incidencia "
            mConsulta &= " and  dni_empl = '" & pDNI & "' "
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFechaDesde & "', 'DD/MM/YYYY') "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFechaHasta & "', 'DD/MM/YYYY') "
            mConsulta &= " group by  J.cod_incidencia, fecha_justificada "
            mConsulta &= " )"
            mConsulta &= " group by codigo)"
            mConsulta &= " group by  cod_incidencia, fecha "
            mConsulta &= " order by cod_incidencia ) "
            mConsulta &= " group by codigo"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Lista_Solicitudes_Dataset_Cuadrante(ByRef pDatos As DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "", Optional ByVal pIncidencias As String = "") As Boolean Implements PresenciaDAO.Lista_Solicitudes_Dataset_Cuadrante
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * "
            'mConsulta = "Select fecha "
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "' "
            If pIncidencias <> "" Then
                mConsulta &= " and cod_incidencia in (" & pIncidencias & ") "
            End If
            mConsulta &= " and estado <> 'R'"
            'Else
            '    mConsulta &= " and estado in  ('" & pEstado & "') "
            'End If
            mConsulta &= " and  fecha = to_date('" & pFecha_Desde & "','DD/MM/YYYY') "
            'mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY') "
            'mConsulta &= " group by fecha "
            'mConsulta &= " group by estado, fecha "
            If pOrden <> "" Then
                mConsulta &= " ORDER BY " & pOrden
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Solicitudes_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Solicitudes_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * "
            'mConsulta = "Select fecha "
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "' "
            'If pEstado = "P" Then
            'mConsulta &= " and estado in  ('E','P') "
            'Else
            '    mConsulta &= " and estado in  ('" & pEstado & "') "
            'End If
            mConsulta &= " and  fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY') "
            mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY') "
            'mConsulta &= " group by fecha "
            'mConsulta &= " group by estado, fecha "
            If pOrden <> "" Then
                mConsulta &= " ORDER BY " & pOrden
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Solicitudes_Dataset_Completo(ByRef pDatos As System.Data.DataSet, ByVal pID_Usuario As String, Optional ByVal pFecha_Desde As String = "", Optional ByVal pFecha_Hasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCod_Incidencia As String = "", Optional ByVal pAgrupadoPorFecha As Boolean = True, Optional ByVal pOrden As String = "", Optional ByVal pCod_Solicitud As String = "") As Boolean Implements PresenciaDAO.Lista_Solicitudes_Dataset_Completo
        Dim mConsulta As String
        Dim mWhere As String

        Try

            If pAgrupadoPorFecha Then
                mConsulta = "Select fecha "
            Else
                mConsulta = "Select * "
            End If
            mConsulta &= " from solicitud "
            mConsulta &= " where dni = '" & pID_Usuario & "' "
            If pListaEstados <> "" Then
                mConsulta &= " and estado in  (" & pListaEstados & ") "
            End If
            If pCod_Incidencia <> "" Then
                mConsulta &= " and cod_incidencia = " & pCod_Incidencia
            End If
            'mConsulta &= " and  fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY') "
            'mConsulta &= " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY') "
            If pFecha_Desde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFecha_Desde & "','DD/MM/YYYY')"
            End If
            If pFecha_Hasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFecha_Hasta & "','DD/MM/YYYY')"
            End If

            If pCod_Solicitud <> "" Then
                mConsulta &= " AND codigo <> " & pCod_Solicitud
            End If

            If pAgrupadoPorFecha Then
                mConsulta &= " group by fecha "
            End If



            If pOrden <> "" Then
                mConsulta &= " ORDER BY " & pOrden
            End If
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Detalle_Resumen_Justificaciones_Dataset(ByRef pDatos As System.Data.DataSet, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pDNI As String, ByVal pCod_Incidencia As String) As Boolean Implements PresenciaDAO.Detalle_Resumen_Justificaciones_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = " select i.COD_INCIDENCIA, i.DESC_INCIDENCIA, fecha_justificada, desde_minutos, hasta_minutos, observaciones "
            mConsulta &= " from justificaciones_t J, incidencias I "
            mConsulta &= " where I.cod_incidencia = J.cod_incidencia and  dni_empl = '" & pDNI & "'"
            mConsulta &= " and fecha_justificada >= TO_DATE ('" & pFechaDesde & "', 'DD/MM/YYYY')  "
            mConsulta &= " and fecha_justificada <= TO_DATE ('" & pFechaHasta & "', 'DD/MM/YYYY') "
            mConsulta &= " and J.cod_incidencia =" & pCod_Incidencia
            mConsulta &= " order by fecha_justificada "

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_AutorizadoJustificar_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Actualiza_Noticias(ByVal pCod_Noticia As String, Optional ByVal pDesc_Noticia As String = "", Optional ByVal pFecha As String = "", Optional ByVal Obs_Noticia As String = "") As Boolean Implements PresenciaDAO.Actualiza_Noticias
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String

        Try

            mSQL = "update noticias set "


            If pDesc_Noticia <> "" Then
                mWhere &= " Desc_Noticia ='" & pDesc_Noticia & "'"
            End If

            If pFecha <> "" Then
                If mWhere <> "" Then mWhere &= ", "
                mWhere &= " fecha = '" & pFecha & "'"
            End If

            If Obs_Noticia = "" Then
                If mWhere <> "" Then mWhere &= ", "
                mWhere &= " Obs_Noticia = null"
            Else
                If mWhere <> "" Then mWhere &= ", "
                mWhere &= " Obs_Noticia = '" & Obs_Noticia & "'"
            End If

            mSQL &= mWhere
            mSQL &= " where cod_noticia = " & pCod_Noticia
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Noticias", ex, mSQL)
        End Try

    End Function

    Public Function Elimina_Noticias(ByVal pCod_Noticia As String) As Boolean Implements PresenciaDAO.Elimina_Noticias
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            mSQL = "delete from noticias "
            mSQL &= "where cod_noticia = " & pCod_Noticia

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_TCIncidencia", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_Noticias(ByVal pDesc_Noticia As String, ByVal pFecha As String, Optional ByVal Obs_Noticia As String = "") As Integer Implements PresenciaDAO.Inserta_Noticias
        Dim mConsulta As String
        Dim mWhere As String

        Dim mSQL As String
        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mauxdbl As Integer


        Try
            '***************************************************
            'Buscamos el último código
            '***************************************************

            mSQL = "SELECT max(cod_noticia) FROM noticias"

            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mSQL
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) Is DBNull.Value Then
                    mauxdbl = 1
                Else
                    mauxdbl = mReader(0) + 1
                End If
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into noticias (cod_noticia, desc_noticia, fecha, obs_noticia) " _
                & " values (" & mauxdbl & ","
            mConsulta &= "'" & pDesc_Noticia & "','" & pFecha & "',"
            mConsulta &= IIf(Obs_Noticia <> "", "'" & Obs_Noticia & "'", "null") & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return mauxdbl
        Catch ex As Exception
            Trata_Error("Error en Inserta_Noticias", ex, mConsulta)
        End Try
    End Function

    Public Function Lista_Noticias_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCod_Noticia As String = "", Optional ByVal pDesc_Noticia As String = "", Optional ByVal pFecha As String = "", Optional ByVal pNumero_Max As String = "") As Boolean Implements PresenciaDAO.Lista_Noticias_Dataset
        Dim mConsulta As String
        Try
            If pNumero_Max <> "" Then mConsulta = "select * from ("

            mConsulta &= "SELECT * FROM noticias"
            Dim mWhere As String
            If pCod_Noticia <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " cod_noticia = " & pCod_Noticia
            End If
            If pDesc_Noticia <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " upper(desc_noticia) like '" & UCase(pDesc_Noticia) & "%'"
            End If
            If pFecha <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                End If
                mWhere = mWhere & " fecha = to_date ('" & pFecha & "', 'DD/MM/YYYY')"
            End If
            If mWhere <> "" Then mConsulta = mConsulta & " WHERE " & mWhere
            mConsulta &= " order by fecha desc"

            If pNumero_Max <> "" Then
                mConsulta &= ") where rownum <=" & pNumero_Max
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Noticias_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Actualiza_Asignacion_Responsable(ByVal pID_Responsable_Antiguo As String, ByVal pID_Responsable_Nuevo As String) As Boolean Implements PresenciaDAO.Actualiza_Asignacion_Responsable
        Dim mConsulta As String


        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim Existe As Boolean = False
        Try


            mConsulta = "update Asignacion_Responsable set id_responsable='" & pID_Responsable_Nuevo & "' where "
            mConsulta &= " id_responsable ='" & pID_Responsable_Antiguo & "'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()


            'puede ser que perteneciera a la persona antigua, y entonces quedaria el mismo como su responsable
            'asi que borramos su posible asignacion cuando esto ocurra
            'En el caso que no perteneciera, el responsable nuevo pertenece a alguien y debe seguir con la asignación
            mConsulta = "delete asignacion_responsable  where id_usuario='" & pID_Responsable_Nuevo & "' and id_responsable ='" & pID_Responsable_Nuevo & "'"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()



            'Si cambiamos el responsable, todas las solicitudes que estaban pendientes del anterior responsable, pasan al nuevo
            mConsulta = "update solicitud set id_siguiente_responsable='" & pID_Responsable_Nuevo & "' where id_siguiente_responsable='" & pID_Responsable_Antiguo & "' and estado ='P'"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()


            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Responsable", ex, mConsulta)
        End Try


    End Function

    Function Mantiene_Asignacion_Responsable_Siguiente(ByVal pID_Responsable) As Boolean Implements PresenciaDAO.Mantiene_Asignacion_Responsable_Siguiente
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mDatos As New DataSet


        Try

            mConsulta = "SELECT siguiente from asignacion_responsable "
            mConsulta &= " where id_responsable = '" & pID_Responsable & "' and siguiente = '1' and id_usuario not like 'G%'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            mConsulta = "update asignacion_responsable set "
            If mDatos.Tables(0).Rows.Count > 0 Then
                mConsulta &= " siguiente= '1'"
            Else
                mConsulta &= " siguiente= null"
            End If
            mDatos.Clear()
            mDatos = Nothing

            mConsulta &= " where id_responsable ='" & pID_Responsable & "' and id_usuario not like 'G%'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Mantiene_Asignacion_Responsable", ex, mConsulta)
        End Try

    End Function


    Public Function Lista_Acumuladores_Dataset(ByRef pDatos As System.Data.DataSet) As Boolean Implements PresenciaDAO.Lista_Acumuladores_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * FROM ACUMULADORES"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Acumuladores_Fijos_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pFavoritos As Boolean = False) As Boolean Implements PresenciaDAO.Lista_Acumuladores_Fijos_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * FROM ACUMULADOR_FIJO "
            If pFavoritos Then
                mConsulta &= " where not favoritos is null order by favoritos"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_Acumulador_Fijo(ByVal pCod As String, Optional ByVal pFavoritos As String = "", Optional ByVal pFormato As String = "") As Boolean Implements PresenciaDAO.Actualiza_Acumulador_Fijo
        Dim mConsulta As String


        Dim mCommand As New OleDb.OleDbCommand
        Try


            mConsulta = "update Acumulador_Fijo set "
            If pFavoritos <> "" Then
                mConsulta &= " favoritos = " & pFavoritos
            Else
                mConsulta &= " favoritos = null"
            End If

            If pFormato <> "" Then
                mConsulta &= " ,formato = '" & pFormato & "'"
            Else
                mConsulta &= ", formato = 'H'"  'Por defecto lo ponemos en formato hora
            End If

            mConsulta &= " where codigo ='" & pCod & "'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Responsable", ex, mConsulta)
        End Try

    End Function

    Public Function Actualiza_Acumulador_Def(ByVal pCod As String, ByVal pDesc_Acumulador As String, Optional ByVal pSeleccion As String = "", Optional ByVal pPeriodicidad As String = "", Optional ByVal pFormato As String = "", Optional ByVal pIntervalo As String = "", Optional ByVal pDesc_Long As String = "", Optional ByVal pFavoritos As String = "") As Boolean Implements PresenciaDAO.Actualiza_Acumulador_Def
        Dim mConsulta As String


        Dim mCommand As New OleDb.OleDbCommand
        Try


            mConsulta = "update Acumuladores set "
            mConsulta &= " desc_acumulador= '" & pDesc_Acumulador & "'"
            mConsulta &= ", seleccion= '" & pSeleccion & "'"
            mConsulta &= ", periodicidad= '" & pPeriodicidad & "'"
            mConsulta &= ", formato= '" & pFormato & "'"
            mConsulta &= ", intervalo= '" & pIntervalo & "'"


            If pDesc_Long <> "" Then
                mConsulta &= " ,desc_long = '" & pDesc_Long & "'"
            Else
                mConsulta &= ", desc_long = null"
            End If

            If pFavoritos <> "" Then
                mConsulta &= " ,favoritos = " & pFavoritos
            Else
                mConsulta &= ", favoritos = 0"
                'es el valor por defecto de la tabla
            End If


            mConsulta &= " where cod_acumulador =" & pCod

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Responsable", ex, mConsulta)
        End Try

    End Function

    Public Function Consulta_Dni_Incidencias(ByRef pDatos As System.Data.DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Dni_Incidencias
        Dim mConsulta As String
        Dim mWhere As String

        Try

            If pCod_Grupo = "" Then
                mConsulta = "select dni dni_empl, ape1, ape2, nombre from empleados order by ape1, ape2, nombre"
            Else
                mConsulta = "select distinct(p.dni_empl),ape1, ape2, nombre from justificaciones j, pertenecena p, empleados e where  p.DNI_EMPL = j.DNI_EMPL  (+)"
                mConsulta &= " and p.dni_empl=e.dni and p.COD_GRUPO in (" & pCod_Grupo & ") and p.TIPO_GRP = 1 "
                mConsulta &= " and  j.FECHA_JUSTIFICADA(+) >= '" & Fecha_Ini & "' and j.FECHA_JUSTIFICADA(+) <= '" & Fecha_Fin & "'"
                mConsulta &= " and cod_incidencia (+)=" & pCod_Inc & " order by e.ape1"
            End If


            'mConsulta = "select distinct(dni_empl),ape1 from justificaciones j, empleados where j.dni_empl = empleados.dni "
            'mConsulta &= " and j.FECHA_JUSTIFICADA >= '" & Fecha_Ini & "' and j.FECHA_JUSTIFICADA <= '" & Fecha_Fin & "'"
            'mConsulta &= " and cod_incidencia =" & pCod_Inc
            'mConsulta &= " order by empleados.ape1"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Consulta_Datos_Incidencias(ByRef pDatos As System.Data.DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Incidencias
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni_empl, fecha_justificada, cod_incidencia "
            mConsulta &= " from justificaciones_t j, pertenecena p where   p.DNI_EMPL = j.DNI_EMPL "
            If pCod_Grupo = "" Then
                mConsulta &= " and p.TIPO_GRP = 1"
            Else
                mConsulta &= " and p.COD_GRUPO in (" & pCod_Grupo & ") and p.TIPO_GRP = 1"
            End If
            mConsulta &= " and j.FECHA_JUSTIFICADA >= '" & Fecha_Ini & "' and j.FECHA_JUSTIFICADA <= '" & Fecha_Fin & "'"
            mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            mConsulta &= " group by j.dni_empl, fecha_justificada , cod_incidencia"
            mConsulta &= " order by j.dni_empl, fecha_justificada, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Consulta_Datos_Incidencias_DNI(ByRef pDatos As System.Data.DataSet, ByVal pDNI As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Incidencias_DNI
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni_empl, fecha_justificada, cod_incidencia "
            mConsulta &= " from justificaciones_t j, pertenecena p where   p.DNI_EMPL = j.DNI_EMPL "
            mConsulta &= "and j.DNI_EMPL in (" & pDNI & ")"
            mConsulta &= " and j.FECHA_JUSTIFICADA >= '" & Fecha_Ini & "' and j.FECHA_JUSTIFICADA <= '" & Fecha_Fin & "'"
            mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            mConsulta &= " group by j.dni_empl, fecha_justificada , cod_incidencia"
            mConsulta &= " order by j.dni_empl, fecha_justificada, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Consulta_Datos_Solicitudes(ByRef pDatos As System.Data.DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Solicitudes
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni, fecha, cod_incidencia "
            mConsulta &= " from solicitud j, pertenecena p, asignacion_responsable where   ((p.DNI_EMPL = j.DNI "
            If pCod_Grupo = "" Then
                mConsulta &= " and p.TIPO_GRP = 1)"
            Else
                mConsulta &= " and p.COD_GRUPO in ('" & pCod_Grupo & "') and p.TIPO_GRP = 1)"
            End If
            Dim grupos() As String
            Dim i As Integer
            grupos = pCod_Grupo.Split(",")
            If grupos.Length > 1 Then
                For i = 0 To grupos.Length - 1
                    If i = 0 Then
                        mConsulta &= " or (asignacion_responsable.id_responsable = j.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0'))) "
                    ElseIf i = grupos.Length - 1 Then
                        mConsulta &= " or (asignacion_responsable.id_responsable = j.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0'))))"
                    Else
                        mConsulta &= " or (asignacion_responsable.id_responsable = j.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0')))"
                    End If

                Next
            ElseIf grupos.Length = 1 Then
                mConsulta &= " or (asignacion_responsable.id_responsable = j.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & pCod_Grupo & ",4,'0'))))"
            End If
            '            mConsulta &= " or (asignacion_responsable.id_responsable = j.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & pCod_Grupo & ",4,'0')))) "
            mConsulta &= " and j.FECHA >= '" & Fecha_Ini & "' and j.FECHA <= '" & Fecha_Fin & "'"
            mConsulta &= " and j.estado in ('E','A','R')"
            If pCod_Inc <> "" Then
                mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            End If
            mConsulta &= " group by j.dni, fecha, cod_incidencia"
            mConsulta &= " order by j.dni, fecha, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Solicitudes", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Consulta_Datos_Solicitudes_DNI(ByRef pDatos As System.Data.DataSet, ByVal pDNIs As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Solicitudes_DNI
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni, fecha, cod_incidencia "
            mConsulta &= " from solicitud j, pertenecena p where   p.DNI_EMPL = j.DNI "
            mConsulta &= "and j.DNI in (" & pDNIs & ")"
            mConsulta &= " and j.FECHA >= '" & Fecha_Ini & "' and j.FECHA <= '" & Fecha_Fin & "'"
            mConsulta &= " and j.estado in ('E','A','R')"
            mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            mConsulta &= " group by j.dni, fecha, cod_incidencia"
            mConsulta &= " order by j.dni, fecha, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Solicitudes", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Consulta_Dni_GruposTrabajo(ByRef pDatos As System.Data.DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Grupo As String) As Boolean Implements PresenciaDAO.Consulta_Dni_GruposTrabajo
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'solo nos hacen falta los usuarios que pertenezcan a los grupos especificados. Porque despues ya coge el grupo correspondiente
            'por cada dia del mes seleccionado.
            'En la otra consulta, la de incidencias, funciona bien los left join y se puede dejar la consulta como está, 
            'aunque realmente solo hacen falta los dni de las personas en la primera consulta
            If pCod_Grupo = "" Then
                mConsulta = "select dni dni_empl, ape1, ape2, nombre from empleados order by ape1, ape2, nombre"
            Else

                mConsulta = "select  distinct(e.dni) as dni_empl, ape1, ape2, nombre "
                mConsulta &= " from pertenecena p ,asociausuariogrupotrabajo a, empleados e,asignacion_responsable  "
                mConsulta &= " where (p.DNI_EMPL = a.DNI_EMPL and  p.DNI_EMPL= e.DNI "
                mConsulta &= " and p.TIPO_GRP =1 and p.COD_GRUPO in (" & pCod_Grupo & ")) or "

                Dim grupos() As String
                Dim i As Integer
                grupos = pCod_Grupo.Split(",")
                If grupos.Length > 1 Then
                    For i = 0 To grupos.Length - 1
                        If i = 0 Then
                            mConsulta &= "(asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0')))"
                        Else
                            mConsulta &= " or (asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0')))"
                        End If

                    Next
                ElseIf grupos.Length = 1 Then
                    mConsulta &= "(asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & pCod_Grupo & ",4,'0')))"
                End If


                mConsulta &= " order by ape1, ape2, nombre"
            End If

            'esta consulta solo muestra las personas del grupo seleccionado que tengan algun grupode trabajo en la fecha. No nos vale, porque nos
            'hacen falta todas las personas del grupo de consulta, por eso usamos la de arriba

            '            mConsulta = "select distinct(dni_empl), ape1, ape2, nombre from asociausuariogrupotrabajo a, empleados e "
            '            mConsulta &= " where a.DNI_EMPL= e.DNI and ((FECHA_DESDE >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
            '            & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            '            mConsulta &= " and a.dni_empl in  (select p.dni_empl from pertenecena p where p.TIPO_GRP =1 and p.COD_GRUPO in (" & pCod_Grupo & ") )"
            '            mConsulta &= "order by ape1, ape2, nombre"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Consulta_Dni_GruposConsulta(ByRef pDatos As System.Data.DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Grupo As String) As Boolean Implements PresenciaDAO.Consulta_Dni_GruposConsulta
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'solo nos hacen falta los usuarios que pertenezcan a los grupos especificados. Porque despues ya coge el grupo correspondiente
            'por cada dia del mes seleccionado.
            'En la otra consulta, la de incidencias, funciona bien los left join y se puede dejar la consulta como está, 
            'aunque realmente solo hacen falta los dni de las personas en la primera consulta
            If pCod_Grupo = "" Then
                mConsulta = "select dni dni_empl, ape1, ape2, nombre from empleados order by ape1, ape2, nombre"
            Else

                mConsulta = "select  distinct(e.dni) as dni_empl, ape1, ape2, nombre "
                mConsulta &= " from pertenecena p ,empleados e,asignacion_responsable  "
                mConsulta &= " where (p.DNI_EMPL= e.DNI "
                mConsulta &= " and p.TIPO_GRP =1 and p.COD_GRUPO in (" & pCod_Grupo & ")) or "

                Dim grupos() As String
                Dim i As Integer
                grupos = pCod_Grupo.Split(",")
                If grupos.Length > 1 Then
                    For i = 0 To grupos.Length - 1
                        If i = 0 Then
                            mConsulta &= "(asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0')))"
                        Else
                            mConsulta &= " or (asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & grupos(i) & ",4,'0')))"
                        End If

                    Next
                ElseIf grupos.Length = 1 Then
                    mConsulta &= "(asignacion_responsable.id_responsable = e.dni and asignacion_responsable.ID_USUARIO in ('GRP' || lpad(" & pCod_Grupo & ",4,'0')))"
                End If


                mConsulta &= " order by ape1, ape2, nombre"
            End If

            'esta consulta solo muestra las personas del grupo seleccionado que tengan algun grupode trabajo en la fecha. No nos vale, porque nos
            'hacen falta todas las personas del grupo de consulta, por eso usamos la de arriba

            '            mConsulta = "select distinct(dni_empl), ape1, ape2, nombre from asociausuariogrupotrabajo a, empleados e "
            '            mConsulta &= " where a.DNI_EMPL= e.DNI and ((FECHA_DESDE >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
            '            & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            '            mConsulta &= " and a.dni_empl in  (select p.dni_empl from pertenecena p where p.TIPO_GRP =1 and p.COD_GRUPO in (" & pCod_Grupo & ") )"
            '            mConsulta &= "order by ape1, ape2, nombre"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Consulta_Dni_Cuadrante(ByRef pDatos As System.Data.DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pDNIs As String) As Boolean Implements PresenciaDAO.Consulta_Dni_Cuadrante
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'solo nos hacen falta los usuarios que pertenezcan a los grupos especificados. Porque despues ya coge el grupo correspondiente
            'por cada dia del mes seleccionado.
            'En la otra consulta, la de incidencias, funciona bien los left join y se puede dejar la consulta como está, 
            'aunque realmente solo hacen falta los dni de las personas en la primera consulta
            If pDNIs <> "" Then
                mConsulta = "select dni dni_empl, ape1, ape2, nombre from empleados"
                mConsulta = mConsulta & " WHERE dni in (" & pDNIs & ")"
                mConsulta = mConsulta & " order by ape1, ape2, nombre"
            End If

            'esta consulta solo muestra las personas del grupo seleccionado que tengan algun grupode trabajo en la fecha. No nos vale, porque nos
            'hacen falta todas las personas del grupo de consulta, por eso usamos la de arriba

            '            mConsulta = "select distinct(dni_empl), ape1, ape2, nombre from asociausuariogrupotrabajo a, empleados e "
            '            mConsulta &= " where a.DNI_EMPL= e.DNI and ((FECHA_DESDE >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
            '            & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha_Fin), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha_Ini), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            '            mConsulta &= " and a.dni_empl in  (select p.dni_empl from pertenecena p where p.TIPO_GRP =1 and p.COD_GRUPO in (" & pCod_Grupo & ") )"
            '            mConsulta &= "order by ape1, ape2, nombre"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Consulta_Datos_GruposTrabajo(ByVal Fecha As String, ByVal pDni As String) As String Implements PresenciaDAO.Consulta_Datos_GruposTrabajo
        Dim mConsulta As String
        Dim Valor As String
        Dim mDatos As New DataSet

        Try

            mConsulta = "select cod_grupotrabajo, fecha_desde, fecha_hasta, dni_empl from asociausuariogrupotrabajo "
            mConsulta &= " where ((FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            mConsulta &= " and dni_empl = '" & pDni & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            'Dim oObj As Object
            'oObj = mDatos.Tables(0).Compute("Min(Cod_GrupoTrabajo)", "((FECHA_DESDE <= '" & Format(CDate(Fecha), "dd/MM/yyyy") & "' and FECHA_HASTA >= '" & Format(CDate(Fecha), "dd/MM/yyyy") & "') or (FECHA_DESDE <= '" & Format(CDate(Fecha), "dd/MM/yyyy") & "' and FECHA_HASTA is null)) and dni_empl = '" & pDni & "'")

            If mDatos.Tables(0).Rows.Count > 0 Then
                Valor = mDatos.Tables(0).Rows(0)("cod_grupotrabajo")
            Else
                Valor = ""
            End If
            mDatos.Clear()
            mDatos = Nothing
            Return Valor

        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias", ex, mConsulta)
        End Try

    End Function

    Public Function Consulta_Datos_GruposTrabajo2(ByRef pDatos As DataSet) As Boolean Implements PresenciaDAO.Consulta_Datos_GruposTrabajo2
        Dim mConsulta As String



        Try

            mConsulta = "select cod_grupotrabajo, fecha_desde, fecha_hasta, dni_empl from asociausuariogrupotrabajo "
            'mConsulta &= " where ((FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(Fecha), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null)) "
            ' mConsulta &= " ((FECHA_DESDE >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_DESDE <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_HASTA >= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA <= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or " _
            '                & "(FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA >= TO_DATE('" & Format(CDate(pHasta), "dd/MM/yyyy") & "','DD/MM/YYYY')) or (FECHA_DESDE <= TO_DATE('" & Format(CDate(pDesde), "dd/MM/yyyy") & "','DD/MM/YYYY') and FECHA_HASTA is null))"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Consulta_Datos_GruposTrabajo2", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_PromaInf(ByRef pDatos As System.Data.DataSet, ByVal pConsulta As String) As Boolean Implements PresenciaDAO.Lista_PromaInf
        Try

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(pConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_PromaInf", ex, pConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Usuarios_GrupoConsulta(ByRef pDatos As System.Data.DataSet, ByVal pCodigo As String, Optional ByVal pDistinct As String = "") As Boolean Implements PresenciaDAO.Lista_Usuarios_GrupoConsulta
        Dim mConsulta As String
        Dim mWhere As String

        Try

            If pDistinct = "" Then
                mConsulta = "SELECT dni, nombre, ape1, ape2, cod_grupo, codpertenece"
                mConsulta = mConsulta & " From PERTENECENA, empleados"
                mConsulta = mConsulta & " WHERE dni_empl = dni and COD_GRUPO in (" & pCodigo & ") and tipo_grp=1"
                mConsulta = mConsulta & " order by ape1, ape2, nombre"
            Else
                mConsulta = "SELECT dni, nombre, ape1, ape2"
                mConsulta = mConsulta & " From PERTENECENA, empleados"
                mConsulta = mConsulta & " WHERE dni_empl = dni and COD_GRUPO in (" & pCodigo & ") and tipo_grp=1"
                mConsulta &= " group by dni, ape1, ape2, nombre "
                mConsulta = mConsulta & " order by ape1, ape2, nombre"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Usuarios_GrupoConsulta", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lee_Diario(ByRef pDatos As DataSet, ByVal pDNI As String, ByVal pFecha As String) As Boolean Implements PresenciaDAO.Lee_Diario
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "SELECT * from diario"
            mConsulta = mConsulta & " WHERE dni = '" & pDNI & "' and fecha ='" & pFecha & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lee_Diario", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_Asignacion_Responsable_Siguiente(ByVal pID_Responsable As Object) As Boolean Implements PresenciaDAO.Actualiza_Asignacion_Responsable_Siguiente
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mDatos As New DataSet
        Dim mTiene_Siguiente As String

        Try

            mTiene_Siguiente = "0"

            mConsulta = "SELECT siguiente from asignacion_responsable "
            mConsulta &= " where id_responsable = '" & pID_Responsable & "' and siguiente = '1' and id_usuario not like 'G%'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            mConsulta = "update asignacion_responsable set "
            If mDatos.Tables(0).Rows.Count > 0 Then
                mConsulta &= " siguiente= null"
            Else
                mConsulta &= " siguiente= '1'"
            End If
            mDatos.Clear()
            mDatos = Nothing

            mConsulta &= " where id_responsable ='" & pID_Responsable & "' and id_usuario not like 'G%'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Responsable", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Empleados_Responsable_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Lista_Empleados_Responsable_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2,"
            mConsulta &= " (select id_responsable || ' ' || nombre || ' ' || ape1 from asignacion_responsable, empleados where dni = id_responsable and id_usuario = e.dni) responsable "
            mConsulta &= " from empleados e "

            If pDNI <> "" Then
                mWhere = " where upper(dni) like '" & UCase(pDNI) & "%'"
            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then mConsulta &= mWhere
            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados_Responsable_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function


    Public Function Lista_Asignacion_Responsable_Por_Grupo(ByRef pDatos As System.Data.DataSet, Optional ByVal pGrupo As String = "", Optional ByVal pResponsable As String = "") As Boolean Implements PresenciaDAO.Lista_Asignacion_Responsable_Por_Grupo
        'da lista de asignaciones de responsables
        Dim mConsulta As String
        Dim mWhere As String
        Dim grupos_usuario() As String

        Try
            grupos_usuario = pGrupo.Split(",")
            'devolvemos todas las asignaciones de responsable por grupos, que empiezan por GRP
            'enlazamos id_responsable porque id_usuario es grp0001 por ejemplo, es un grupo, y
            'no un empleado. Los empleados son los responsables del grupo
            mConsulta = "SELECT ID_USUARIO, ID_RESPONSABLE, dni, nombre, ape1, ape2 "
            mConsulta = mConsulta & " FROM asignacion_responsable"
            mConsulta = mConsulta & " , empleados "
            mWhere = " WHERE asignacion_responsable.ID_responsable = empleados.dni and id_usuario like 'GRP%'"
            If pGrupo <> "" Then
                If grupos_usuario.Length > 1 Then
                    Dim i As Integer = 0
                    mWhere = mWhere & " AND ID_USUARIO in ("
                    While i < grupos_usuario.Length
                        If i >= 1 Then
                            mWhere = mWhere & ","
                        End If
                        mWhere = mWhere & "'GRP" & Format(CInt(grupos_usuario(i).ToString()), "0000") & "'"
                        i = i + 1
                    End While
                    mWhere = mWhere & ")"
                Else
                    mWhere = mWhere & " AND ID_USUARIO in ('GRP" & Format(CInt(pGrupo), "0000") & "')"
                End If
                'mWhere = mWhere & " AND ID_USUARIO in ('GRP" & Format(CInt(pGrupo), "0000") & "')"
            End If
            If pResponsable <> "" Then
                mWhere = mWhere & " AND ID_RESPONSABLE in ('" & pResponsable & "')"
            End If

            mConsulta = mConsulta & mWhere & " order by empleados.ape1, empleados.ape2, empleados.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Asignacion_Responsable_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Asignacion_Consultor_Por_Grupo(ByRef pDatos As System.Data.DataSet, Optional ByVal pGrupo As String = "", Optional ByVal pResponsable As String = "") As Boolean Implements PresenciaDAO.Lista_Asignacion_Consultor_Por_Grupo
        'da lista de asignaciones de responsables
        Dim mConsulta As String
        Dim mWhere As String
        Dim grupos_usuario() As String

        Try
            grupos_usuario = pGrupo.Split(",")
            'devolvemos todas las asignaciones de responsable por grupos, que empiezan por GRP
            'enlazamos id_responsable porque id_usuario es grp0001 por ejemplo, es un grupo, y
            'no un empleado. Los empleados son los responsables del grupo
            mConsulta = "SELECT GRUPO, ID_RESPONSABLE, dni, nombre, ape1, ape2 "
            mConsulta = mConsulta & " FROM AUTORIZADOCONSULTAR"
            mConsulta = mConsulta & " , empleados "
            mWhere = " WHERE AUTORIZADOCONSULTAR.ID_responsable = empleados.dni"
            If pGrupo <> "" Then
                If grupos_usuario.Length > 1 Then
                    Dim i As Integer = 0
                    mWhere = mWhere & " AND GRUPO in ("
                    While i < grupos_usuario.Length
                        If i >= 1 Then
                            mWhere = mWhere & ","
                        End If
                        mWhere = mWhere & "'" & grupos_usuario(i).ToString() & "'"
                        i = i + 1
                    End While
                    mWhere = mWhere & ")"
                Else
                    mWhere = mWhere & " AND GRUPO in ('" & pGrupo & "')"
                End If
                'mWhere = mWhere & " AND ID_USUARIO in ('GRP" & Format(CInt(pGrupo), "0000") & "')"
            End If
            If pResponsable <> "" Then
                mWhere = mWhere & " AND ID_RESPONSABLE in ('" & pResponsable & "')"
            End If

            mConsulta = mConsulta & mWhere & " order by empleados.ape1, empleados.ape2, empleados.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Asignacion_Consultor_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Asignacion_Responsable_Por_Grupo(Optional ByVal pGrupo As String = "", Optional ByVal pID_Responsable As String = "") As Boolean Implements PresenciaDAO.Elimina_Asignacion_Responsable_Por_Grupo
        Dim mConsulta As String
        Dim mWhere As String
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "DELETE asignacion_responsable "
            If pGrupo <> "" Then
                mWhere = " WHERE ID_USUARIO = 'GRP" & pGrupo & "'"
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & mWhere
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asignacion_Responsable_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Elimina_Asignacion_Consultor_Por_Grupo(Optional ByVal pGrupo As String = "", Optional ByVal pID_Responsable As String = "") As Boolean Implements PresenciaDAO.Elimina_Asignacion_Consultor_Por_Grupo
        Dim mConsulta As String
        Dim mWhere As String
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "DELETE AUTORIZADOCONSULTAR "
            If pGrupo <> "" Then
                mWhere = " WHERE GRUPO = '" & pGrupo & "'"
            End If
            If pID_Responsable <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " ID_RESPONSABLE = '" & pID_Responsable & "'"
            End If
            mConsulta = mConsulta & mWhere
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asignacion_Consultor_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Empleados_Responsables_Grupos_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Lista_Empleados_Responsables_Grupos_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select dni, nombre, ape1, ape2 , desc_grupo grupo" _
                        & " from empleados e,  " _
                        & "(select id_responsable, desc_grupo  from gruposconsulta g, asignacion_responsable where cod_grupo = to_number(substr(id_usuario, 4, 4))  and id_usuario like 'GRP%') grupo " _
                        & " where id_responsable (+) = e.dni "

            If pDNI <> "" Then
                mWhere = " AND upper(dni) like '" & UCase(pDNI) & "%'"
            End If

            If pNombre <> "" Then
                mWhere = mWhere & " AND SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                mWhere = mWhere & " AND SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                mWhere = mWhere & " AND SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pApellidos <> "" Then
                mWhere = mWhere & " AND SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If


            If mWhere <> "" Then mConsulta &= mWhere
            mConsulta &= " order by ape1, ape2, nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Empleados_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Inserta_Asignacion_Responsable_Por_Grupo(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Inserta_Asignacion_Responsable_Por_Grupo
        Dim mConsulta As String
        Dim mWhere As String
        Try


            'antes buscamos si alguien que esté en ese grupo tiene el campo siguiente a 1
            'si es así, éste también debe estar a uno.
            Dim mDatos As DataSet
            Dim mConsulta_Siguiente As String
            Dim mSiguiente As String
            mConsulta_Siguiente = "select * from asignacion_responsable where id_usuario = 'GRP" & Format(CInt(pID_Usuario), "0000") & "' and " _
                    & " siguiente  = 1"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta_Siguiente, mConexion)
            mDatos = New DataSet
            mDataAdapter.Fill(mDatos)


            If mDatos.Tables(0).Rows.Count > 0 Then
                mSiguiente = "'1'"
            Else
                mSiguiente = "null"
            End If
            mDatos.Clear()
            mDatos = Nothing

            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "insert into asignacion_responsable values"
            mConsulta &= "('GRP" & Format(CInt(pID_Usuario), "0000") & "','" & pID_Responsable & "'," & mSiguiente & ")"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asignacion_Responsable_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Inserta_Asignacion_Consultor_Por_Grupo(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean Implements PresenciaDAO.Inserta_Asignacion_Consultor_Por_Grupo
        Dim mConsulta As String
        Dim mWhere As String
        Try


            'antes buscamos si alguien que esté en ese grupo tiene el campo siguiente a 1
            'si es así, éste también debe estar a uno.
            Dim mDatos As DataSet
            Dim mConsulta_Siguiente As String
            Dim mSiguiente As String
            mConsulta_Siguiente = "select * from AUTORIZADOCONSULTAR where grupo = 'GRP" & Format(CInt(pID_Usuario), "0000") & "'"
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta_Siguiente, mConexion)
            mDatos = New DataSet
            mDataAdapter.Fill(mDatos)


            'If mDatos.Tables(0).Rows.Count > 0 Then
            'mSiguiente = "'1'"
            'Else
            '    mSiguiente = "null"
            'End If
            mDatos.Clear()
            mDatos = Nothing

            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "insert into AUTORIZADOCONSULTAR values"
            mConsulta &= "('" & pID_Usuario & "','" & pID_Responsable & "')"
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_Asignacion_Consultor_Por_Grupo", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function EjecutaConsultaSQL(ByVal pSQL As String, ByRef pDataSet As DataSet, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.EjecutaConsultaSQL
        Try
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(pSQL, mConexion)
            mDataAdapter.Fill(pDataSet)
            pError = ""
            Return True
        Catch ex As Exception
            pError = ex.Message
            Trata_Error("Error en EjecutaConsultaSQL", ex, pSQL)
            Return False
        End Try

    End Function

    Public Function Actualiza_Asignacion_Responsable_Siguiente_Por_Grupo(ByVal pGrupo As Object) As Boolean Implements PresenciaDAO.Actualiza_Asignacion_Responsable_Siguiente_Por_Grupo
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mDatos As New DataSet


        Try


            mConsulta = "SELECT siguiente from asignacion_responsable "
            mConsulta &= " where id_usuario = 'GRP" & Format(CInt(pGrupo), "0000") & "' and siguiente = '1'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            mConsulta = "update asignacion_responsable set "
            If mDatos.Tables(0).Rows.Count > 0 Then
                mConsulta &= " siguiente= null"
            Else
                mConsulta &= " siguiente= '1'"
            End If
            mDatos.Clear()
            mDatos = Nothing

            mConsulta &= " where id_usuario = 'GRP" & Format(CInt(pGrupo), "0000") & "'"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Responsable_Por_Grupo", ex, mConsulta)
        End Try
    End Function

    Public Function Actualiza_Asignacion_Consultor_Por_Grupo(ByVal pGrupo As Object) As Boolean Implements PresenciaDAO.Actualiza_Asignacion_Consultor_Por_Grupo
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mDatos As New DataSet


        Try


            'mConsulta = "SELECT siguiente from AUTORIZADOCONSULTAR "
            'mConsulta &= " where id_usuario = 'GRP" & Format(CInt(pGrupo), "0000") & "' and siguiente = '1'"


            'Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'mDataAdapter.Fill(mDatos)

            'mConsulta = "update AUTORIZADOCONSULTAR set "
            'If mDatos.Tables(0).Rows.Count > 0 Then
            'mConsulta &= " siguiente= null"
            'Else
            '    mConsulta &= " siguiente= '1'"
            'End If
            'mDatos.Clear()
            'mDatos = Nothing

            'mConsulta &= " where id_usuario = 'GRP" & Format(CInt(pGrupo), "0000") & "'"

            'mCommand.Connection = mConexion
            'mCommand.CommandText = mConsulta
            'mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_Asignacion_Consultor_Por_Grupo", ex, mConsulta)
        End Try
    End Function

    Public Function Lista_Asignacion_Responsable_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Lista_Responsables As String = "") As Boolean Implements PresenciaDAO.Lista_Asignacion_Responsable_Dataset
        'da lista de asignaciones de responsables
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT ID_USUARIO, ID_RESPONSABLE, SIGUIENTE "
            mConsulta = mConsulta & " FROM asignacion_responsable"
            mConsulta = mConsulta & " , empleados "
            mWhere = " WHERE asignacion_responsable.ID_USUARIO = empleados.dni "
            If pID_Usuario <> "" Then
                mWhere = mWhere & " AND ID_USUARIO = '" & pID_Usuario & "'"
            End If
            If pID_Lista_Responsables <> "" Then
                mWhere = mWhere & " AND ID_RESPONSABLE IN ( '" & pID_Lista_Responsables & "')"
            End If
            'mConsulta = mConsulta & mWhere & " order by ID_USUARIO, ID_RESPONSABLE "
            mConsulta = mConsulta & mWhere & " order by empleados.ape1, empleados.ape2, empleados.nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Asignacion_Responsable_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Grupos_Consulta_Usuario(ByRef pDatos As System.Data.DataSet, ByVal pCodigo As String) As Boolean Implements PresenciaDAO.Grupos_Consulta_Usuario
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * from pertenecena,gruposconsulta"
            mConsulta &= " where pertenecena.tipo_Grp = '1' and dni_empl = '" & pCodigo & "' and gruposconsulta.cod_grupo=pertenecena.cod_grupo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Grupos_Consulta_Usuario", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Grupos_Consulta_Usuario_DIPU(ByRef pDatos As System.Data.DataSet, ByVal pCodigo As String) As Boolean Implements PresenciaDAO.Grupos_Consulta_Usuario_DIPU
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * from pertenecena,gruposconsulta"
            mConsulta &= " where pertenecena.tipo_Grp = '1' and dni_empl = '" & pCodigo & "' and gruposconsulta.cod_grupo=pertenecena.cod_grupo  and gruposconsulta.GRUPO_PADRE <> 16 and gruposconsulta.GRUPO_PADRE <> 95  and gruposconsulta.GRUPO_PADRE <> 348"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Grupos_Consulta_Usuario", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Grupos_Consulta_Usuario_DIPU_Aprobar(ByRef pDatos As System.Data.DataSet, ByVal pCodigo As String) As Boolean Implements PresenciaDAO.Grupos_Consulta_Usuario_DIPU_Aprobar
        Dim mConsulta As String
        Try
            mConsulta = "SELECT * from pertenecena,gruposconsulta"
            mConsulta &= " where pertenecena.tipo_Grp = '1' and dni_empl = '" & pCodigo & "' and gruposconsulta.cod_grupo=pertenecena.cod_grupo  and gruposconsulta.GRUPO_PADRE = 16 and gruposconsulta.GRUPO_PADRE <> 95  and gruposconsulta.GRUPO_PADRE <> 348"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Grupos_Consulta_Usuario", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Grupos_Consulta_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Boolean Implements PresenciaDAO.Lista_Grupos_Consulta_Dataset
        Dim mSQL As String
        Dim mWhere As String
        Try
            mSQL = "SELECT cod_grupo, desc_grupo, grupo_padre"
            mSQL = mSQL & " FROM gruposconsulta "
            If pCodigo > 0 Then
                mWhere = " WHERE cod_grupo = " & pCodigo
            End If
            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "Desc_grupo LIKE '" & pNombre & "'"
            End If
            If pPadre > 0 Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & "Grupo_Padre = " & pPadre
            End If

            If mWhere <> "" Then
                mSQL = mSQL & mWhere
            End If
            mSQL = mSQL & " order by desc_grupo"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mSQL, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)

            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta_dataset", ex, mSQL)
            Return False
        End Try
    End Function

    Public Function Elimina_Siguientes_Solicitud(Optional ByVal pcod_solicitud As String = "", Optional ByVal plista_dni As String = "") As Boolean Implements PresenciaDAO.Elimina_Siguientes_Solicitud
        Dim mSQL As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try

            mSQL = "delete SIGUIENTES_SOLICITUD "
            If pcod_solicitud <> "" Then
                mWhere = " cod_solicitud = " & pcod_solicitud
            End If
            If plista_dni <> "" Then
                If mWhere <> "" Then mWhere &= " and "
                mWhere &= "dni in (" & plista_dni & ")"
            End If

            If mWhere <> "" Then mSQL = mSQL & " where " & mWhere

            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            Trata_Error("Error en Elimina_Siguientes_Solicitud", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Inserta_Siguientes_Solicitud(ByVal pcod_solicitud As String, ByVal plista_dni As String) As Boolean Implements PresenciaDAO.Inserta_Siguientes_Solicitud
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim i As Integer
        '    Dim Lista() As String = Split(plista_dni, ",")
        '    Dim mdni As String

        Try

            '       For i = 0 To Lista.Length - 1
            '      mdni = Lista(i)
            mSQL = "insert into SIGUIENTES_SOLICITUD values(" & pcod_solicitud & ",'" & plista_dni & "')"
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()
            '     Next
            Return True

        Catch ex As Exception
            Trata_Error("Error en Inserta_Siguientes_Solicitud", ex, mSQL)
            Return False
        End Try

    End Function

    Public Function Lista_Siguientes_Solicitud_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pcod_solicitud As String = "", Optional ByVal plista_dni As String = "") As Boolean Implements PresenciaDAO.Lista_Siguientes_Solicitud_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT * from SIGUIENTES_SOLICITUD"
            If pcod_solicitud <> "" Then
                mWhere = " cod_solicitud = " & pcod_solicitud
            End If
            If plista_dni <> "" Then
                If mWhere <> "" Then mWhere &= " and "
                mWhere &= "dni in (" & plista_dni & ")"
            End If

            If mWhere <> "" Then mConsulta = mConsulta & " where " & mWhere

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Siguientes_Solicitud_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Lista_Delegados_Solicitud(ByVal pID_Delegado As String, ByVal pID_Solicitud As String) As Object Implements PresenciaDAO.Lista_Delegados_Solicitud
        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT Delegados.ID_Responsable, Delegados.ID_Delegado, Empleados.Ape1, Empleados.Ape2, Empleados.Nombre "
            mConsulta &= " FROM Delegados, Empleados"
            mConsulta &= " WHERE Delegados.ID_Delegado = Empleados.Dni"
            mConsulta &= " AND ID_DELEGADO = '" & pID_Delegado & "'"
            mConsulta &= " and ID_RESPONSABLE IN (select dni from siguientes_solicitud where cod_solicitud = " & pID_Solicitud & ")"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mReader = mCommand.ExecuteReader()
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Delegados_Solicitud", ex, mConsulta)
        End Try
    End Function


    Public Function Lista_Solicitudes_Grupos(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object Implements PresenciaDAO.Lista_Solicitudes_Grupos
        Dim mConsulta As String
        Dim mOrden As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, solicitud.DNI , solicitud.DNI || ' ' ||  ape1 || ' ' || ape2 || ' ' || nombre Nombre, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias, empleados"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND SOLICITUD.dni = EMPLEADOS.dni"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND SOLICITUD.dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND SOLICITUD.dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND exists (select * from siguientes_solicitud where cod_solicitud = solicitud.codigo and dni in ('" & pID_Lista_Responsables & "'))  "
            End If
            If pLista_Ultimo_Responsable <> "" Then
                mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            End If
            If pCodigoIncidencia >= 0 Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia =  " & pCodigoIncidencia
            End If
            If Not Cambio_Grupo Is Nothing Then
                mConsulta = mConsulta & " AND cambio_grupo =  '" & Cambio_Grupo & "'"
            End If
            If pOrden = "Usuario" Then
                mOrden = " ORDER BY ape1,ape2, nombre, fecha, desde, codigo"
            ElseIf pOrden = "Fecha" Then
                mOrden = " ORDER BY fecha, desde, codigo"
            ElseIf pOrden = "Incidencia" Then
                mOrden = " ORDER BY cod_incidencia,ape1,ape2, nombre, fecha, desde"
            ElseIf pOrden = "Desde" Then
                mOrden = " ORDER BY desde "
            ElseIf pOrden = "Hasta" Then
                mOrden = " ORDER BY hasta "
            End If

            mCommand.CommandText = mConsulta & mOrden
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta & mOrden)
        End Try
    End Function

    Public Function Lista_Solicitudes_Grupos_Todos(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object Implements PresenciaDAO.Lista_Solicitudes_Grupos_Todos
        Dim mConsulta As String
        Dim mOrden As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, solicitud.DNI , solicitud.DNI || ' ' ||  ape1 || ' ' || ape2 || ' ' || nombre Nombre, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias, empleados"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND SOLICITUD.dni = EMPLEADOS.dni"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND SOLICITUD.dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND SOLICITUD.dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            'If pID_Lista_Responsables <> "" Then
            'mConsulta = mConsulta & " AND exists (select * from siguientes_solicitud where cod_solicitud = solicitud.codigo and dni in ('" & pID_Lista_Responsables & "'))  "
            'End If
            'If pLista_Ultimo_Responsable <> "" Then
            'mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            'End If
            If pCodigoIncidencia >= 0 Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia =  " & pCodigoIncidencia
            End If
            If Not Cambio_Grupo Is Nothing Then
                mConsulta = mConsulta & " AND cambio_grupo =  '" & Cambio_Grupo & "'"
            End If
            If pOrden = "Usuario" Then
                mOrden = " ORDER BY ape1,ape2, nombre, fecha, desde, codigo"
            ElseIf pOrden = "Fecha" Then
                mOrden = " ORDER BY fecha, desde, codigo"
            ElseIf pOrden = "Incidencia" Then
                mOrden = " ORDER BY cod_incidencia,ape1,ape2, nombre, fecha, desde"
            ElseIf pOrden = "Desde" Then
                mOrden = " ORDER BY desde "
            ElseIf pOrden = "Hasta" Then
                mOrden = " ORDER BY hasta "
            End If

            mCommand.CommandText = mConsulta & mOrden
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_Grupos_Todos", ex, mConsulta & mOrden)
        End Try
    End Function

    Public Function Lista_Solicitudes_Grupos_Cuadrantes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Lista_Responsables As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As String = "", Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object Implements PresenciaDAO.Lista_Solicitudes_Grupos_Cuadrantes
        Dim mConsulta As String
        Dim mOrden As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT codigo, fecha, estado, solicitud.DNI , solicitud.DNI || ' ' ||  ape1 || ' ' || ape2 || ' ' || nombre Nombre, "
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO "
            mConsulta = mConsulta & " FROM solicitud, incidencias, empleados"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND SOLICITUD.dni = EMPLEADOS.dni"
            If pCodigo > 0 Then
                mConsulta = mConsulta & " AND codigo =  " & pCodigo
            End If
            If pDNI <> "" Then
                If Left(pDNI, 1) = "(" Then
                    mConsulta = mConsulta & " AND SOLICITUD.dni in " & pDNI
                Else
                    mConsulta = mConsulta & " AND SOLICITUD.dni = '" & pDNI & "'"
                End If
            End If
            If pListaEstados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pListaEstados & ")"
            End If
            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If
            If pID_Lista_Responsables <> "" Then
                mConsulta = mConsulta & " AND exists (select * from siguientes_solicitud where cod_solicitud = solicitud.codigo and dni in ('" & pID_Lista_Responsables & "'))  "
            End If
            'If pLista_Ultimo_Responsable <> "" Then
            'mConsulta = mConsulta & " AND ULTIMO_RESPONSABLE IN ('" & pLista_Ultimo_Responsable & "')"
            'End If
            If pCodigoIncidencia <> "" Then
                mConsulta = mConsulta & " AND incidencias.cod_incidencia in (" & pCodigoIncidencia & ")"
            End If
            'If Not Cambio_Grupo Is Nothing Then
            'mConsulta = mConsulta & " AND cambio_grupo =  '" & Cambio_Grupo & "'"
            'End If
            If pOrden = "Usuario" Then
                mOrden = " ORDER BY ape1,ape2, nombre, fecha, desde, codigo"
            ElseIf pOrden = "Fecha" Then
                mOrden = " ORDER BY fecha, desde, codigo"
            ElseIf pOrden = "Incidencia" Then
                mOrden = " ORDER BY cod_incidencia,ape1,ape2, nombre, fecha, desde"
            ElseIf pOrden = "Desde" Then
                mOrden = " ORDER BY desde "
            ElseIf pOrden = "Hasta" Then
                mOrden = " ORDER BY hasta "
            End If

            mCommand.CommandText = mConsulta & mOrden
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes", ex, mConsulta & mOrden)
        End Try
    End Function


    Public Function Lista_Grupos_Consulta_Autorizados_a_Justificar(ByVal pID_Responsable As String) As Object Implements PresenciaDAO.Lista_Grupos_Consulta_Autorizados_a_Justificar
        Dim mSQL As String
        Dim mSQL1 As String
        Dim mWhere As String
        Try


            mSQL1 = mSQL1 & " select to_char(grupos) "
            mSQL1 = mSQL1 & " from autorizadojustificar "
            mSQL1 = mSQL1 & " where dni= '" & pID_Responsable & "'"
            mSQL1 = mSQL1 & " and fecha_desde <= sysdate"
            mSQL1 = mSQL1 & " and (fecha_hasta >= sysdate or fecha_hasta is null)"

            mSQL = mSQL & " select * from gruposconsulta"
            mSQL = mSQL & " where cod_grupo in (" & mSQL1 & " )"
            mSQL = mSQL & " and ( grupo_padre not in (" & mSQL1 & " )"
            mSQL = mSQL & " or grupo_padre is null )"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            Me.Haz_Log("Ejecutando SQL:" & mSQL, 3)
            mReader = mCommand.ExecuteReader
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta_Autorizados_a_Justificar", ex, mSQL)
        End Try
    End Function

    Public Function Lista_Grupos_Consulta_Autorizados_a_Consultar(ByVal pID_Responsable As String) As Object Implements PresenciaDAO.Lista_Grupos_Consulta_Autorizados_a_Consultar
        Dim mSQL As String
        Dim mSQL1 As String
        Dim mWhere As String
        Try


            mSQL1 = mSQL1 & " select to_char(grupo) "
            mSQL1 = mSQL1 & " from autorizadoconsultar "
            mSQL1 = mSQL1 & " where ID_RESPONSABLE= '" & pID_Responsable & "'"

            mSQL = mSQL & " select * from gruposconsulta"
            mSQL = mSQL & " where cod_grupo in (" & mSQL1 & " )"
            mSQL = mSQL & " and ( grupo_padre not in (" & mSQL1 & " )"
            mSQL = mSQL & " or grupo_padre is null )"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            Me.Haz_Log("Ejecutando SQL:" & mSQL, 3)
            mReader = mCommand.ExecuteReader
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta_Autorizados_a_Consultar", ex, mSQL)
        End Try
    End Function

    Public Function Lista_Grupos_Consulta_Autorizados_a_Autorizar(ByVal pID_Responsable As String) As Object Implements PresenciaDAO.Lista_Grupos_Consulta_Autorizados_a_Autorizar
        Dim mSQL As String
        Dim mSQL1 As String
        Dim mWhere As String
        Try


            mSQL1 = mSQL1 & " select SUBSTR(to_char(id_usuario),4) "
            mSQL1 = mSQL1 & " from asignacion_responsable "
            mSQL1 = mSQL1 & " where ID_RESPONSABLE= '" & pID_Responsable & "'"

            mSQL = mSQL & " select * from gruposconsulta"
            mSQL = mSQL & " where cod_grupo in (" & mSQL1 & " )"
            mSQL = mSQL & " and ( grupo_padre not in (" & mSQL1 & " )"
            mSQL = mSQL & " or grupo_padre is null )"

            Dim mReader As Object
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            Me.Haz_Log("Ejecutando SQL:" & mSQL, 3)
            mReader = mCommand.ExecuteReader
            Return mReader
        Catch ex As Exception
            Trata_Error("Error en Lista_Grupos_Consulta_Autorizados_a_Consultar", ex, mSQL)
        End Try
    End Function

    Public Function Lista_Acumuladores_Todos_Dataset(ByRef pDatos As System.Data.DataSet, Optional ByVal pFavoritos As Boolean = False) As Boolean Implements PresenciaDAO.Lista_Acumuladores_Todos_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * FROM ACUMULADORES_TODOS "
            If pFavoritos Then
                mConsulta &= " where not favoritos is null and favoritos <> 0 order by favoritos"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Acumuladores_Todos_Dataset", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Numero_Justificaciones_Anio(ByVal pDNI As String, ByVal pAnio As String, ByVal pCod_Incidencia As String) As Integer Implements PresenciaDAO.Numero_Justificaciones_Anio
        Dim mConsulta As String
        Dim mWhere As String
        Dim mDatos As New DataSet

        Try

            mConsulta = "select nvl(count (Numero), 0) Dias from ( "
            mConsulta &= " select fecha_justificada Numero from justificaciones_t justificaciones "
            mConsulta &= " where(cod_incidencia =" & pCod_Incidencia & ") "
            mConsulta &= " and dni_empl = '" & pDNI & "' "
            mConsulta &= " and fecha_justificada >= '01/01/" & pAnio & "' "
            mConsulta &= " and fecha_justificada <= '31/12/" & pAnio & "' "
            mConsulta &= " group by fecha_justificada) "




            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(mDatos)

            If mDatos.Tables(0).Rows.Count > 0 Then
                Return mDatos.Tables(0).Rows(0)("Dias")
            Else
                Return 0
            End If
        Catch ex As Exception
            Trata_Error("Error en Numero_Justificaciones_Anio", ex, mConsulta)
            Return 0
        End Try

    End Function

    Public Function EjecutaComandoSQL(ByVal pSQL As String, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.EjecutaComandoSQL
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = pSQL
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en EjecutaComandoSQL", ex, pSQL)
            pError = ex.Message
            Return False
        End Try
    End Function



    Public Function Lista_Incidencias_TVR_Dataset(ByRef pDatos As System.Data.DataSet) As Boolean Implements PresenciaDAO.Lista_Incidencias_TVR_Dataset
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "Select * from incidencias where seleccionable_tvr = 'S' or seleccionable_tvr is null order by cod_incidencia"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Incidencias_TVR_Dataset", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Busqueda_Empleados_De_ResponsablesAut(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Busqueda_Empleados_De_ResponsablesAut
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'abro la conexion para el control


            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, dni dni_empl from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then

                'porque si son Todos, no tiene que seleccionar por grupos
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If


                'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                mWhere = mWhere & " dni in (select dni from empleados where " _
                                   & " dni in (select id_usuario from asignacion_responsable " _
                                   & " where id_usuario not like 'GRP%' and (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "'))) " _
                                   & " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                                   & " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                                   & "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"

            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Busqueda_Empleados_De_ResponsablesAut_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Busqueda_Empleados_De_ResponsablesAut_Dataset
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'abro la conexion para el control


            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre || ' ' || Ape1 || ' ' || Ape2 NOMBRE from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then

                'porque si son Todos, no tiene que seleccionar por grupos
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If


                'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                mWhere = mWhere & " dni in (select dni from empleados where " _
                                   & " dni in (select id_usuario from asignacion_responsable " _
                                   & " where id_usuario not like 'GRP%' and (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "'))) " _
                                   & " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                                   & " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                                   & "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"

            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Busqueda_Empleados_De_ResponsablesJust_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Busqueda_Empleados_De_ResponsablesJust_Dataset
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'abro la conexion para el control


            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre || ' ' || Ape1 || ' ' || Ape2 NOMBRE from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then

                'porque si son Todos, no tiene que seleccionar por grupos
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If


                'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                mWhere = mWhere & " dni in (select dni_empl from pertenecena where " _
                                  & " cod_grupo in (select grupos from autorizadojustificar " _
                                  & " where dni = '" & pResponsable & "' ))"

            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Busqueda_Empleados_De_ResponsablesConsultar(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Busqueda_Empleados_De_ResponsablesConsultar
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'abro la conexion para el control


            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre, Ape1, Ape2, Email, Clave_web, centro, cargo, clave_emp, telefono, calcula_saldo, dni dni_empl from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then

                'porque si son Todos, no tiene que seleccionar por grupos
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If


                'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                mWhere = mWhere & " dni in (select dni_empl from pertenecena where " _
                                   & " cod_grupo in (select grupo from autorizadoconsultar " _
                                   & " where id_responsable = '" & pResponsable & "' ))" _
                                   '& " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                '& " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                '& "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"

            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Busqueda_Empleados_De_ResponsablesConsultar_DataSet(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean Implements PresenciaDAO.Busqueda_Empleados_De_ResponsablesConsultar_Dataset
        'devuelve la tabla de empleados
        Dim mConsulta As String
        Dim mWhere As String

        Try
            'abro la conexion para el control


            'pone los datos del dia seleccionado
            mConsulta = "SELECT DNI, Nombre || ' ' || Ape1 || ' ' || Ape2 NOMBRE from empleados"

            If pDNI <> "" Then
                mWhere = mWhere & " WHERE UPPER(DNI) LIKE '" & UCase(pDNI) & "%'"
            End If

            If pResponsable <> "" Then

                'porque si son Todos, no tiene que seleccionar por grupos
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If


                'con esta consulta, estariamos cogiendo las personas que pertenecen a todos los grupos de la lista, cuando nos interesa que esté en cualquiera de ellos
                'mWhere = mWhere & " dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni having count(distinct cod_grupo) = " & contador & ")"
                mWhere = mWhere & " dni in (select dni_empl from pertenecena where " _
                                   & " cod_grupo in (select grupo from autorizadoconsultar " _
                                   & " where id_responsable = '" & pResponsable & "' ))"
                '& " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                '& " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                '& "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"

            End If

            If pNombre <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(NOMBRE) LIKE '" & UCase(pNombre) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(NOMBRE)) LIKE SUPR_ACCENT(UPPER('" & pNombre & "%'))"
            End If
            If pApe1 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE1) LIKE '" & UCase(pApe1) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE1)) LIKE SUPR_ACCENT(UPPER('" & pApe1 & "%'))"
            End If
            If pApe2 <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(APE2) LIKE '" & UCase(pApe2) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(APE2)) LIKE SUPR_ACCENT(UPPER('" & pApe2 & "%'))"
            End If
            If pClave_Empleado <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " UPPER(CLAVE_EMP) LIKE '" & UCase(pClave_Empleado) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(CLAVE_EMP)) LIKE SUPR_ACCENT(UPPER('" & pClave_Empleado & "%'))"
            End If
            If pApellidos <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                'mWhere = mWhere & " upper(ape1 || ' ' || ape2) LIKE '" & UCase(pApellidos) & "%'"
                mWhere = mWhere & " SUPR_ACCENT(UPPER(ape1 || ' ' || ape2)) LIKE SUPR_ACCENT(UPPER('" & pApellidos & "%'))"
            End If

            If mWhere <> "" Then
                mConsulta = mConsulta & " " & mWhere
            End If

            mConsulta = mConsulta & " ORDER BY Ape1,Ape2,Nombre"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Busca_Empleados", ex, mConsulta)
            Return False
        End Try
    End Function



    Public Function Consulta_Datos_Incidencias_Aut(ByRef pDatos As System.Data.DataSet, ByVal pResponsable As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Incidencias_Aut
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni_empl, fecha_justificada, cod_incidencia "
            mConsulta &= " from justificaciones_t j, empleados e where   e.DNI = j.DNI_EMPL "
            mConsulta &= " and e.dni in (select dni from empleados where " _
                                   & " dni in (select id_usuario from asignacion_responsable " _
                                   & " where id_usuario not like 'GRP%' and (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "'))) " _
                                   & " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                                   & " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                                   & "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"

            mConsulta &= " and j.FECHA_JUSTIFICADA >= '" & Fecha_Ini & "' and j.FECHA_JUSTIFICADA <= '" & Fecha_Fin & "'"
            mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            mConsulta &= " group by j.dni_empl, fecha_justificada , cod_incidencia"
            mConsulta &= " order by j.dni_empl, fecha_justificada, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Incidencias_Aut", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Consulta_Datos_Solicitudes_Aut(ByRef pDatos As System.Data.DataSet, ByVal pResponsable As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean Implements PresenciaDAO.Consulta_Datos_Solicitudes_Aut
        Dim mConsulta As String
        Dim mWhere As String

        Try

            mConsulta = "select j.dni, fecha, cod_incidencia "
            mConsulta &= " from solicitud j, empleados e where   e.DNI = j.DNI "
            mConsulta &= " and e.dni in (select dni from empleados where " _
                                   & " dni in (select id_usuario from asignacion_responsable " _
                                   & " where id_usuario not like 'GRP%' and (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "'))) " _
                                   & " or dni in (select dni_empl from pertenecena where tipo_grp = 1 " _
                                   & " and cod_grupo in (select to_number(substr(id_usuario, 4, 4)) " _
                                   & "	from asignacion_responsable where (id_responsable = '" & pResponsable & "' or id_responsable in (select id_responsable from delegados where id_delegado = '" & pResponsable & "')) and id_usuario like 'GRP%') ))"
            mConsulta &= " and j.FECHA >= '" & Fecha_Ini & "' and j.FECHA <= '" & Fecha_Fin & "'"
            mConsulta &= " and cod_incidencia in (" & pCod_Inc & ")"
            mConsulta &= " group by j.dni, fecha, cod_incidencia"
            mConsulta &= " order by j.dni, fecha, cod_incidencia"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Consulta_Dni_Solicitudes_Aut", ex, mConsulta)
            Return False
        End Try
    End Function


    Public Overloads Function Lista_Solicitudes_de_Grupos2(ByRef pDatos As Object, ByVal pListaGrupos As String, ByVal pDni As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pLista_Estados As String, ByVal pIncidencia As String, Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Solicitudes_de_Grupos2
        Dim mConsulta As String
        Dim mWhere As String
        Dim contador As Integer
        Try

            Dim mGruposAutorizados As String

            'Dim mBD As PresenciaDAO
            'mBD = DAOFactory.GetFactory(CTE_Tipo_BD).getPresenciaDAO(CTE_Cadena_conexion)
            'mBD.Conecta()

            contador = NumeroDeGrupos(pListaGrupos)

            mConsulta &= "SELECT codigo, fecha, estado, solicitud.DNI, empleados.ape1 || ' ' || empleados.ape2 || ', ' || empleados.nombre as Nombre,"
            mConsulta = mConsulta & " solicitud.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, desde, hasta, observaciones, "
            mConsulta = mConsulta & " fecha_sol, id_siguiente_responsable, "
            mConsulta = mConsulta & " DESDE_ORIGINAL, HASTA_ORIGINAL, OBSERVACIONES_ORIGINAL, INCIDENCIA_ORIGINAL, "
            mConsulta = mConsulta & " COD_JUSTIFICACION, ULTIMO_RESPONSABLE, solicitud.tipo, solicitud.CAMBIO_GRUPO"
            mConsulta = mConsulta & " FROM solicitud, incidencias, empleados"
            mConsulta = mConsulta & " WHERE solicitud.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND   solicitud.dni = empleados.dni(+)"
            If pDni = "" Then
                If pListaGrupos <> "" And pListaGrupos <> "Todos" Then
                    mConsulta = mConsulta & " AND solicitud.dni IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni)"
                End If
            Else
                mConsulta = mConsulta & " AND solicitud.dni = '" & pDni & "'"
            End If


            If pLista_Estados <> "" Then
                mConsulta = mConsulta & " AND estado in (" & pLista_Estados & ")"
            End If
            If pIncidencia <> "" Then
                mConsulta = mConsulta & " AND solicitud.cod_incidencia = " & pIncidencia
            End If

            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If

            If pOrden <> "" Then
                mConsulta = mConsulta & " order by " & pOrden & ",fecha, desde, solicitud.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha, desde, incidencias.desc_incidencia"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Solicitudes_de_Grupos2", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Justificaciones_de_Grupos2(ByRef pDatos As Object, ByVal pListaGrupos As String, ByVal pdni As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pLista_Estados As String, ByVal pIncidencia As String, Optional ByVal pOrden As String = "") As Boolean Implements PresenciaDAO.Lista_Justificaciones_de_Grupos2
        Dim mConsulta As String
        Dim mWhere As String
        Dim contador As Integer
        Try

            Dim mGruposAutorizados As String

            'Dim mBD As PresenciaDAO
            'mBD = DAOFactory.GetFactory(CTE_Tipo_BD).getPresenciaDAO(CTE_Cadena_conexion)
            'mBD.Conecta()

            contador = NumeroDeGrupos(pListaGrupos)

            mConsulta &= "SELECT cod_justificacion as codigo, fecha_justificada as fecha, 'R' as estado, justificaciones.dni_empl, empleados.ape1 || ' ' || empleados.ape2 || ', ' || empleados.nombre as Nombre,"
            mConsulta = mConsulta & " justificaciones.cod_incidencia, nvl(incidencias.desc_incidencia,'') desc_Incidencia, formatea_hora(desde_minutos) as desde , formatea_hora(hasta_minutos) as hasta, observaciones, "
            mConsulta = mConsulta & " 'N' as CAMBIO_GRUPO"
            mConsulta = mConsulta & " FROM justificaciones, incidencias, empleados"
            mConsulta = mConsulta & " WHERE justificaciones.cod_incidencia = incidencias.cod_incidencia(+)"
            mConsulta = mConsulta & " AND   justificaciones.dni_empl = empleados.dni(+)"
            If pdni = "" Then
                If pListaGrupos <> "" And pListaGrupos <> "Todos" Then
                    mConsulta = mConsulta & " AND justificaciones.dni_empl IN (SELECT dni FROM perteneceagrupo WHERE COD_GRUPO IN  ( " & pListaGrupos & " ) group by dni)"
                End If
            Else
                mConsulta = mConsulta & " AND justificaciones.dni_empl = '" & pdni & "'"
            End If


            'If pLista_Estados <> "" Then
            '    mConsulta = mConsulta & " AND estado in (" & pLista_Estados & ")"
            'End If
            If pIncidencia <> "" Then
                mConsulta = mConsulta & " AND justificaciones.cod_incidencia = " & pIncidencia
            End If

            If pFechaDesde <> "" Then
                mConsulta = mConsulta & " AND fecha_justificada >= to_date('" & pFechaDesde & "','DD/MM/YYYY')"
            End If
            If pFechaHasta <> "" Then
                mConsulta = mConsulta & " AND fecha_justificada <= to_date('" & pFechaHasta & "','DD/MM/YYYY')"
            End If

            If pOrden <> "" Then
                mConsulta = mConsulta & " order by " & pOrden & ",fecha, desde, justificaciones.cod_incidencia"
            Else
                mConsulta = mConsulta & " order by fecha_hora, desde_minutos, incidencias.desc_incidencia"
            End If

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_justificacioneses_de_Grupos2", ex, mConsulta)
        End Try
    End Function

    Public Overloads Function Elimina_Justificacion1(ByVal pListaCodigos As String) As Boolean Implements PresenciaDAO.Elimina_Justificacion1
        Dim mConsulta As String
        Try
            Dim mWhere As String
            mConsulta = "DELETE JUSTIFICACIONES WHERE "
            mWhere = " COD_JUSTIFICACION in (" & pListaCodigos & ")"
            mConsulta = mConsulta & mWhere
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True


        Catch ex As Exception
            Trata_Error("Error en Elimina_Justificacion", ex, mConsulta)
            Return False
        End Try
    End Function



    Public Function Lista_Eventos_Visor(ByRef pDatos As Object, Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pHoraDesde As String = "", Optional ByVal pHoraHasta As String = "", Optional ByVal pListaUsuarios As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal pSinHerencia As Boolean = True, Optional ByVal pTipoEvento As String = "", Optional ByVal pAgruparEvento As Boolean = False, Optional ByVal pPermitido As String = "", Optional ByVal pListaCodRecurso As String = "", Optional ByVal pListaGrupoRecurso As String = "", Optional ByVal pOrden As String = "") As Object Implements PresenciaDAO.Lista_Eventos_Visor
        Dim mConsulta As String
        Dim mCodGrupo As String
        Dim mDescGrupo As String
        Dim mGrupo As String
        Dim mGruposHijo As String
        Dim mGruposHijo2 As String
        Dim mWhere As String
        Dim mTipo As String

        Try

            mConsulta &= " SELECT "
            If Not pAgruparEvento Then
                mConsulta &= " cod_evento Evento, "
            End If

            If Not pAgruparEvento Then
                mConsulta &= " to_char(FECHA,'DD/MM/YYYY') FECHA,"
                mConsulta &= "  HORA, E_S, "
            Else
                mConsulta &= " distinct (to_char(FECHA,'DD/MM/YYYY')) FECHA,"
                mConsulta &= " '' Evento, "
                mConsulta &= " '' HORA, '' E_S,"
            End If
            mConsulta &= " dni_empl dni,dni_empl dni_justif,  "
            mConsulta &= " ape1|| ' ' || ape2 || ', ' || nombre empleado, "
            If Not pAgruparEvento Then
                mConsulta &= " eventos.pan_tarjeta, eventos.cod_recurso cod, recursos.desc_recurso Recurso, PERMITIDO, '' Picadas,"
            Else
                mConsulta &= " '' pan_tarjeta, '' cod,'' Recurso, '' permitido, picadas_usuario(dni_empl,fecha) Picadas,"
            End If
            'mConsulta &= " empleados.telefono, email, centro, cargo, empresa"
            mConsulta &= " empresa "

            If pListaUsuarios <> "" Then
                'If pAgruparEvento = False Then
                If InStr(pListaUsuarios, "'", CompareMethod.Text) = 0 Then
                    pListaUsuarios = "'" & pListaUsuarios & "'"
                End If
                mTipo = "1"
                mCodGrupo &= " (select "
                mCodGrupo &= " max(b.cod_grupo) cod_grupo " '--, max((select a.desc_grupo from gruposconsulta a where a.cod_grupo = b.cod_grupo)) desc_grupo "
                mGruposHijo &= " from pertenecena , "
                mGruposHijo &= " (select cod_grupo , desc_grupo, grupo_padre, level nivel  "
                mGruposHijo &= " from gruposconsulta start with cod_grupo in (" & pListaGrupos & ")"
                mGruposHijo &= " connect by prior cod_grupo = grupo_padre ) b"
                mGruposHijo &= " where pertenecena.cod_grupo = b.cod_grupo"
                mGruposHijo &= " and dni_empl = eventos.dni_empl"
                mGruposHijo2 &= " and nivel = ("
                mGruposHijo2 &= " select max(nivel) " '--, desc_grupo --desc_grupo,  nivel "
                mGruposHijo2 &= mGruposHijo
                mGruposHijo2 &= ")"
                mGruposHijo = mGruposHijo & mGruposHijo2
                mCodGrupo &= mGruposHijo
                mCodGrupo &= " ) as cod"



                mDescGrupo &= " (select "
                mDescGrupo &= " max((select a.desc_grupo from gruposconsulta a where a.cod_grupo = b.cod_grupo)) desc_grupo "
                mDescGrupo &= mGruposHijo
                mDescGrupo &= ") as grupo"
                'End If

            Else
                If pListaGrupos <> "" Then
                    If pSinHerencia Then
                        mTipo = "2"
                        'mCodGrupo &= " (select cod_grupo from "
                        mCodGrupo &= " max((select "
                        mCodGrupo &= " pertenecena.cod_grupo cod_grupo " ', max((select a.desc_grupo from gruposconsulta a where a.cod_grupo = pertenecena.cod_grupo)) desc_grupo "
                        mGruposHijo &= " from pertenecena , "
                        mGruposHijo &= "( select cod_grupo , desc_grupo from gruposconsulta where cod_grupo in (" & pListaGrupos & ")"
                        mGruposHijo &= " ) b where pertenecena.cod_grupo = b.cod_grupo and pertenecena.tipo_grp=2"
                        mGruposHijo &= " and dni_empl = eventos.dni_empl"
                        mCodGrupo &= mGruposHijo
                        'mCodGrupo &= " )) cod_grupo"
                        mCodGrupo &= " )) as cod"

                        'mDescGrupo &= "(select desc_grupo from ("
                        mDescGrupo &= " max((select"
                        mDescGrupo &= " (select a.desc_grupo from gruposconsulta a where a.cod_grupo = pertenecena.cod_grupo and pertenecena.tipo_grp=2) desc_grupo "
                        mDescGrupo &= mGruposHijo
                        'mDescGrupo &= ")) desc_grupo"
                        mDescGrupo &= ")) as grupo"

                    Else
                        mTipo = "2"
                        'mCodGrupo &= " (select cod_grupo from ("
                        mCodGrupo &= " max((select "
                        mCodGrupo &= " b.cod_grupo cod_grupo " '--, max((select a.desc_grupo from gruposconsulta a where a.cod_grupo = b.cod_grupo)) desc_grupo "
                        mGruposHijo &= " from pertenecena , "
                        mGruposHijo &= " (select cod_grupo , desc_grupo, grupo_padre, level nivel  "
                        mGruposHijo &= " from gruposconsulta start with cod_grupo in (" & pListaGrupos & ")"
                        mGruposHijo &= " connect by prior cod_grupo = grupo_padre ) b"
                        mGruposHijo &= " where pertenecena.cod_grupo = b.cod_grupo and pertenecena.tipo_grp=2"
                        mGruposHijo &= " and dni_empl = eventos.dni_empl"
                        mGruposHijo2 &= " and nivel = ("
                        mGruposHijo2 &= " select max(nivel) " '--, desc_grupo --desc_grupo,  nivel "
                        mGruposHijo2 &= mGruposHijo
                        mGruposHijo2 &= ")"
                        mGruposHijo = mGruposHijo & mGruposHijo2
                        mCodGrupo &= mGruposHijo
                        mCodGrupo &= " )) as cod"



                        mDescGrupo &= "max((select "
                        mDescGrupo &= " (select a.desc_grupo from gruposconsulta a where a.cod_grupo = b.cod_grupo) desc_grupo "
                        mDescGrupo &= mGruposHijo
                        mDescGrupo &= ")) as grupo"
                    End If
                End If
            End If
            If mCodGrupo <> "" And mDescGrupo <> "" Then
                mConsulta &= ", " & mCodGrupo & ", " & mDescGrupo
            Else
                mConsulta &= ",'' cod, '' Grupo"
            End If
            mConsulta &= " FROM eventos, empleados, recursos "
            mConsulta &= " where empleados.dni = eventos.dni_empl "
            mConsulta &= " and recursos.cod_recurso = eventos.cod_recurso"
            If pListaUsuarios <> "" Then
                If InStr(pListaUsuarios, "'", CompareMethod.Text) = 0 Then
                    pListaUsuarios = "'" & pListaUsuarios & "'"
                End If
                pListaUsuarios = " and ( empleados.dni in (" & pListaUsuarios & " ) "
            Else
                If pListaGrupos <> "" Then
                    If pSinHerencia Then
                        If pListaUsuarios = "" Then
                            pListaUsuarios &= " and ( empleados.dni in (select pertenecena.dni_empl from pertenecena where cod_grupo in (" & pListaGrupos & "))"
                        Else
                            pListaUsuarios &= " or ( empleados.dni in (select pertenecena.dni_empl from pertenecena where cod_grupo in (" & pListaGrupos & ")))"
                        End If
                    Else
                        If pListaUsuarios = "" Then
                            pListaUsuarios &= " and ( empleados.dni in (select pertenecena.dni_empl from pertenecena where cod_grupo in "

                            pListaUsuarios &= " ("
                            pListaUsuarios &= " select cod_grupo "
                            pListaUsuarios &= " from gruposconsulta"
                            pListaUsuarios &= " start with cod_grupo in (" & pListaGrupos & ")"
                            pListaUsuarios &= " connect by prior cod_grupo = grupo_padre"
                            pListaUsuarios &= " ))"
                        Else
                            pListaUsuarios &= " or ( empleados.dni in (select dni from pertenecena where cod_grupo in "

                            pListaUsuarios &= " ("
                            pListaUsuarios &= " select cod_grupo "
                            pListaUsuarios &= " from gruposconsulta"
                            pListaUsuarios &= " start with cod_grupo in (" & pListaGrupos & ")"
                            pListaUsuarios &= " connect by prior cod_grupo = grupo_padre"
                            pListaUsuarios &= " ))"
                        End If
                    End If
                End If
            End If
            If pListaUsuarios <> "" Then
                pListaUsuarios &= ")"
                mConsulta &= pListaUsuarios
            End If

            If pListaCodRecurso <> "" Then
                mConsulta &= " and eventos.cod_recurso in (" & pListaCodRecurso & ")"
            End If
            If pListaCodRecurso = "" And pListaGrupoRecurso <> "" Then
                If pListaGrupoRecurso = "null" Then
                    mConsulta &= " and eventos.cod_recurso in ( select cod_recurso from recursos where cod_gruporecursos is " & pListaGrupoRecurso & ")"
                Else
                    mConsulta &= " and eventos.cod_recurso in ( select cod_recurso from recursos where cod_gruporecursos = " & pListaGrupoRecurso & ")"
                End If
            End If
            If pPermitido <> "" Then
                mConsulta &= " and permitido = '" & pPermitido & "'"
            End If
            If pFechaDesde <> "" Then
                mConsulta &= " and fecha >= '" & pFechaDesde & "'"
            End If
            If pFechaHasta <> "" Then
                mConsulta &= " and fecha <= '" & pFechaHasta & "'"
            End If
            If pHoraDesde <> "" Then
                mConsulta &= " and hora >= '" & pHoraDesde & "'"
            End If
            If pHoraHasta <> "" Then
                mConsulta &= " and hora <= '" & pHoraHasta & "'"
            End If

            If pTipoEvento <> "" Then
                mConsulta &= " and e_s = '" & pTipoEvento & "'"
            End If

            If pAgruparEvento Then
                If pListaUsuarios <> "" Then

                    If pListaGrupos <> "" And mTipo = "2" Then
                        mConsulta &= " group by FECHA, dni_empl, ape1|| ' ' || ape2 || ', ' || nombre, empresa"
                    ElseIf pListaGrupos <> "" And mTipo = "" Then
                        mConsulta &= " group by FECHA, dni_empl, ape1|| ' ' || ape2 || ', ' || nombre, empresa"
                    End If
                Else

                    If pListaGrupos <> "" And mTipo = "2" Then
                        mConsulta &= " group by FECHA, dni_empl, ape1|| ' ' || ape2 || ', ' || nombre, empresa"
                    End If
                End If
            Else
                If pListaUsuarios <> "" Then
                    If pListaGrupos <> "" And mTipo = "2" Then
                        mConsulta &= " group by cod_evento,fecha,hora,e_s,dni_empl,ape1|| ' ' || ape2 || ', ' || nombre,eventos.pan_tarjeta, eventos.cod_recurso , recursos.desc_recurso, PERMITIDO, empresa"
                    ElseIf pListaGrupos <> "" And mTipo = "" Then
                        mConsulta &= " group by FECHA, dni_empl, ape1|| ' ' || ape2 || ', ' || nombre, empresa"
                    End If
                Else
                    If pListaGrupos <> "" And mTipo = "2" Then
                        mConsulta &= " group by cod_evento,fecha,hora,e_s,dni_empl,ape1|| ' ' || ape2 || ', ' || nombre,eventos.pan_tarjeta, eventos.cod_recurso , recursos.desc_recurso, PERMITIDO, empresa"
                    End If
                End If

            End If
            If pAgruparEvento Then
                If pOrden <> "" Then
                    mConsulta = mConsulta & " order by " & pOrden
                Else
                    mConsulta = mConsulta & " order by fecha, hora, ape1|| ' ' || ape2 || ', ' || nombre"
                End If
            Else
                mConsulta = mConsulta & " order by fecha, hora, ape1|| ' ' || ape2 || ', ' || nombre"
            End If


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            mDataAdapter.Fill(pDatos)
            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Eventos_Visor", ex, mConsulta)
        End Try
    End Function

    Public Function TratamientoIncidencia(ByVal pCodInci As Integer, _
                                  ByVal pDni As String, _
                                  ByVal pFecha As String, _
                                  ByVal pFechaDesde As String, _
                                  ByVal pFechaHasta As String, _
                                  ByVal pHoraDesde As String, _
                                  ByVal pHoraHasta As String, _
                                  ByVal pTipoIntervalo As String, _
                                  ByVal pTipoOperacion As String, _
                                  ByRef pMaximo As Integer, _
                                  ByRef pMinimo As Integer, _
                                  ByRef pTipoLimite As String, _
                                  ByRef pPermitido As Integer, _
                                  ByRef pDescDenegacion As String) As Integer Implements PresenciaDAO.TratamientoIncidencia

        Dim mConsulta As String
        '--  pTipoIntervalo IN varchar2, (L.laboral,N:Natural)
        '--  pTipoLimite  IN VARCHAR2) (D:Día/H:Horas)
        '--  pTipoOperacion (S:Solicitud/J:Justificación/L:Límites máximo y mínimo)

        Try

            '// build command
            Dim mCommand As New OleDb.OleDbCommand("TratamientoIncidencia")
            Dim myParameter As New OleDb.OleDbParameter

            mCommand.Connection = mConexion
            mCommand.CommandType = CommandType.StoredProcedure
            '// add parameters
            myParameter = mCommand.Parameters.Add("@pCodInci", OleDb.OleDbType.UnsignedInt)
            myParameter.Value = pCodInci
            myParameter = mCommand.Parameters.Add("@pDni", OleDb.OleDbType.VarChar)
            myParameter.Value = pDni
            myParameter = mCommand.Parameters.Add("@pFecha", OleDb.OleDbType.Date)
            myParameter.Value = pFecha
            myParameter = mCommand.Parameters.Add("@pFechaDesde", OleDb.OleDbType.Date)
            myParameter.Value = pFechaDesde
            myParameter = mCommand.Parameters.Add("@pFechaHasta", OleDb.OleDbType.Date)
            myParameter.Value = pFechaHasta
            myParameter = mCommand.Parameters.Add("@pHoraDesde", OleDb.OleDbType.VarChar)
            myParameter.Value = pHoraDesde
            myParameter = mCommand.Parameters.Add("@pHoraHasta", OleDb.OleDbType.VarChar)
            myParameter.Value = pHoraHasta
            myParameter = mCommand.Parameters.Add("@pTipoIntervalo", OleDb.OleDbType.VarChar)
            myParameter.Value = pTipoIntervalo
            myParameter = mCommand.Parameters.Add("@pTipoOperacion", OleDb.OleDbType.VarChar)
            myParameter.Value = pTipoOperacion

            mCommand.Parameters.Add("@pMaximo", OleDb.OleDbType.Integer)
            mCommand.Parameters.Add("@pMinimo", OleDb.OleDbType.Integer)
            mCommand.Parameters.Add("@pTipoLimite", OleDb.OleDbType.VarChar, 1)
            mCommand.Parameters.Add("@pPermitido", OleDb.OleDbType.VarChar, 1)
            mCommand.Parameters.Add("@pDescDenegacion", OleDb.OleDbType.VarChar, 200)

            mCommand.Parameters("@pMaximo").Direction = ParameterDirection.Output
            mCommand.Parameters("@pMinimo").Direction = ParameterDirection.Output
            mCommand.Parameters("@pTipoLimite").Direction = ParameterDirection.Output
            mCommand.Parameters("@pPermitido").Direction = ParameterDirection.Output
            mCommand.Parameters("@pDescDenegacion").Direction = ParameterDirection.Output

            mCommand.ExecuteNonQuery()

            pMaximo = NVL(mCommand.Parameters("@pMaximo").Value, 0)
            pMinimo = NVL(mCommand.Parameters("@pMinimo").Value, 0)
            pTipoLimite = NVL(mCommand.Parameters("@pTipoLimite").Value, 0)
            pPermitido = NVL(mCommand.Parameters("@pPermitido").Value, 0)
            pDescDenegacion = NVL(mCommand.Parameters("@pDescDenegacion").Value, "")

        Catch ex As Exception
            Trata_Error("Error en TratamientoIncidencia", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Cadena_de_Intervalos_Opcional(ByVal pcod_horario As Integer) As String Implements PresenciaDAO.Cadena_de_Intervalos_Opcional
        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String

        Try

            mCommand.Connection = mConexion
            mSQL = "SELECT inicio_intervalo, fin_intervalo FROM IntervaloOpcional "
            mSQL = mSQL & " WHERE cod_horario = " & pcod_horario
            mSQL = mSQL & " ORDER BY inicio_intervalo"
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            mLista = ""
            While mReader.Read
                If mLista <> "" Then
                    mLista = mLista & ";"
                End If
                mLista = mLista & Format(mReader(0), "0000") & "-" & Format(mReader(1), "0000")
            End While
            mReader.Close()

            Return mLista

        Catch ex As Exception
            Trata_Error("Error en Cadena_de_Intervalos_Recuperacion", ex, mSQL)
        End Try
    End Function

    Public Function Intervalo_minimo_incidencia(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean Implements PresenciaDAO.Intervalo_minimo_incidencia

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT Tiempo_Minimo from incidencias"
            mConsulta = mConsulta & " WHERE COD_INCIDENCIA = " & pCod_Incidencia

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Intervalo_minimo_Incidencia", ex, mConsulta)
        End Try

    End Function

    Public Function Intervalo_minimo_duracion(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean Implements PresenciaDAO.Intervalo_minimo_duracion

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT Minimo_Duracion from incidencias"
            mConsulta = mConsulta & " WHERE COD_INCIDENCIA = " & pCod_Incidencia

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Intervalo_minimo_duracion", ex, mConsulta)
        End Try

    End Function

    Public Function Intervalo_maximo_duracion(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean Implements PresenciaDAO.Intervalo_maximo_duracion

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT Maximo_Duracion from incidencias"
            mConsulta = mConsulta & " WHERE COD_INCIDENCIA = " & pCod_Incidencia

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Intervalo_maximo_duracion", ex, mConsulta)
        End Try

    End Function

    Public Function Intervalo_maximo_duracion_Incidencia_TipoContrato(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCodContrato As Integer = -1) As Boolean Implements PresenciaDAO.Intervalo_maximo_duracion_Incidencia_TipoContrato

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT Maximo_Duracion from tipocontrato_incidencia"
            mConsulta = mConsulta & " WHERE COD_TIPOCONTRATO=" & pCodContrato & " AND COD_INCIDENCIA = " & pCod_Incidencia

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Intervalo_maximo_duracion_Incidencia_TipoContrato", ex, mConsulta)
        End Try

    End Function


    Public Function Intervalo_minimo_duracion_Incidencia_TipoContrato(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCodContrato As Integer = -1) As Boolean Implements PresenciaDAO.Intervalo_minimo_duracion_Incidencia_TipoContrato

        Dim mConsulta As String
        Dim mWhere As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            'abro la conexion para el control
            mCommand.Connection = mConexion

            mConsulta = "SELECT Minimo_Duracion from tipocontrato_incidencia"
            mConsulta = mConsulta & " WHERE COD_TIPOCONTRATO=" & pCodContrato & " AND COD_INCIDENCIA = " & pCod_Incidencia

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Intervalo_minimo_duracion_Incidencia_TipoContrato", ex, mConsulta)
        End Try

    End Function

    Function IntervalosFactorCompensacion(ByVal pCodInci As Integer, _
                                ByVal pDni As String, _
                                ByVal pFecha As String, _
                                ByVal pFechaDesde As String, _
                                ByVal pFechaHasta As String, _
                                ByVal pHoraDesde As String, _
                                ByVal pHoraHasta As String, _
                                ByRef pIntervalos As String) As Integer Implements PresenciaDAO.IntervalosFactorCompensacion

        Dim mConsulta As String
        '--  pTipoIntervalo IN varchar2, (L.laboral,N:Natural)
        '--  pTipoLimite  IN VARCHAR2) (D:Día/H:Horas)
        '--  pTipoOperacion (S:Solicitud/J:Justificación/L:Límites máximo y mínimo)

        Try

            '// build command
            Dim mCommand As New OleDb.OleDbCommand("IntervalosFactorCompensacion")
            Dim myParameter As New OleDb.OleDbParameter

            mCommand.Connection = mConexion
            mCommand.CommandType = CommandType.StoredProcedure
            '// add parameters
            myParameter = mCommand.Parameters.Add("@pCodInci", OleDb.OleDbType.UnsignedInt)
            myParameter.Value = pCodInci
            myParameter = mCommand.Parameters.Add("@pDni", OleDb.OleDbType.VarChar)
            myParameter.Value = pDni
            myParameter = mCommand.Parameters.Add("@pFecha", OleDb.OleDbType.Date)
            myParameter.Value = pFecha
            myParameter = mCommand.Parameters.Add("@pFechaDesde", OleDb.OleDbType.Date)
            myParameter.Value = pFechaDesde
            myParameter = mCommand.Parameters.Add("@pFechaHasta", OleDb.OleDbType.Date)
            myParameter.Value = pFechaHasta
            myParameter = mCommand.Parameters.Add("@pHoraDesde", OleDb.OleDbType.VarChar)
            myParameter.Value = pHoraDesde
            myParameter = mCommand.Parameters.Add("@pHoraHasta", OleDb.OleDbType.VarChar)
            myParameter.Value = pHoraHasta
            myParameter = mCommand.Parameters.Add("@pTipoIntervalo", OleDb.OleDbType.VarChar)

            mCommand.Parameters.Add("@pIntervalos", OleDb.OleDbType.VarChar, 200)

            mCommand.Parameters("@pIntervalos").Direction = ParameterDirection.Output

            mCommand.ExecuteNonQuery()

            pIntervalos = NVL(mCommand.Parameters("@pIntervalos").Value, "")

        Catch ex As Exception
            Trata_Error("Error en IntervalosFactorCompensacion", ex, mConsulta)
            Return False
        End Try
    End Function

    Public Function Inserta_JustificacionACompensar(ByVal pCodJust As Long, ByVal pDuracion As Integer, ByVal pFactor As Double, ByVal pTipoFactor As String) As Long Implements PresenciaDAO.Inserta_JustificacionACompensar

        Dim mConsulta As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim auxcodigo As Integer


            auxcodigo = 1

            mConsulta = "INSERT INTO JUSTIFICACIONES_A_COMPENSAR(COD_JUSTIFICACION,DURACION,FACTOR_COMPENSACION,TIPO_FACTOR,ESTADO)"
            mConsulta = mConsulta & " VALUES(" & pCodJust & ", " & pDuracion & ", " & CStr(pFactor).Replace(",", ".") & ", '" & pTipoFactor & "','P')"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            auxcodigo = mCommand.ExecuteNonQuery()

            'mConsulta = "SELECT MAX(ID_Justificacion) FROM JUSTIFICACIONES_A_COMPENSAR WHERE Cod_Justificacion = " & pCodJust
            'mCommand.Connection = mConexion
            'mCommand.CommandText = mConsulta
            'mReader = mCommand.ExecuteReader
            'If mReader.Read Then
            '    auxcodigo = NVL(mReader(0), 0)
            'End If
            'mReader.Close()


            Return auxcodigo


        Catch ex As Exception
            Trata_Error("Error en Inserta_JustificacionACompensar", ex, mConsulta)
            Return -1
        End Try

    End Function

    Public Function Comprobar_Solicitud_Base(ByVal pCod_Solicitud As String) As Object Implements PresenciaDAO.Comprobar_Solicitud_Base
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT cod_solicitud_base from solicitud where codigo=" & pCod_Solicitud

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Comprobar_Solicitud_Base", ex, mConsulta)
        End Try

    End Function

    Public Function Numero_Solicitudes_con_Solicitud_Base_Justificaciones(ByVal pCod_Solicitud As String) As Object Implements PresenciaDAO.Numero_Solicitudes_con_Solicitud_Base_Justificaciones
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT count(*) as total_solicitudes from solicitud where cod_solicitud_base=" & pCod_Solicitud & " and ESTADO not in ('D','R','P')"

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Numero_Solicitudes_con_Solicitud_Base", ex, mConsulta)
        End Try

    End Function

    Public Function Numero_Solicitudes_con_Solicitud_Base_Aprobaciones(ByVal pCod_Solicitud As String) As Object Implements PresenciaDAO.Numero_Solicitudes_con_Solicitud_Base_Aprobaciones
        Dim mConsulta As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            mConsulta = "SELECT count(*) as total_solicitudes from solicitud where cod_solicitud_base=" & pCod_Solicitud & " and ESTADO not in ('D','R','A')"

            mCommand.CommandText = mConsulta
            mCommand.Connection = mConexion
            mReader = mCommand.ExecuteReader()
            Return mReader

        Catch ex As Exception
            Trata_Error("Error en Numero_Solicitudes_con_Solicitud_Base", ex, mConsulta)
        End Try

    End Function

    Public Function Lista_Justificaciones_A_Compensar(ByRef pDatos As Object, _
                                    Optional ByVal pCodigoIncidencia As Integer = -1, _
                                    Optional ByVal pID_Usuario As String = Nothing, _
                                    Optional ByVal pEstado As String = Nothing _
                                    ) As Boolean Implements PresenciaDAO.Lista_Justificaciones_A_Compensar
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrder As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            mCommand.Connection = mConexion

            mConsulta = " SELECT ID_Justificacion, JC.Cod_Justificacion, Duracion, Factor_Compensacion, Tipo_Factor, "
            mConsulta &= " Dni_Empl, Fecha_Justificada, Estado, "
            mConsulta &= " ( SELECT SUM(Duracion_Minutos) "
            mConsulta &= "   FROM COMPENSACIONES "
            mConsulta &= "   WHERE Id_Justificacion_Compensada = JC.Id_Justificacion"
            mConsulta &= " ) Compensado, "
            mConsulta &= " add_months"
            mConsulta &= " 	("
            mConsulta &= " 	add_months"
            mConsulta &= " 		("
            mConsulta &= " 		J.fecha_justificada + substr(I.caducidad_compensacion,5,2),"
            mConsulta &= " 		substr(I.caducidad_compensacion,3,2)"
            mConsulta &= " 		), "
            mConsulta &= " 	substr(I.caducidad_compensacion,0,2)*12"
            mConsulta &= " 	) Fecha_Caducidad"
            mConsulta &= " FROM JUSTIFICACIONES_A_COMPENSAR JC, JUSTIFICACIONES J, INCIDENCIAS I"
            mWhere = " WHERE JC.Cod_Justificacion = J.Cod_Justificacion "
            mWhere &= " AND J.Cod_Incidencia = I.Cod_Incidencia "


            If pCodigoIncidencia <> -1 Then
                mWhere &= " AND J.Cod_Incidencia = " & pCodigoIncidencia
            End If

            If Not pID_Usuario Is Nothing Then
                mWhere &= " AND Dni_Empl = '" & pID_Usuario & "'"
            End If

            If Not pEstado Is Nothing Then
                mWhere &= " AND Estado IN ('" & pEstado & "')"
            End If

            mOrder = " ORDER BY Fecha_Justificada, ID_Justificacion"
            mConsulta &= mWhere & mOrder
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Justificaciones_A_Compensar", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_JustificacionACompensar(ByVal pIDJust As Long, ByVal pEstado As String) As Long Implements PresenciaDAO.Actualiza_JustificacionACompensar

        'actualiza la solicitud

        Dim mConsulta As String

        Try
            mConsulta = " UPDATE JUSTIFICACIONES_A_COMPENSAR SET "
            If Not IsNothing(pEstado) Then
                mConsulta = mConsulta & " ESTADO = '" & pEstado & "' "
            End If
            'mConsulta = mConsulta & " ,DNI = '" & pDNI & "'"
            mConsulta = mConsulta & " WHERE Id_Justificacion = " & pIDJust
            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_JustificacionACompensar", ex, mConsulta)
        End Try
    End Function
    Public Function Inserta_CompensacionDeJustificacion(ByVal pCodJustCompensacion As Long, _
                                        ByVal pCodJustCompensada As Long, _
                                        ByVal pDuracion As Long) As Boolean Implements PresenciaDAO.Inserta_CompensacionDeJustificacion

        Dim mConsulta As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim auxcodigo As Integer


            auxcodigo = 1

            mConsulta = "INSERT INTO COMPENSACIONES(COD_JUSTIFICACION_COMPENSACION,ID_JUSTIFICACION_COMPENSADA,DURACION_MINUTOS)"
            mConsulta = mConsulta & " VALUES(" & pCodJustCompensacion & ", " & pCodJustCompensada & ", " & pDuracion & ")"

            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            auxcodigo = mCommand.ExecuteNonQuery()

            'mConsulta = "SELECT MAX(ID_Justificacion) FROM JUSTIFICACIONES_A_COMPENSAR WHERE Cod_Justificacion = " & pCodJust
            'mCommand.Connection = mConexion
            'mCommand.CommandText = mConsulta
            'mReader = mCommand.ExecuteReader
            'If mReader.Read Then
            '    auxcodigo = NVL(mReader(0), 0)
            'End If
            'mReader.Close()


            Return auxcodigo


        Catch ex As Exception
            Trata_Error("Error en Inserta_CompensacionDeJustificacion", ex, mConsulta)
            Return -1
        End Try

    End Function

    Public Function Inserta_ConsultaInforme(ByVal pSql As String, ByRef pfilas As Integer) As Boolean Implements PresenciaDAO.Inserta_ConsultaInforme

        Dim mConsulta As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object

            mCommand.Connection = mConexion
            mCommand.CommandText = pSql
            pfilas = mCommand.ExecuteNonQuery()

            Return True


        Catch ex As Exception
            Trata_Error("Error en Inserta_ConsultaInforme", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Actualiza_ConsultaInforme(ByVal pSql As String, ByRef pfilas As Integer) As Boolean Implements PresenciaDAO.Actualiza_ConsultaInforme

        Dim mConsulta As String

        Try
            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object

            mCommand.Connection = mConexion
            mCommand.CommandText = pSql
            pfilas = mCommand.ExecuteNonQuery()

            Return True


        Catch ex As Exception
            Trata_Error("Error en Actualiza_ConsultaInforme", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Lista_Parametros_Word(ByRef pDatos As System.Data.DataSet, Optional ByVal pInforme As String = "") As Boolean Implements PresenciaDAO.Lista_Parametros_Word
        'da lista de parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT * "
            mConsulta = mConsulta & " FROM promainf_parametro"
            mConsulta = mConsulta & " WHERE codigo_inf='" & pInforme & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Parametros_Word", ex, mConsulta)
            Return False
        End Try

    End Function

    Function DameValorCampo_Word(ByRef pDatos As System.Data.DataSet, Optional ByVal pInforme As String = "", Optional ByVal pCod_Solic As String = "", Optional ByVal pSQL As String = "") As Boolean Implements PresenciaDAO.DameValorCampo_Word
        'devuelve datos de un parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT * "
            mConsulta = mConsulta & " FROM promainf_parametros "
            mConsulta = mConsulta & " WHERE codigo_inf=" & pInforme


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Parametros_Word", ex, mConsulta)
            Return False
        End Try
    End Function


    Function Lista_Compensaciones(ByRef pDatos As Object, _
                                Optional ByVal pCodigoJustificacionCompensacion As Long = -1, _
                                Optional ByVal pIDJustificacionCompensada As Long = -1, _
                                Optional ByVal pCodJustificacion As Long = -1) As Boolean Implements PresenciaDAO.Lista_Compensaciones
        Dim mConsulta As String
        Dim mWhere As String
        Dim mOrder As String
        Dim mCommand As New OleDb.OleDbCommand

        Try
            mCommand.Connection = mConexion

            mConsulta = " SELECT ID_JUSTIFICACION_COMPENSADA,COD_JUSTIFICACION_COMPENSACION,DURACION_MINUTOS "
            mConsulta &= " FROM COMPENSACIONES, JUSTIFICACIONES_A_COMPENSAR "
            mConsulta &= " WHERE COMPENSACIONES.ID_JUSTIFICACION_COMPENSADA  = JUSTIFICACIONES_A_COMPENSAR.ID_JUSTIFICACION"




            If pCodigoJustificacionCompensacion <> -1 Then
                mWhere &= " AND Cod_Justificacion_Compensacion = " & pCodigoJustificacionCompensacion
            End If

            If pIDJustificacionCompensada <> -1 Then
                mWhere &= " AND ID_justificacion_compensada = " & pIDJustificacionCompensada
            End If

            If pCodJustificacion <> -1 Then
                mWhere &= " AND cod_justificacion = " & pCodJustificacion
            End If



            mConsulta &= mWhere & mOrder
            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            'Conectarse, buscar datos y desconectarse de la base de datos 
            mDataAdapter.Fill(pDatos)
            Return True
        Catch ex As Exception
            Trata_Error("Error en Lista_Compensaciones", ex, mConsulta)
            Return False
        End Try

    End Function
    Public Function IntervalosCumplimiento(ByRef pDatos As System.Data.DataSet, Optional ByVal pHorario As String = "") As Boolean Implements PresenciaDAO.IntervalosCumplimiento
        'da lista de parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT * "
            mConsulta = mConsulta & " FROM IntervaloCumplimiento"
            mConsulta = mConsulta & " WHERE cod_horario='" & pHorario & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Parametros_Word", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Intervalos(ByRef pDatos As System.Data.DataSet, Optional ByVal pHorario As String = "") As Boolean Implements PresenciaDAO.Intervalos
        'da lista de parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "SELECT * "
            mConsulta = mConsulta & " FROM Intervalos"
            mConsulta = mConsulta & " WHERE cod_horario='" & pHorario & "'"


            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en Lista_Parametros_Word", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Grupos_Consulta_Nombre(ByVal pCodigo As String) As String Implements PresenciaDAO.Grupos_Consulta_Nombre
        Dim mSQL As String

        Try
            mSQL = "SELECT desc_grupo from gruposconsulta where cod_grupo=" & pCodigo

            Dim mCommand As New OleDb.OleDbCommand
            Dim mReader As Object
            Dim mSalida As String
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            If mReader.Read Then
                mSalida = NVL(mReader(0), "")
            End If
            mReader.Close()

            If mSalida = "NO DISP" Then mSalida = "0"
            Return mSalida
        Catch ex As Exception
            Trata_Error("Error en Valor_acumulador", ex, mSQL)
        End Try
    End Function

    Public Function Estado_Solicitud(ByVal pCod_Solicitud As String) As String Implements DAO.PresenciaDAO.Estado_Solicitud
        'Calcula el Horario asignado a una persona en un dia.
        Dim mLista As String
        Dim mCommand As New OleDb.OleDbCommand
        Dim mReader As Object
        Dim mSQL As String
        Dim mCodigo As String

        Try
            mCommand.Connection = mConexion
            mSQL = "select estado from solicitud where cod_solicitud=" & pCod_Solicitud
            mCommand.CommandText = mSQL
            mReader = mCommand.ExecuteReader()
            mReader.Read()
            mCodigo = mReader(0)
            mReader.Close()
            Return mCodigo

        Catch ex As Exception
            Trata_Error("Error en Estado_Solicitud", ex, mSQL)
        End Try
    End Function

    Public Function ListaEmpleadosGrupoTrabajo(ByRef pDatos As DataSet, Optional ByVal pCodGrupo As Integer = 0, Optional ByRef pError As String = "") As Boolean Implements PresenciaDAO.ListaEmpleadosGrupoTrabajo

        'da lista de parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select count(*) total from asociausuariogrupotrabajo "
            mConsulta = mConsulta & " where COD_GRUPOTRABAJO = " & pCodGrupo & " AND (FECHA_HASTA IS NULL OR FECHA_HASTA > TO_CHAR(SYSDATE,'DD/MM/YYYY')) "

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en ListaEmpleadosGrupoTrabajo", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function ListaAsociacionesTarjetasUsuario(ByRef pDatos As DataSet, Optional ByVal pID_Usuario As String = Nothing) As Boolean Implements PresenciaDAO.ListaAsociacionesTarjetasUsuario

        'da lista de parametros de un documento word
        Dim mConsulta As String
        Dim mWhere As String

        Try
            mConsulta = "select * from tarjetasasociadas "
            mConsulta = mConsulta & " where DNI_EMPL = '" & pID_Usuario & "'"
            mConsulta = mConsulta & " order by fecha_hora_alta asc"

            Dim mDataAdapter As New OleDb.OleDbDataAdapter(mConsulta, mConexion)
            pDatos = New DataSet
            mDataAdapter.Fill(pDatos)

            Return True

        Catch ex As Exception
            Trata_Error("Error en ListaAsociacionesTarjetasUsuario", ex, mConsulta)
            Return False
        End Try

    End Function

    Public Function Elimina_Asignacion_Tarjeta_Usuario(Optional ByVal pID_Usuario As String = "", Optional ByVal pPan_Tarjeta As String = "") As Boolean Implements PresenciaDAO.Elimina_Asignacion_Tarjeta_Usuario
        Dim mConsulta As String
        Dim mWhere As String
        Try
            Dim mCommand As New OleDb.OleDbCommand
            mConsulta = "DELETE TarjetasAsociadas "
            If pID_Usuario <> "" Then
                mWhere = " WHERE DNI_EMPL = '" & pID_Usuario & "'"
            End If
            If pPan_Tarjeta <> "" Then
                If mWhere <> "" Then
                    mWhere = mWhere & " AND "
                Else
                    mWhere = " WHERE "
                End If
                mWhere = mWhere & " PAN_TARJETA = '" & pPan_Tarjeta & "'"
            End If
            mConsulta = mConsulta & mWhere
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trata_Error("Error en Elimina_Asignacion_Tarjeta_Usuario", ex, mConsulta)
        End Try

    End Function

    Public Function Inserta_TarjetaAsociada(ByVal pDNI As String, ByVal pPan_tarjeta As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean Implements PresenciaDAO.Inserta_TarjetaAsociada

        Dim mConsulta As String
        Dim mFechas As String

        Dim mCommand_valor As New OleDb.OleDbCommand
        Dim mReader As Object

        Try
            'antes comprobamos si existe otro asociacion la cual pise ésta:

            mFechas = "SELECT count(*) FROM TarjetasAsociadas where dni_empl='" & pDNI & "'"
            mFechas = mFechas & " and (FECHA_HORA_BAJA is null or FECHA_HORA_BAJA > to_char(sysdate,'dd/mm/yyyy'))"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mFechas
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()

            '***************************************************
            'Grabamos
            '***************************************************

            mConsulta = "insert into TarjetasAsociadas (dni_empl, pan_tarjeta,fecha_hora_alta, fecha_hora_baja) " _
                & " values ('" & pDNI & "'," & pPan_tarjeta & ",'" & pFecha_Alta & "',"
            mConsulta &= IIf(pFecha_Baja <> "", "'" & pFecha_Baja & "'", "null") & ")"

            Dim mCommand As New OleDb.OleDbCommand
            mCommand.Connection = mConexion
            mCommand.CommandText = mConsulta
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Inserta_TarjetaAsociada", ex, mConsulta)
        End Try

    End Function

    Public Function Actualiza_TarjetaAsociada(ByVal pDNI As String, ByVal pPan_tarjeta As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean Implements PresenciaDAO.Actualiza_TarjetaAsociada
        Dim mSQL As String
        Dim mCommand As New OleDb.OleDbCommand

        Dim mWhere As String
        Dim mFechas As String
        Try


            Dim mCommand_valor As New OleDb.OleDbCommand
            Dim mReader As Object

            mFechas = "SELECT count(*) FROM TarjetasAsociadas where dni_empl='" & pDNI & "'"
            mFechas = mFechas & " and (FECHA_HORA_BAJA is null or FECHA_HORA_BAJA > to_char(sysdate,'dd/mm/yyyy')) and pan_tarjeta <> '" & pPan_tarjeta & "'"
            mCommand_valor.Connection = mConexion
            mCommand_valor.CommandText = mFechas
            mReader = mCommand_valor.ExecuteReader

            If mReader.Read Then
                If mReader(0) > 0 Then
                    Return False
                End If
            Else
                Return False
            End If
            mReader.Close()


            mSQL = "update TarjetasAsociadas set "

            mSQL &= "fecha_hora_alta ='" & pFecha_Alta & "'"


            If pFecha_Baja = "" Then
                mWhere &= ", fecha_hora_baja =null"
            Else
                mWhere &= ", fecha_hora_baja = '" & pFecha_Baja & "'"
            End If

            mSQL &= mWhere
            mSQL &= " where dni_empl = '" & pDNI & "' and pan_tarjeta =" & pPan_tarjeta & " and to_char(fecha_hora_alta,'dd/mm/yyyy')='" & pFecha_Alta & "'"
            mCommand.Connection = mConexion
            mCommand.CommandText = mSQL
            mCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trata_Error("Error en Actualiza_TarjetaAsociada", ex, mSQL)
        End Try


    End Function

End Class


