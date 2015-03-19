Public Interface PresenciaDAO

    Function Conecta() As Boolean
    Function DesConecta() As Boolean

    Function Elimina_PicadaUsuario(Optional ByVal DNI As String = "", Optional ByVal pan_tarjeta As String = "") As Data.DataSet
    Function Comprueba_PermisoTarjeta(Optional ByVal pDni As String = "", Optional ByVal pTarjeta As String = "") As Data.DataSet
    Function Busqueda_Visita_Tarjeta_Asignada(Optional ByVal pTarjeta As String = "") As Data.DataSet
    Function Comprueba_TarjetaUsuario(Optional ByVal DNI As String = "", Optional ByVal pan_tarjeta As String = "") As Data.DataSet
    Function CompruebaUltimaActivaVisitantes(ByVal pDni As String, ByVal pActivas As String) As Data.DataSet
    Function Comprueba_Permiso(Optional ByVal DNI As String = "", Optional ByVal tipo As String = "") As Data.DataSet
    Function Inserta_PANVisitante(Optional ByVal DNI As String = "", Optional ByVal pPanTarjeta As String = "") As Boolean
    Function Inserta_Visita(Optional ByVal DNI_VISITANTE As String = "", Optional ByVal DNI_VISITADO As String = "") As Boolean
    Function Inserta_Visitante(Optional ByVal DNI As String = "", Optional ByVal NOMBRE As String = "", Optional ByVal APE1 As String = "", Optional ByVal APE2 As String = "", Optional ByVal PAN_TARJETA As String = "", Optional ByVal EMPRESA As String = "") As Boolean
    Function Actualiza_Visitante(Optional ByVal DNI As String = "", Optional ByVal NOMBRE As String = "", Optional ByVal APE1 As String = "", Optional ByVal APE2 As String = "", Optional ByVal PAN_TARJETA As String = "", Optional ByVal EMPRESA As String = "") As Boolean
    Function Comprueba_Visitante(Optional ByVal DNI As String = "") As Data.DataSet
    Function Lista_Visitados_Filtros_Todos(Optional ByVal Filtro As String = "") As Data.DataSet
    Function Lista_Visitados_Filtros_varios(Optional ByVal FiltroNombre As String = "", Optional ByVal FiltroApellidos As String = "") As Data.DataSet
    Function Lista_Visitados_Filtro(Optional ByVal FiltroDNI As String = "") As Data.DataSet
    Function Lista_Visitantes_Filtros_Todos(Optional ByVal Filtro As String = "") As Data.DataSet
    Function Lista_Visitantes_Filtros_varios(Optional ByVal FiltroNombre As String = "", Optional ByVal FiltroApellidos As String = "") As Data.DataSet
    Function Lista_Visitantes_Filtro(Optional ByVal FiltroDNI As String = "") As Data.DataSet
    Function Lista_Visitantes(Optional ByVal FiltroDNI As String = "") As Data.DataSet
    Function CompruebaNombreyClave(ByVal pDni As String, ByVal pClave As String) As Data.DataSet
    Function CompruebaAcceso(ByVal pDni As String) As Data.DataSet
    Function UsuarioTieneClave(ByVal pDni As String) As Data.DataSet
    Function Existe_Clave(ByVal pClave As String, ByVal pTipo As String) As Data.DataSet
    Function Existe_DniUsuario(ByVal pDni As String) As Data.DataSet
    Function TarjetaAsociada(ByVal pNumTarjeta As String) As Data.DataSet
    Function TarjetaAsociada2(ByVal pDni As String) As Data.DataSet
    Function TarjetaTemporal(ByVal pNumTarjeta As String) As Data.DataSet
    Function TarjetaAlta(ByVal pNumTarjeta As String) As Data.DataSet
    Function CompruebaUltimaActiva(ByVal pDni As String, ByVal pActivas As String) As Data.DataSet
    Function ExisteTarjeta(ByVal pPan As String, ByVal tipo As String) As Data.DataSet
    Function DevuelveSecuenciaTarjetas(ByVal pPan As String, ByVal pDni As String, ByVal tipo As String, ByVal pNumTarjeta As String) As Data.DataSet
    Function NuevoNumeroDIP() As Data.DataSet
    Function NuevoNumero(ByVal tipo As String) As Data.DataSet
    Function NuevoClaveEmpleado() As Data.DataSet
    Function InsertaTarjeta(ByVal pPan As String, ByVal pEstado As String, ByVal pTipo As String, ByVal pFechaCaducidad As String) As Boolean
    Function ActualizaTarjeta(ByVal pNumTarjeta As String, ByVal pEstado As String, ByVal pEstadoModificado As String) As Boolean
    Function ActualizaTarjetasAsociadas(ByVal pFechaHora As String, ByVal pDni As String) As Boolean
    Function ActualizaTarjetasAsociadasVisitantes(ByVal pFechaHora As String, ByVal pDni As String) As Boolean
    Function ActualizaAnulacion(ByVal pNumTarjeta As String, ByVal pFechaAnulacion As String) As Boolean
    Function EliminaTarjeta(ByVal pPan As String) As Boolean
    Function EliminaTarjetaTemporal(ByVal pPan As String) As Boolean
    Function EliminaTarjetaAsociada(ByVal pPan As String) As Boolean
    Function ActualizaEmpTarjeta(ByVal pNumTarjeta As String, ByVal pDni As String) As Boolean
    Function InsertaTAsociada(ByVal pPan As String, ByVal pDni As String, ByVal pFechaHoraAlta As String, ByVal pFechaHoraBaja As String) As Boolean
    Function InsertaTAsociadaVis(ByVal pPan As String, ByVal pDni As String, ByVal pFechaHoraAlta As String, ByVal pFechaHoraBaja As String) As Boolean

    'tabla grupos Scap32
    Function Lista_GruposConsulta() As Data.DataSet
    Function Lista_GruposTrabajo() As Data.DataSet
    Function Lista_GruposPrivilegio() As Data.DataSet
    Function Lista_GruposConsultaPerteneceA(ByVal pDni As String) As Data.DataSet
    Function Lista_AsociaGrupoTrabajo(ByVal pDni As String) As Data.DataSet
    Function Lista_GruposPrivilegioPerteneceA(ByVal pDni As String) As Data.DataSet
    Function MeteGruposConsultaHijos(ByVal pCodigo As String) As Data.DataSet

    'MOD_SCAPINI
    Function ValorScapini_Personal(ByVal pDni As String) As Data.DataSet
    Function ValorScapini_Personal2(ByVal pDni As String) As Data.DataSet
    Function Guardar_ValorScapini_Personal(ByVal pDni As String, ByVal pCampo As String, ByVal pValor As String, ByVal pTipo As String) As Boolean
    Function Guardar_ValorScapini(ByVal pDni As String, ByVal pCampo As String, ByVal pValor As String, ByVal pTipo As String) As Boolean
    Function ValorScapini1(ByVal pDni As String) As Data.DataSet
    Function ValorScapini2(ByVal pDni As String) As Data.DataSet
    Function Empleados_ProveedoresTipo(ByVal pDni As String) As Data.DataSet

    'tablas de tarjetas y perfiles (SCAPImpTarjeta)
    Function CargaCampo(ByVal pCPerfil As String, ByVal pCDato As String) As Data.DataSet
    Function Obten_Valor(ByVal PCadSQL As String) As Data.DataSet
    Function CargaPerfiles() As Data.DataSet
    Function VisualizaTarjeta(ByVal pCodigoPerfil As String) As Data.DataSet
    Function VisualizarTarjeta(ByVal pCodigo As String) As Data.DataSet
    Function Guarda_DatosPerfilTarjeta_Coordenadas(ByVal pCPerfil As String, ByVal pCDato As String, ByVal pX As String, ByVal pY As String) As Boolean
    Function Guarda_DatosPerfilTarjeta(ByVal pCPerfil As String, ByVal pCDato As String, ByVal pFuente As String, ByVal pNegrita As String, ByVal pFoto As String, ByVal pX As String, ByVal pY As String, ByVal pAlto As String, ByVal pAncho As String, ByVal pTamano As String, ByVal pColor As String, ByVal pSource As String) As Boolean
    Function Muestra_DatosPerfilTarjeta(ByVal pCPerfil As String, ByVal pCDato As String) As Data.DataSet
    Function ListaDatosPerfilTarjeta(ByVal pCodigo As String) As Data.DataSet
    Function InsertaDatosPerfilTarjeta(ByVal pCodigo As String, ByVal pNuevoCodigo As String) As Boolean
    Function CargaDatosPerfilTarjeta(ByVal pCodigo As Integer) As Data.DataSet
    Function ListaPerfilesTarjeta() As Data.DataSet
    Function Elimina_DatosPerfilTarjeta(ByVal pCodigo As String, ByVal pDato As String) As Boolean
    Function Elimina_Perfil(ByVal pCodigo As String) As Boolean
    Function Muestra_Perfil(ByVal pCodigo As String) As Data.DataSet
    Function Guardar_Perfil(ByVal pCodigo As Integer, ByVal pNombrePerfil As String, ByVal pFondo As String) As Boolean
    Function Insertar_Perfil(ByVal pCodigo As String, ByVal pNombrePerfil As String) As Boolean


    'tabla Empleados

    Function Comprueba_clave(Optional ByVal pClave As String = "", Optional ByVal pDni As String = "") As Data.DataSet
    Function ObtenerDNI(Optional ByVal usuario As String = "") As Data.DataSet
    Function ListaEmpleadosDNI(Optional ByVal pDNI As String = "") As Data.DataSet
    Function Ultima_Tarjeta(Optional ByVal pDNI As String = "", Optional ByVal tipo As String = "") As Data.DataSet
    Function ListaProveedoresDNI(Optional ByVal pDNI As String = "") As Data.DataSet
    Function Lista_Empleados(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Object
    Function Lista_Empleados2(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Data.DataSet
    Function Listado_Empleados(Optional ByVal FiltroDNI As String = "", Optional ByVal FiltroEmail As String = "", Optional ByVal Orden As String = "", Optional ByVal pUsuariosSinResponsable As Boolean = False) As Data.DataSet
    Function Busqueda_Empleados(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pOrden As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pCalcula_Saldo As String = "") As Object
    Function Busqueda_Personal(Optional ByVal pDNI As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pTarjeta As String = "") As Data.DataSet
    Function Busqueda_Personal_Tarjeta(Optional ByVal pDNI As String = "", Optional ByVal pApellidos As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pTarjeta As String = "", Optional ByVal pTipo As Integer = 0) As Data.DataSet
    Function Inserta_Empleado(ByVal pDNI As String, ByVal pNombre As String, ByVal pApe1 As String, ByVal pApe2 As String, ByVal pClave_Empleado As String, ByVal pCentro As String, ByVal pCargo As String, ByVal pEmail As String, ByVal pTelefono As String, ByVal pCalcula_Saldo As Boolean, Optional ByVal pAdmin As Boolean = False, Optional ByVal pEmpresa As String = "", Optional ByVal pFecha_Antiguedad As String = "", Optional ByVal pUsuarioLdap As String = "") As Boolean
    Function Elimina_Empleado(ByVal pDNI As String) As Boolean
    Function Elimina_TarjetasAsociadas(ByVal pDNI As String) As Boolean
    Function Elimina_VisitasRealizadas(ByVal pDNI As String) As Boolean
    Function Elimina_VisitasRegistradas(ByVal pDNI As String) As Boolean
    Function Elimina_AsociaUsuarioGrupoTrabajo(ByVal pDNI As String) As Boolean
    Function Elimina_PerteneceA(ByVal pDNI As String) As Boolean
    Function Actualiza_Empleado(ByVal pDNI As String, Optional ByVal pNombre As String = Nothing, Optional ByVal pApe1 As String = Nothing, Optional ByVal pApe2 As String = Nothing, Optional ByVal pClave_Empleado As String = Nothing, Optional ByVal pCentro As String = Nothing, Optional ByVal pCargo As String = Nothing, Optional ByVal pEmail As String = Nothing, Optional ByVal pTelefono As String = Nothing, Optional ByVal pCalcula_Saldo As Integer = -10, Optional ByVal pAdmin As Integer = -10, Optional ByVal pEmpresa As String = "", Optional ByVal pFecha_Antiguedad As String = "", Optional ByVal pUsuarioLdap As String = "") As Boolean
    Function Actualiza_Campo_Empleado(ByVal pDNI As String, ByVal pCampo As String, ByVal pValor As String) As Boolean
    Function Lista_Empleados_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Lista_Empleados_DNI_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Lista_Empleado(ByRef pDatos As DataSet, ByVal pDNI As String) As Boolean
    Function Lista_Empleados_Responsable_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Lista_Empleados_Asignacion_Responsable(ByRef pDatos As DataSet, ByVal Lista_DNI As String) As Boolean
    Function Lista_Empleados_Responsables_Grupos_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function VisualizaEmpleadoTarjeta() As Data.DataSet
    Function MuestraHistorial(ByVal pDNI As String, ByVal tipo As Boolean) As Data.DataSet

    'tabla log_visita
    Function Registra_LOG_Visita(ByVal pUsuario As String, ByVal pObserva As String) As Boolean

    'tabla eventos
    Function Lista_Eventos(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object
    Function Lista_EventosGrupos(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object
    Function Lista_Primer_Evento(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object
    Function Lista_Ultimo_Evento(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pUsuario As String = "") As Object
    Function Inserta_Evento(ByVal pDNI As String, ByVal pFecha As DateTime, ByVal pSentido As String, ByVal pCod_recurso As Long, Optional ByVal pTarjeta As String = "", Optional ByVal pCod_incidencia As Long = 0, Optional ByVal pPermitido As Boolean = False, Optional ByVal pIP As String = "") As Boolean
    Function Elimina_Evento(ByVal pCod_Evento As String) As Boolean

    Function Lista_Eventos_Visor(ByRef pDatos As Object, _
                                Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", _
                                Optional ByVal pHoraDesde As String = "", Optional ByVal pHoraHasta As String = "", _
                                Optional ByVal pListaUsuarios As String = "", _
                                Optional ByVal pListaGrupos As String = "", _
                                Optional ByVal pSinHerencia As Boolean = True, _
                                Optional ByVal pTipoEvento As String = "", _
                                Optional ByVal pAgruparEvento As Boolean = False, _
                                Optional ByVal pPermitido As String = "", _
                                Optional ByVal pListaCodRecurso As String = "", _
                                Optional ByVal pListaGrupoRecurso As String = "", _
                                Optional ByVal pOrden As String = "") As Object

    Function CambiaSentido_Evento(ByVal pCod_Evento As String, ByVal pSentido As String) As Boolean

    'tabla horarios
    Function Lista_Horarios(Optional ByVal pCod_Horario As String = "") As Object
    Function Lista_Horarios_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_Horario As String = "") As Boolean
    Function Lista_Intervalos_Horario(ByVal pcod_horario As String) As DataSet
    Function Horario_Dia(ByVal pDni As String, ByVal pFecha_Desde As Date) As String

    'tabla incidencias
    Function Lista_Incidencias(Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As Boolean = False) As Object

    Function Lista_Incidencias_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDescripcion As String = "", Optional ByVal pCodigoIncidenciaDeCompensacion As Integer = -1) As Boolean
    Function Lista_Incidencias_Contrato_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDescripcion As String = "", Optional ByVal pCod_Contrato As Integer = -1) As Boolean
    Function Inserta_Incidencia(ByVal pCod_Incidencia As String, ByVal pDesc_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal pSeleccionable_TVR As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Integer
    Function Actualiza_Incidencia(ByVal pCod_Incidencia As String, Optional ByVal pDesc_Incidencia As String = "", Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal pSeleccionable_TVR As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean
    Function Intervalo_minimo_incidencia(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean
    Function Intervalo_minimo_duracion(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean
    Function Intervalo_minimo_duracion_Incidencia_TipoContrato(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCodContrato As Integer = -1) As Boolean
    Function Intervalo_maximo_duracion(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean
    Function Intervalo_maximo_duracion_Incidencia_TipoContrato(ByRef pDatos As DataSet, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCodContrato As Integer = -1) As Boolean

    Function Elimina_Incidencia(ByVal pCod_Incidencia As String) As Boolean
    Function Lista_Incidencias_TVR_Dataset(ByRef pDatos As DataSet) As Boolean


    'tabla tipo_contrato
    Function Lista_Tipo_Contrato_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_Tipo As String = "", Optional ByVal pDescripcion As String = "") As Boolean
    Function Inserta_Tipo_Contrato(ByVal pDesc_Tipo As String, Optional ByVal pObs_Tipo As String = "") As Integer
    Function Actualiza_Tipo_Contrato(ByVal pCod_Tipo As String, Optional ByVal pDesc_Tipo As String = "", Optional ByVal pObs_Tipo As String = "") As Boolean
    Function Elimina_Tipo_Contrato(ByVal pCod_Tipo As String) As Boolean


    'tabla tipocontrato_incidencia
    Function Lista_TC_Incidencias_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_TC As Integer = -1, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pSeleccionable As String = "", Optional ByVal pCodigoIncidenciaDeCompensacion As Integer = -1) As Boolean
    Function Lista_TC_Incidencias_Excepciones_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_TC As Integer = -1, Optional ByVal pCod_Incidencia As Integer = -1) As Boolean
    Function Inserta_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean
    Function Actualiza_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String, Optional ByVal pTipo As String = "", Optional ByVal pTipoFijo As String = "", Optional ByVal pMaximo As String = "", Optional ByVal pFecha_Base As String = "", Optional ByVal pFecha_Termino As String = "", Optional ByVal pTiempo_Maximo As String = "", Optional ByVal pOrden As String = "", Optional ByVal pSeleccionable As String = "", Optional ByVal pMaximo_Horas As String = "", Optional ByVal Maximo_Duracion As Integer = 0, Optional ByVal Minimo_Duracion As Integer = 0, Optional ByVal Tiempo_Minimo As Integer = 0, Optional ByVal Naturales As Integer = 1) As Boolean
    Function Elimina_TC_Incidencia(ByVal pCod_TC As String, ByVal pCod_Incidencia As String) As Boolean

    'tabla asocia_emp_tipocontrato
    Function Lista_Asocia_Emp_Tipocontrato_Dataset(ByRef pDatos As DataSet, ByVal pDNI As String) As Boolean
    Function Lista_Asocia_Emp_Tipocontrato_Fechas_Dataset(ByRef pDatos As DataSet, ByVal pDNI As String, ByVal pDesde As String, ByVal pHasta As String) As Boolean
    Function Lista_Asocia_Emp_Tipocontrato_Dia_Dataset(ByRef pDatos As DataSet, ByVal pDNI As String, ByVal pDia As String) As Boolean
    Function Inserta_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean
    Function Actualiza_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo_Antiguo As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta_Antigua As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean
    Function Elimina_Asocia_Emp_Tipocontrato(ByVal pDNI As String, ByVal pCod_Tipo As String, ByVal pFecha_Alta As String) As Boolean

    'tabla diario y auxiliares
    Function Datos_Diario(ByVal pDNI As String, ByVal pFechaDesde As Date, ByVal pFechaHasta As Date) As Object
    Function Lee_Diario(ByRef pDatos As DataSet, ByVal pDNI As String, ByVal pFecha As String) As Boolean
    Function Datos_Diario_DataSet(ByVal pDNI As String, ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, ByRef pSalida As DataSet, Optional ByVal pError As String = "") As Boolean
    Sub Calcula_Diario(ByVal pDNI As String, ByVal pFecha As Date)
    Function Elimina_Diario(ByVal pDNI As String, ByVal pFechaDesde As String, Optional ByVal pFechaHasta As String = "", Optional ByRef pError As String = "") As Boolean
    Function Lista_Intervalos_Presencia(Optional ByVal pDNI As String = "", Optional ByVal pFecha As String = "") As Object
    Function Cadena_de_Picadas(ByVal pDNI As String, ByVal pFecha As Date) As String
    Function Cadena_de_Eventos(ByVal pDNI As String, ByVal pFecha As Date) As String
    Function Cadena_de_Justificaciones(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pTipo As String = "") As String
    Function Cadena_de_Solicitudes(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pTipo As String = "") As String
    Function Cadena_de_Intervalos(ByVal pcod_horario As Integer, Optional ByVal pFormato As String = "N") As String
    Function Cadena_de_Intervalos_Recuperacion(ByVal pcod_horario As Integer) As String
    Function Cadena_de_Intervalos_Opcional(ByVal pcod_horario As Integer) As String
    Function Acumuladores_Usuario(ByVal pDNI As String, ByVal pFecha As Date, Optional ByVal pSoloFavoritos As Boolean = False) As Object
    Function Elimina_del_Diario(ByVal pTipo As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pDNI As String = "", Optional ByVal pGrupo As String = "") As Boolean


    'estas tablas son para la actualizacion del diario cuando se modifican asignaciones en calendarios
    'o en grupos de trabajo.
    Function Cambio_Calendario_Festivo(ByVal pFecha As String, ByVal pCod_Cal As String) As Boolean
    Function Cambio_Calendario_Laborable(ByVal pCod_Cal As String, ByVal pDia As Integer) As Boolean


    'tabla Solicitud
    Function Lista_Solicitudes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False, Optional ByVal Cambio_Grupo As String = Nothing) As Object
    Function Lista_SolicitudesCuadrantes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As String = "", Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False, Optional ByVal Cambio_Grupo As String = Nothing) As Object
    Function Lista_Solicitudes_Grupos(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object
    Function Lista_Solicitudes_Grupos_Todos(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object
    Function Lista_Solicitudes_Grupos_Cuadrantes(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As String = "", Optional ByVal pOrden As String = "Usuario", Optional ByVal Cambio_Grupo As String = Nothing) As Object
    Function Inserta_Solicitud(ByVal pCODIGO As Long, ByVal pEstado As String, ByVal pDNI As String, ByVal pFecha As Date, ByVal pCod_Incidencia As Integer, ByVal pDesde As String, ByVal pHasta As String, ByVal pObservaciones As String, ByVal pSiguiente_Responsable As String, Optional ByVal pTipoEfecto As Integer = 1, Optional ByVal pCambioGrupo As Boolean = False, Optional ByVal pCod_solicitud_base As Long = 0) As Boolean
    Function Actualiza_Solicitud(ByVal pCODIGO As Long, Optional ByVal pEstado As String = Nothing, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pDesde As String = Nothing, Optional ByVal pHasta As String = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pID_Sig_Responsable As String = Nothing, Optional ByVal pCod_Justificacion As Long = -1, Optional ByVal pUltimo_Responsable As String = Nothing, Optional ByVal pTipo As String = Nothing, Optional ByVal pCambioGrupo As String = Nothing, Optional ByVal pCod_solicitud_base As Long = 0) As Boolean
    Function Elimina_Solicitud(ByVal pCodigo As Long) As Boolean
    Function Elimina_Solicitud(ByVal pListaCodigos As String) As Boolean
    Function Elimina_Justificacion1(ByVal pListaCodigos As String) As Boolean
    Function Lista_SolicitudesAprobadas(ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Object
    Function Lista_Solicitudes_Movimiento(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean
    Function Lista_SolicitudesAprobadas_Movimiento(ByRef pDatos As DataSet, ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean
    Function Lista_SolicitudesAprobadas_Grupos(ByVal pLista_Responsables As String, Optional ByVal pOrden As String = "Usuario", Optional ByVal pEstado As String = "A") As Object
    Function Lista_Solicitudes_Movimiento_Grupos(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCodigo As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pLista_Ultimo_Responsable As String = "", Optional ByVal pCodigoIncidencia As Long = -1, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean
    Function Lista_SolicitudesAprobadas_Movimiento_Grupos(ByRef pDatos As DataSet, ByVal pLista_Responsables As String, Optional ByVal pOrdenUsuario As Boolean = False, Optional ByVal pOrdenFechasol As Boolean = False) As Boolean
    Function Comprobar_Solicitud_Base(ByVal pCod_Solicitud As String) As Object
    Function Numero_Solicitudes_con_Solicitud_Base_Justificaciones(ByVal pCod_Solicitud As String) As Object
    Function Numero_Solicitudes_con_Solicitud_Base_Aprobaciones(ByVal pCod_Solicitud As String) As Object
    Function Estado_Solicitud(ByVal pCod_Solicitud As String) As String

    'tabla Justificacion
    Function Lista_Justificaciones(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pDNI As String = "", Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pCod_Justificacion As Long = -1) As Object
    Function Inserta_Justificacion(ByVal pFecha_Justificada As Date, ByVal pDesde_minutos As Integer, ByVal pHasta_minutos As Integer, ByVal pDni As String, ByVal pOperador As String, ByVal pCod_Incidencia As Integer, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing, Optional ByVal pEfecto As Integer = 1) As Long
    Function Actualiza_Justificacion(ByVal pCodigo As Long, ByVal pFecha_Justificada As Date, ByVal pDesde_minutos As Integer, ByVal pHasta_minutos As Integer, Optional ByVal pDni As String = Nothing, Optional ByVal pOperador As String = Nothing, Optional ByVal pCod_Incidencia As Integer = Nothing, Optional ByVal pObservaciones As String = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing, Optional ByVal pEfecto As Integer = 1) As Boolean
    Function Elimina_Justificacion(Optional ByVal pCodigo As Long = Nothing, Optional ByVal pCod_Solicitud As Long = Nothing) As Boolean
    Function Resumen_Justificaciones(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, Optional ByVal pListaDNI As String = "", Optional ByVal pListaIncidencia As String = "") As Object
    Function Resumen_Justificaciones_Dataset(ByRef pDatos As DataSet, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pDNI As String) As Boolean
    Function Detalle_Resumen_Justificaciones_Dataset(ByRef pDatos As DataSet, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pDNI As String, ByVal pCod_Incidencia As String) As Boolean
    Function Actualiza_Justificacion_Cod_Base(ByVal pCodigoJustf As Long, ByVal pCodigoBase As Long) As Boolean


    'tabla Asignacion_Responsable
    Function Lista_Asignacion_Responsable(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Lista_Responsables As String = "") As Object
    Function Inserta_Asignacion_Responsable(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean
    Function Actualiza_Asignacion_Responsable(ByVal pID_Responsable_Antiguo As String, ByVal pID_Responsable_Nuevo As String) As Boolean
    Function Actualiza_Asignacion_Responsable_Siguiente(ByVal pID_Responsable) As Boolean
    Function Mantiene_Asignacion_Responsable_Siguiente(ByVal pID_Responsable) As Boolean
    Function Inserta_Asignacion_Responsable_Comprobaciones(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean
    Function Elimina_Asignacion_Responsable(Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Responsable As String = "") As Boolean
    Function Lista_Asignaciones_Responsables(ByRef pDatos As DataSet, Optional ByVal pResponsable As String = "") As Boolean
    Function Lista_Usuarios_Responsable(ByRef pDatos As DataSet, Optional ByVal pResponsable As String = "", Optional ByVal pSinAsignar As Boolean = False) As Boolean
    Function Lista_Asignacion_Responsable_Por_Grupo(ByRef pDatos As DataSet, Optional ByVal pGrupo As String = "", Optional ByVal pResponsable As String = "") As Boolean
    Function Lista_Asignacion_Consultor_Por_Grupo(ByRef pDatos As DataSet, Optional ByVal pGrupo As String = "", Optional ByVal pResponsable As String = "") As Boolean
    Function Elimina_Asignacion_Responsable_Por_Grupo(Optional ByVal pGrupo As String = "", Optional ByVal pID_Responsable As String = "") As Boolean
    Function Elimina_Asignacion_Consultor_Por_Grupo(Optional ByVal pGrupo As String = "", Optional ByVal pID_Responsable As String = "") As Boolean
    Function Inserta_Asignacion_Responsable_Por_Grupo(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean
    Function Inserta_Asignacion_Consultor_Por_Grupo(ByVal pID_Usuario As String, ByVal pID_Responsable As String) As Boolean
    Function Actualiza_Asignacion_Responsable_Siguiente_Por_Grupo(ByVal pGrupo) As Boolean
    Function Actualiza_Asignacion_Consultor_Por_Grupo(ByVal pGrupo) As Boolean
    Function Lista_Asignacion_Responsable_Dataset(ByRef pDatos As DataSet, Optional ByVal pID_Usuario As String = "", Optional ByVal pID_Lista_Responsables As String = "") As Boolean
    Function Lista_Parametros_Word(ByRef pDatos As System.Data.DataSet, Optional ByVal pInforme As String = "") As Boolean
    Function IntervalosCumplimiento(ByRef pDatos As System.Data.DataSet, Optional ByVal pHorario As String = "") As Boolean
    Function Intervalos(ByRef pDatos As System.Data.DataSet, Optional ByVal pHorario As String = "") As Boolean
    Function DameValorCampo_Word(ByRef pDatos As System.Data.DataSet, Optional ByVal pInforme As String = "", Optional ByVal pCod_Solic As String = "", Optional ByVal pSQL As String = "") As Boolean

    'tabla Aprobacion
    Function Lista_Aprobaciones(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "") As Object
    Function Lista_OperacionesPendientes(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "") As Object
    Function Inserta_Aprobacion(ByVal pCod_Solicitud As Long, ByVal pID_Responsable As String, Optional ByVal pOperacion As String = "A", Optional ByVal pDelegado As String = "") As Boolean
    Function Actualiza_Aprobacion(ByVal pCod_Solicitud As Long, ByVal pID_Responsable As String, ByVal pOperacion As String, Optional ByVal pDelegado As String = Nothing, Optional ByVal pCausaDenegacion As String = Nothing) As Boolean
    Function Elimina_Aprobacion(Optional ByVal pCod_Solicitud As Long = -1, Optional ByVal pID_Responsable As String = "", Optional ByVal pOperacion As String = "") As Boolean

    'tabla AutorizadoJustificar
    Function Lista_AutorizadoJustificar(Optional ByVal pDNI As String = "", Optional ByVal pFechaDesde As String = Nothing, Optional ByVal pFechaHasta As String = Nothing) As Object
    Function Lista_AutorizadoConsultar(Optional ByVal pDNI As String = "") As Object
    Function Lista_AutorizadoAprobar(Optional ByVal pDNI As String = "") As Object
    Function Usuario_ResponsableSolicitud(Optional ByVal pUsuario As String = "", Optional ByVal pSolicitud As String = "") As Object
    Function Lista_AutorizadoJustificar_Dataset(ByRef pDatos As DataSet, Optional ByVal pGrupo As String = "") As Boolean
    Function Lista_AutorizadoConsultar_Dataset(ByRef pDatos As DataSet, Optional ByVal pGrupo As String = "") As Boolean
    Function Inserta_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, Optional ByVal pFecha_Hasta As String = "") As Boolean
    Function Existe_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, Optional ByVal pFecha_Hasta As String = "") As Boolean
    Function Actualiza_AutorizadoJustificar(ByVal pDni As String, ByVal pGrupo As String, ByVal pFecha_Desde As String, ByVal pFecha_Desde_Nueva As String, Optional ByVal pFecha_Hasta As String = "") As Boolean
    Function Elimina_AutorizadoJustificar(Optional ByVal pDni As String = "", Optional ByVal pGrupo As String = "", Optional ByVal pFecha_Desde As String = "") As Boolean

    Function Lista_Solicitudes_de_Grupos(Optional ByVal pDni As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal pFechaDesde As String = "", Optional ByVal pFechaHasta As String = "", Optional ByVal pLista_Estados As String = "", Optional ByVal pCodIncidencia As String = "", Optional ByVal pOrden As String = "Usuario") As Object
    Function Lista_Solicitudes_de_Grupos2(ByRef pDatos As Object, ByVal pListaGrupos As String, ByVal pdni As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pLista_Estados As String, ByVal pIncidencia As String, Optional ByVal pOrden As String = "") As Boolean
    Function Lista_Justificaciones_de_Grupos2(ByRef pDatos As Object, ByVal pListaGrupos As String, ByVal pdni As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pLista_Estados As String, ByVal pIncidencia As String, Optional ByVal pOrden As String = "") As Boolean
    Function Lista_Solicitudes_de_Grupos_RowNum(ByRef pDatos As DataSet, ByVal pListaGrupos As String, ByVal pFechaDesde As String, ByVal pFechaHasta As String, ByVal pLista_Estados As String, Optional ByVal pOrden As String = "") As Boolean
    Function Busqueda_Empleados_en_Grupos(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal pApellidos As String = "") As Object
    Function Busqueda_Empleados_en_Asignaciones(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Object
    Function Busqueda_Empleados_De_ResponsablesAut(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Busqueda_Empleados_De_ResponsablesAut_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Busqueda_Empleados_De_ResponsablesJust_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Busqueda_Empleados_De_ResponsablesConsultar(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Busqueda_Empleados_De_ResponsablesConsultar_Dataset(ByRef pDatos As DataSet, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pResponsable As String = "", Optional ByVal pApellidos As String = "") As Boolean
    Function Busqueda_Empleados_Grupos_Privilegios(Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pClave_Empleado As String = "", Optional ByVal pListaGrupos As String = "", Optional ByVal Tipo_Grp As String = "", Optional ByVal pApellidos As String = "") As Object
    Function Actualiza_Saldo(ByVal pDNI As String, ByVal pFECHA As Date) As Boolean

    'Tabla Delegados
    Function Lista_Delegados(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "", Optional ByVal pOrden As String = "Ape1,Ape2,Nombre")
    Function Lista_Delegados_Solicitud(ByVal pID_Delegado As String, ByVal pID_Solicitud As String)
    Function Inserta_Delegados(ByVal pID_Responsable As String, ByVal pID_Delegado As String) As Boolean
    Function Elimina_Delegados(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "")
    Function Elimina_Delegados_Rodas(Optional ByVal pID_Responsable As String = "", Optional ByVal pID_Delegado As String = "") As Boolean
    Function Lista_Responsables(ByRef pDatos As DataSet, ByVal pID_Responsable As String) As Boolean
    Function Lista_Responsables_Libres(ByRef pDatos As DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "") As Boolean
    Function Lista_Responsables_Libres_Justificadores(ByRef pDatos As DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pGrupo As String = "") As Boolean
    Function Lista_Responsables_Libres_Aprobadores(ByRef pDatos As DataSet, ByVal pID_Responsable As String, Optional ByVal pDNI As String = "", Optional ByVal pNombre As String = "", Optional ByVal pApe1 As String = "", Optional ByVal pApe2 As String = "", Optional ByVal pGrupo As String = "") As Boolean

    'tabla grupos consulta
    Function Lista_Grupos_Consulta(Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Object
    Function Lista_Grupos_Consulta_Autorizados_a_Justificar(ByVal pID_Responsable As String) As Object
    Function Lista_Grupos_Consulta_Autorizados_a_Consultar(ByVal pID_Responsable As String) As Object
    Function Lista_Grupos_Consulta_Autorizados_a_Autorizar(ByVal pID_Responsable As String) As Object
    Function Actualiza_Grupos_Consulta(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean
    Function Inserta_Grupos_Consulta(ByVal pDesc_Grupo As String, Optional ByVal pPadre As String = "") As Boolean
    Function Elimina_Grupos_Consulta(ByVal pCod_Grupo As Integer) As Boolean
    Function Lista_Grupos_Consulta_Usuario(ByRef pDatos As DataSet, Optional ByVal pCodigo As String = "", Optional ByVal pPadre As Long = 0, Optional ByVal pcoleccion As Collection = Nothing) As Boolean
    Function Lista_Usuarios_GrupoConsulta(ByRef pDatos As DataSet, ByVal pCodigo As String, Optional ByVal pDistinct As String = "") As Boolean
    Function Grupos_Consulta_Usuario(ByRef pDatos As DataSet, ByVal pCodigo As String) As Boolean
    Function Grupos_Consulta_Usuario_DIPU(ByRef pDatos As DataSet, ByVal pCodigo As String) As Boolean
    Function Grupos_Consulta_Usuario_DIPU_Aprobar(ByRef pDatos As DataSet, ByVal pCodigo As String) As Boolean
    Function Lista_Grupos_Consulta_Dataset(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Boolean
    Function Grupos_Consulta_Nombre(ByVal pCodigo As String) As String

    'tabla grupos privilegios
    Function Lista_Grupos_Privilegios(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByRef pError As String = "") As Boolean
    Function Actualiza_Grupos_Privilegios(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean
    Function Inserta_Grupos_Privilegios(ByVal pDesc_Grupo As String) As Boolean
    Function Elimina_Grupos_Privilegios(ByVal pCod_Grupo As Integer) As Boolean

    'tabla grupos recursos
    Function Lista_Grupos_Recursos(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pCodigoPadre As Long = 0, Optional ByRef pError As String = "") As Boolean
    Function Lista_Grupos_Recursos2(Optional ByVal pCodigo As Long = 0, Optional ByVal pNombre As String = "", Optional ByVal pPadre As Long = 0) As Object
    Function Actualiza_Grupos_Recursos(ByVal pCod_Grupo As Integer, ByVal pDesc_Grupo As String) As Boolean
    Function Inserta_Grupos_Recursos(ByVal pDesc_Grupo As String, Optional ByVal pPadre As String = "") As Boolean
    Function Elimina_Grupos_Recursos(ByVal pCod_Grupo As Integer) As Boolean

    'tabla recursos
    Function Lista_Recursos(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByRef pError As String = "") As Boolean
    Function Lista_Recursos_SinAsignar() As Object
    Function Asigna_Recursos(ByVal pCod_Recurso As Integer, ByVal pPadre As String) As Boolean


    'vista Recursos de Grupo
    Function Lista_RecursosDeGrupo(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByRef pError As String = "") As Boolean


    'tabla diascalendario
    Function Lista_DiasCalendario(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pFecha As String = "", Optional ByRef pError As String = "") As Boolean
    Function Inserta_DiasCalendario(Optional ByVal pCodigo As String = "", Optional ByVal pFecha As String = "", Optional ByVal pDesc_Fecha As String = "", Optional ByVal pCod_Horario As String = "") As Boolean

    'tabla pertenecena
    Function Lista_Pertenecena(ByRef pDatos As DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pDNI As String = "", Optional ByVal pTipoGrp As Long = 0, Optional ByVal pError As String = "") As Boolean
    Function Lee_Pertenecena(ByVal pCodigoGrupo As Long, ByVal pDNI As String, ByVal pTipoGrp As Long) As Boolean
    Function Inserta_Pertenecena(ByVal pCod_Grupo As Integer, ByVal pTipo_Grp As Integer, ByVal pDni_Empl As String) As Boolean
    Function Elimina_Pertenecena(ByVal pCodpertenece As Integer) As Boolean

    'tabla AccesosCalendarioGrupos
    Function Lista_AccesosCalendarioGrupos(ByRef pDatos As DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pCodigoCalendario As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByVal pError As String = "") As Boolean
    Function Lista_AccesosCalendarioGrupos_Extendida(ByRef pDatos As DataSet, Optional ByVal pCodigoGrupo As Long = 0, Optional ByVal pCodigoCalendario As Long = 0, Optional ByVal pCodigoGrupoRecursos As Long = 0, Optional ByVal pError As String = "") As Boolean
    Function Inserta_AccesosCalendarioGrupos(ByVal pCod_Grupo As String, ByVal pCod_Calendario As Integer, ByVal pCod_Recurso As String) As Boolean
    Function Elimina_AccesosCalendarioGrupos(ByVal pCod_Grupo As String, ByVal pCod_Calendario As Integer, ByVal pCod_Recurso As String) As Boolean


    'tabla Calendarios
    Function Lista_Calendarios(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pTipoCalendario As Long = 0, Optional ByVal pClaseCalendario As Long = 0, Optional ByVal pDescCalendario As String = "", Optional ByVal pAnio As String = "", Optional ByVal pError As String = "") As Boolean
    Function Lista_Calendarios_Asociacion(ByRef pDatos As DataSet, ByVal pCod_Grupo As String, ByVal pCod_Recurso As String) As Boolean

    'tabla CalendariosLaborables
    Function Lista_CalendariosLaborables(ByRef pDatos As DataSet, Optional ByVal pCodigo As Long = 0, Optional ByVal pError As String = "") As Boolean

    Function Lista_AsociacionesGruposTrabajo(ByRef pDatos As DataSet, Optional ByVal pID_Usuario As String = Nothing, Optional ByVal pCodGrupo As String = Nothing, Optional ByVal pAño As String = Nothing) As Boolean



    'acumuladores
    Function Valor_Acumulador(ByVal pDNI As String, ByVal pFecha As Date, ByVal pAcumulador As String) As String

    'fecha hora de la base de datos
    Function Ahora() As DateTime

    'Grupos de trabajo
    Function Lista_Grupo_Trabajo(ByRef pDatos As DataSet, Optional ByVal pCodigo As Integer = 0, Optional ByVal pDescGrupo As String = "", Optional ByVal pOrden As String = "") As Boolean

    'asignaciones a Grupos de trabajo
    Function Lista_Asocia_Usuario_Grupo_Trabajo(Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0) As Object
    Function Lista_Asocia_Usuario_Grupo_Trabajo(ByRef pDatos As DataSet, Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0, Optional ByRef pError As String = "") As Boolean
    Function Lista_Asocia_Usuario_Grupo_Trabajo_Fechas(ByRef pDatos As DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, Optional ByVal pCod_Asoc As Integer = 0, Optional ByVal pDNI As String = "", Optional ByVal pCodigo As Integer = 0, Optional ByRef pError As String = "") As Boolean
    Function Actualiza_Asocia_Usuario_Grupo_Trabajo(ByVal pCod_Asoc As Integer, ByVal pDNI As String, ByVal pCodigo As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean
    Function Inserta_Asocia_Usuario_Grupo_Trabajo(ByVal pDNI As String, ByVal pCodigo As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean
    Function Elimina_Asocia_Usuario_Grupo_Trabajo(ByVal pCod_Asoc As Integer) As Boolean

    'asignaciones de Calendarios a Grupos de trabajo

    Function Actualiza_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_DesdeAnt As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean
    Function Inserta_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Boolean
    Function Elimina_Asocia_Grupo_Trabajo_Calendario(ByVal pCodigoGrupo As Integer, ByVal pCodigoCalendario As Integer, ByVal pFecha_Desde As String) As Boolean
    Function Lista_GrupoTrabajo_Calendarios(ByRef pDatos As DataSet, Optional ByVal pCodigo_Grupo As Integer = 0, Optional ByVal pTipo_Cal As Integer = 1) As Boolean


    'tabla noticias
    Function Lista_Noticias_Dataset(ByRef pDatos As DataSet, Optional ByVal pCod_Noticia As String = "", Optional ByVal pDesc_Noticia As String = "", Optional ByVal pFecha As String = "", Optional ByVal pNumero_Max As String = "") As Boolean
    Function Inserta_Noticias(ByVal pDesc_Noticia As String, ByVal pFecha As String, Optional ByVal Obs_Noticia As String = "") As Integer
    Function Actualiza_Noticias(ByVal pCod_Noticia As String, Optional ByVal pDesc_Noticia As String = "", Optional ByVal pFecha As String = "", Optional ByVal Obs_Noticia As String = "") As Boolean
    Function Elimina_Noticias(ByVal pCod_Noticia As String) As Boolean



    'TVR_IP
    Function Lista_TVR_IP(Optional ByVal pIP As String = "", Optional ByVal pOrdenDescripcion As Boolean = False) As DataSet
    Function Inserta_TVR_IP(ByVal pIP As String, Optional ByVal pDescripcion As String = "") As Boolean
    Function Modifica_TVR_IP(ByVal pIP As String, Optional ByVal pDescripcion As String = "") As Boolean
    Function Elimina_TVR_IP(ByVal pIP As String) As Boolean

    Function Lista_Justificaciones_Solicitud(ByVal pFechaDesde As Date, ByVal pFechaHasta As Date, ByVal pDNI As String, Optional ByVal pCod_Incidencia As Integer = -1, Optional ByVal pOrden As String = "", Optional ByVal pFechaOmitir As String = "") As Object
    Function Lista_Justificaciones_Dataset(ByRef pDatos As DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "") As Boolean
    Function Lista_Justificaciones_Solicitud_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pFecha_Omitir As String = "", Optional ByVal pCodigo_Justificacion As String = "") As Integer
    Function Numero_Solicitudes_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String) As Integer
    Function Lista_Justificaciones_Solicitud_Horas_Dataset(ByVal pID_Usuario As String, ByVal pCodigoIncidencia As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, Optional ByVal pCodigo_Solicitud As String = "", Optional ByVal pCodigo_Justificacion As String = "") As Integer
    Function Lista_Solicitudes_Dataset(ByRef pDatos As DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "") As Boolean
    Function Lista_Solicitudes_Dataset_Cuadrante(ByRef pDatos As DataSet, ByVal pID_Usuario As String, ByVal pFecha_Desde As String, ByVal pFecha_Hasta As String, ByVal pEstado As String, Optional ByVal pOrden As String = "", Optional ByVal pIncidencias As String = "") As Boolean
    Function Lista_Solicitudes_Dataset_Completo(ByRef pDatos As System.Data.DataSet, ByVal pID_Usuario As String, Optional ByVal pFecha_Desde As String = "", Optional ByVal pFecha_Hasta As String = "", Optional ByVal pListaEstados As String = "", Optional ByVal pCod_Incidencia As String = "", Optional ByVal pAgrupadoPorFecha As Boolean = True, Optional ByVal pOrden As String = "", Optional ByVal pCod_Solicitud As String = "") As Boolean

    'Asocia_GrupoTrabajo_Calendario
    Function Lista_Asocia_GrupoTrabajo_Calendario(ByRef pDatos As DataSet, Optional ByVal pCodGrupo As Integer = 0, Optional ByVal pCodCalendario As Integer = 0, Optional ByVal pDesde As String = "", Optional ByVal pHasta As String = "", Optional ByVal pFestivo As Integer = 0, Optional ByVal pAnyoFestivo As Integer = 0, Optional ByRef pError As String = "") As Boolean

    'Empleados asignados al grupo de trabajo
    Function ListaEmpleadosGrupoTrabajo(ByRef pDatos As DataSet, Optional ByVal pCodGrupo As Integer = 0, Optional ByRef pError As String = "") As Boolean

    'Acumuladores
    Function Lista_Acumuladores_Dataset(ByRef pDatos As DataSet) As Boolean
    Function Lista_Acumuladores_Fijos_Dataset(ByRef pDatos As DataSet, Optional ByVal pFavoritos As Boolean = False) As Boolean
    Function Lista_Acumuladores_Todos_Dataset(ByRef pDatos As DataSet, Optional ByVal pFavoritos As Boolean = False) As Boolean
    Function Actualiza_Acumulador_Fijo(ByVal pCod As String, Optional ByVal pFavoritos As String = "", Optional ByVal pFormato As String = "") As Boolean
    Function Actualiza_Acumulador_Def(ByVal pCod As String, ByVal pDesc_Acumulador As String, Optional ByVal pSeleccion As String = "", Optional ByVal pPeriodicidad As String = "", Optional ByVal pFormato As String = "", Optional ByVal pIntervalo As String = "", Optional ByVal pDesc_Long As String = "", Optional ByVal pFavoritos As String = "") As Boolean

    Function Consulta_Dni_Incidencias(ByRef pDatos As DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Datos_Incidencias(ByRef pDatos As DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Datos_Incidencias_DNI(ByRef pDatos As DataSet, ByVal pDNI As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Datos_Solicitudes(ByRef pDatos As DataSet, ByVal pCod_Grupo As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Datos_Solicitudes_DNI(ByRef pDatos As DataSet, ByVal pDNIs As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Dni_GruposTrabajo(ByRef pDatos As DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Grupo As String) As Boolean
    Function Consulta_Dni_GruposConsulta(ByRef pDatos As DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Grupo As String) As Boolean
    Function Consulta_Dni_Cuadrante(ByRef pDatos As DataSet, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Grupo As String) As Boolean
    Function Consulta_Datos_GruposTrabajo(ByVal Fecha As String, ByVal pDni As String) As String
    Function Consulta_Datos_GruposTrabajo2(ByRef pDatos As DataSet) As Boolean

    Function Consulta_Datos_Incidencias_Aut(ByRef pDatos As DataSet, ByVal pResponsable As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean
    Function Consulta_Datos_Solicitudes_Aut(ByRef pDatos As DataSet, ByVal pResponsable As String, ByVal Fecha_Ini As String, ByVal Fecha_Fin As String, ByVal pCod_Inc As String) As Boolean


    'para la Tabla SIGUIENTES_SOLICITUD
    Function Inserta_Siguientes_Solicitud(ByVal pcod_solicitud As String, ByVal plista_dni As String) As Boolean
    Function Elimina_Siguientes_Solicitud(Optional ByVal pcod_solicitud As String = "", Optional ByVal plista_dni As String = "") As Boolean
    Function Lista_Siguientes_Solicitud_Dataset(ByRef pDatos As DataSet, Optional ByVal pcod_solicitud As String = "", Optional ByVal plista_dni As String = "") As Boolean

    'Para promainf
    Function Lista_PromaInf(ByRef pDatos As DataSet, ByVal pConsulta As String) As Boolean

    Function Actualiza_ConsultaInforme(ByVal pSql As String, ByRef pfilas As Integer) As Boolean
    Function Inserta_ConsultaInforme(ByVal pSql As String, ByRef pfilas As Integer) As Boolean
    Function EjecutaConsultaSQL(ByVal pSQL As String, ByRef pDataSet As DataSet, Optional ByRef pError As String = "") As Boolean
    Function EjecutaComandoSQL(ByVal pSQL As String, Optional ByRef pError As String = "") As Boolean


    'Para ver el numero de justificaciones por año de una incidencia en concreto y una persona
    Function Numero_Justificaciones_Anio(ByVal pDNI As String, ByVal pAnio As String, ByVal pCod_Incidencia As String) As Integer


    'Tratamiento de incidencias.
    'Se utilizan para llamar a procedimientos almacenados que procesan lógicas complejas para:
    '- La validación de solicitudes.
    '- cálculo de máximos y mínimos de justificaciones.
    '- Validación de Justificaciones. Y para el procesamiento de inserciones necesarias en el control 
    '  de justificaciones compensables.

    Function TratamientoIncidencia(ByVal pCodInci As Integer, _
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
                                    ByRef pDescDenegacion As String) As Integer

    Function IntervalosFactorCompensacion(ByVal pCodInci As Integer, _
                                    ByVal pDni As String, _
                                    ByVal pFecha As String, _
                                    ByVal pFechaDesde As String, _
                                    ByVal pFechaHasta As String, _
                                    ByVal pHoraDesde As String, _
                                    ByVal pHoraHasta As String, _
                                    ByRef pIntervalos As String) As Integer


    'Tabla Justificaciones a Compensar
    Function Inserta_JustificacionACompensar(ByVal pCodJust As Long, _
                                    ByVal pDuracion As Integer, _
                                    ByVal pFactor As Double, _
                                    ByVal pTipoFactor As String) As Long

    Function Lista_Justificaciones_A_Compensar(ByRef pDatos As Object, _
                                    Optional ByVal pCodigoIncidencia As Integer = -1, _
                                    Optional ByVal pID_Usuario As String = Nothing, _
                                    Optional ByVal pEstado As String = Nothing _
                                    ) As Boolean
    Function Actualiza_JustificacionACompensar(ByVal pIDJust As Long, ByVal pEstado As String) As Long
    ' Tabla Compensaciones
    Function Inserta_CompensacionDeJustificacion(ByVal pCodJustCompensacion As Long, _
                                    ByVal pCodJustCompensada As Long, _
                                    ByVal pDuracion As Long) As Boolean

    Function Lista_Compensaciones(ByRef pDatos As Object, _
                                    Optional ByVal pCodigoJustificacionCompensacion As Long = -1, _
                                    Optional ByVal pIDJustificacionCompensada As Long = -1, _
                                    Optional ByVal pCodJustificacion As Long = -1) As Boolean

    'Tarjetas
    Function ListaAsociacionesTarjetasUsuario(ByRef pDatos As DataSet, Optional ByVal pID_Usuario As String = Nothing) As Boolean
    Function Elimina_Asignacion_Tarjeta_Usuario(Optional ByVal pID_Usuario As String = "", Optional ByVal pPan_Tarjeta As String = "") As Boolean
    Function Inserta_TarjetaAsociada(ByVal pDNI As String, ByVal pPan_tarjeta As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean
    Function Actualiza_TarjetaAsociada(ByVal pDNI As String, ByVal pPan_tarjeta As String, ByVal pFecha_Alta As String, Optional ByVal pFecha_Baja As String = "") As Boolean

End Interface
