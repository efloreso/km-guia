Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Public Class GuiaFiltros
#Region "Numeraciones"
    Private Enum TipoCampo
        Numerico = 10
        Alfanumerico = 1
        Entero32 = 11
        Fixed = 20
        Fecha = 30
        NoAplica = 99
    End Enum
    Public Enum TipoEstado As Byte
        Libre = 1
        EsperaAutentificacion = 2
        UsuarioConectado = 3
    End Enum
#End Region
#Region "Variables de módulo"
    Private VMblnValidaCola As Boolean                  'Si es true uno de los timers de la cola está validando y el otro debe esperar
    Private WithEvents VMtmrColaEntrada As Timers.Timer
    Private WithEvents VMtmrColaSalida As Timers.Timer
    Public VMipClientes As New Hashtable    ' detecta lo clientes conectados 
    Private WithEvents VMtmrReconecta As Timers.Timer   'En caso de que se caiga el socket lo reinicia
    Private VMobjCola As Queue              'Se almacena la cadena recibida del Servidor
    Private _Procesando As Boolean          'Variable que controla la ejecucion y creación del Hilo
    Private WithEvents VMtmrProcesa As Timers.Timer     'Envía los datos al cliente
    Private _Reconectando As Boolean                    'Está en proceso de reconectar el socket
    'Private VMobjHilo As Threading.Thread   'Hilo que ejecuta el procesamiento de la Transacción
    Public WithEvents VMtcpServidor As New Servidor
    Dim VMPosiciones As Hashtable
    Dim VMmatNivelesXTran() As NivelesTran  'Sirve para el control del filtro (si cumple o no con el filtro)
    Private VMFIltros() As Filtro           'Matriz que almacena los filtros segun el grupo
    Private VMblnUsuarioAutentificado As Boolean
    Dim VMGuiaHTTP As GuiaHTTP
#End Region
#Region "Variables utilizadas para la funcion fncRecibirRespuesta"
    '**************************************************************
    'Public VGblnFaltanDatos As Boolean
    'Public VGstrBúferGlobal As Object
    'Public VGlngNúmTotTran As Long
    'Public VGlngNúmPerdidas As Long
    'Public VGstrTransacción As String
    ''Public VGlngLongHeader As Long
    'Public VGintLongitudTransaccion As Integer
    'Public VGstrCabecera As String
    'Public VGintFechaErronea As Integer
    'Public VGblnMasInform As Boolean
    'Public VGstrBufferEnv As Object
    'Public VGstrNomServer As String
    '**************************************************************
#End Region
#Region "Propiedades privadas"
    Private _LevantandoSocket As Boolean                'True si se está llevando a cabo el proceso de listen, 
    Private _ultimoError As String
    Private _aplicacion As String                       'Nombre de la aplicacion que levantara sus cliente
    Private _servidor As String                         'Nombre de la aplicacion que levantara sus cliente
    Private _puerto As String                           'Nombre de la aplicacion que levantara sus cliente
    Private _echo As Boolean                            'Nombre de la aplicacion que levantara sus cliente
    Private _formato As String                          'Nombre de la aplicacion que levantara sus cliente
    Private _layout As String                           'Nombre de la aplicacion que levantara sus cliente
    Private _numFiltro As String                        'Nombre de la aplicacion que levantara sus cliente
    Private _BaseCentral As tipBD                       'Contiene todos los datos de la base central
    Private _Ignoradas As Integer                       'Cuenta el numero de transacciones ignoradas
    Private Shared _UsuariosConectados As Integer       'Numero de usuarios conectados
    Private _IP As String                               'La IP del cliente que esta conectado
    Private _NombrePC As String = ""                    'Nombre del equipo que esta conectado
    Private _Usuario As String = ""                     'Usuario que esta conectado
    Private _Procesados As Integer                      'Numero de bloques procesados
    Private _Recibidos As Integer                       'Numero de bloques recibidos
    Private Shared _IgnoradasTotales As Integer         'Numero de transacciones ignoradas por todos los hilos
    Private _UltimaTransaccion As String                'La fecha de la ultima transaccion recibida
    Private _TransaccionesFiltradas As Integer          'El numero de transacciones filtradas
    Private _TransaccionesTotales As Integer            'El numero de transacciones totales
    Private _Filtro As String                           'El filtro que se esta usando
    Private _FechaConexion As String                    'La fecha cuando se conecto el cliente
    Private _EstadoPuerto As String = "Libre"           'El estado en el que se encuentra el puerto de conexion con el cliente
    Private _TransaccionesEnviadas As Integer           'Numero de transacciones que se han enviado
#End Region
#Region "Propiedades públicas"
    Private _longitudEncabezado As Integer
    Public ReadOnly Property UltimoError() As String
        Get
            Return _ultimoError
        End Get
    End Property
    Public ReadOnly Property TransaccionesEnviadas() As Integer
        Get
            Return _TransaccionesEnviadas
        End Get
    End Property
    Public Property NumFiltro() As Integer
        Get
            NumFiltro = _numFiltro
        End Get
        Set(ByVal value As Integer)
            _numFiltro = value
        End Set
    End Property
    Public Property BaseCentral() As Object
        Get
            BaseCentral = _BaseCentral
        End Get
        Set(ByVal value As Object)
            _BaseCentral = value
        End Set
    End Property
    Public Property Layout() As Integer
        Get
            Layout = _layout
        End Get
        Set(ByVal value As Integer)
            _layout = value
        End Set
    End Property
    Public Property Formato() As Integer
        Get
            Formato = _formato
        End Get
        Set(ByVal value As Integer)
            _formato = value
        End Set
    End Property
    Public Property LongitudEncabezado() As Integer
        Get
            Return _longitudEncabezado
        End Get
        Set(ByVal value As Integer)
            _longitudEncabezado = value
        End Set
    End Property
    Public Property Aplicacion() As Integer
        Get
            Aplicacion = _aplicacion
        End Get
        Set(ByVal value As Integer)
            _aplicacion = value
        End Set
    End Property
    Public Property Servidor() As String
        Get
            Servidor = _servidor
        End Get
        Set(ByVal value As String)
            _servidor = value
        End Set
    End Property
    Public Property Puerto() As Integer
        Get
            Puerto = _puerto
        End Get
        Set(ByVal value As Integer)
            _puerto = value
        End Set
    End Property
    Public ReadOnly Property IP() As String
        Get
            Return _IP
        End Get
    End Property
    Public ReadOnly Property NombrePC() As String
        Get
            Return _NombrePC
        End Get
    End Property
    Public ReadOnly Property Usuario() As String
        Get
            Return _Usuario
        End Get
    End Property
    Public ReadOnly Property Procesados() As Integer
        Get
            Return _Procesados
        End Get
    End Property
    Public ReadOnly Property Memoria() As Integer
        Get
            Return VMobjCola.Count
        End Get
    End Property
    Public ReadOnly Property Recibidos() As Integer
        Get
            Return _Recibidos
        End Get
    End Property
    Public ReadOnly Property Ignoradas() As Integer
        Get
            Return _Ignoradas
        End Get
    End Property
    Public ReadOnly Property UltimaTransaccion() As String
        Get
            Return _UltimaTransaccion
        End Get
    End Property
    Public ReadOnly Property TransaccionesFiltradas() As Integer
        Get
            Return _TransaccionesFiltradas
        End Get
    End Property
    Public ReadOnly Property TransaccionesTotales() As Integer
        Get
            Return _TransaccionesTotales
        End Get
    End Property
    Public ReadOnly Property FechaConexion() As String
        Get
            Return _FechaConexion
        End Get
    End Property
    Public ReadOnly Property EstadoPuerto() As String
        Get
            Return _EstadoPuerto
        End Get
    End Property
    Public Shared ReadOnly Property UsuariosConectados() As Integer
        Get
            Return _UsuariosConectados
        End Get
    End Property
    Public Shared ReadOnly Property IgnoradasTotal() As Integer
        Get
            Return _IgnoradasTotales
        End Get
    End Property
    Public ReadOnly Property Reconectando() As Boolean
        'SEAC 26-05-10
        Get
            Return _Reconectando
        End Get
    End Property
#End Region
#Region "Propiedades Del Filtro"
    Private Structure NivelesTran
        Dim cumpleTotal As Integer
        Dim conector As String
    End Structure
    Private Structure Filtro
        Dim Cabecera As String
        Dim IdFiltro As Integer
        Dim Operador As String
        Dim Valor1 As String
        Dim Parentesis As String
        Dim Negacion As String
        Dim Valor2 As String
        Dim Conector As String
        Dim Query As String
        Dim TipoDato As TipoCampo
        Dim IdCampo As Integer
        Dim Nombre As String
        Dim Posicion As Integer
        Dim Longitud As Integer
    End Structure
#End Region
    Public Sub New(ByVal VPrincipal As GuiaHTTP)
        Try
            VMobjCola = New Queue
            VMGuiaHTTP = VPrincipal

            VMtmrReconecta = New Timers.Timer
            VMtmrReconecta.Enabled = False
            VMtmrReconecta.Interval = 100

            VMtmrColaEntrada = New Timers.Timer
            VMtmrColaEntrada.Enabled = False
            VMtmrColaEntrada.Interval = 60000

            VMtmrColaSalida = New Timers.Timer
            VMtmrColaSalida.Enabled = False
            VMtmrColaSalida.Interval = 60000

            VMtmrProcesa = New Timers.Timer
            VMtmrProcesa.Interval = 100
            VMtmrProcesa.Enabled = False
        Catch ex As Exception
            Notifica("No se pudo inicializar la clase para el envío de transacciones " & vbNewLine & ex.Message, VPrincipal, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
        End Try
    End Sub
    Public Sub RestaurarConexion()
        Try
            Dim VLstrRespuesta As String
            VMtmrReconecta.Enabled = False
            _Reconectando = True
            Notifica("Reiniciando servidor en puerto " & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
            Try
                VMtcpServidor.Cerrar()
            Catch ex As Exception

            End Try
            VMtcpServidor = New Servidor
            If LevantaServidor() Then
                _UsuariosConectados -= 1
                _EstadoPuerto = "Libre"
                _Procesados = 0
                _Ignoradas = 0
                _Recibidos = 0
                _IP = ""
                _NombrePC = ""
                _aplicacion = 0
                _Usuario = "Libre"
                _UltimaTransaccion = ""
                VMblnUsuarioAutentificado = False
                VMipClientes.Clear()
                VLstrRespuesta = STX & 8012 & FS & ACK & FS & "RC" & FS & _puerto & ETX
                VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                Notifica("Se ha restablecido el servidor en el puerto " & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
            Else
                VLstrRespuesta = STX & 8012 & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & _puerto & ETX
                VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                Notifica("No se pudo restablecer el servidor en el puerto " & _puerto & "  [" & _ultimoError & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
            End If
            _Reconectando = False
        Catch ex As Exception
            VMGuiaHTTP.VMtcpAdmin.EnviarDatos(STX & 8012 & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & _puerto & "[" & ex.Message & "]" & ETX)
            Notifica("No se pudo restablecer el servidor en el puerto " & _puerto & "  [" & ex.Message & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
        End Try
    End Sub
    Public Function LevantaServidor() As Boolean
        If CargaFiltro() Then
            If calculaposicion() Then
                Try
                    'VMtcpServidor.ConexionUnica = True
                    _LevantandoSocket = True
                    VMtcpServidor.Escuchar(_puerto)
                    _LevantandoSocket = False
                    If VMtcpServidor.Estado = KMGuiaHTTP.Servidor.enuEstado.Escuchando Then
                        Return True
                    Else
                        Return False
                    End If

                    'If VMtcpServidor.Estado = Global.KMGuiaHTTP.Servidor.sckEstado.sckError Then
                    '    If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                    '        escribeLog("Error al abrir el servidor")
                    '    ElseIf _Log = enuDepuracion.Consola Then
                    '        VMGuiaHTTP.escribeConsola("Error al abrir el servidor")
                    '    End If
                    'Else

                    'End If
                Catch ex As Exception
                    _LevantandoSocket = False
                    If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                        escribeLog("Error al abrir el servidor: " & ex.Message)
                    ElseIf _Log = enuDepuracion.Consola Then
                        VMGuiaHTTP.escribeConsola("Error al abrir el servidor: " & ex.Message)
                    End If
                    _ultimoError = "Error al abrir el servidor: " & ex.Message
                End Try
            End If
        End If
    End Function
    Private Function CargaFiltro() As Boolean

        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
            escribeLog("--> Cargando Parametros del Filtro")
        ElseIf _Log = enuDepuracion.Consola Then
            VMGuiaHTTP.escribeConsola("--> Cargando Parametros del Filtro")
        End If
        Dim VLdbsBase As Object
        Dim VLrcsFiltro As Object
        Dim VLstrSQL As String
        Dim VLcmdComando As Object
        Dim VLrcsDatos As Object
        Dim VLbytCont As Byte

        VLdbsBase = Nothing
        VLrcsDatos = Nothing
        VLcmdComando = Nothing
        VLstrSQL = "select * from tdetdconfiltros where num_filtro=" & _numFiltro

        Select Case _BaseCentral.Tipo
            Case TipoBase.MySQL
                VLdbsBase = New MySqlConnection
                VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                             ";database = " & _BaseCentral.Base & _
                                             ";uid = " & _BaseCentral.Usuario & _
                                             ";pwd = " & _BaseCentral.Contrasena
            Case TipoBase.Oracle
                VLdbsBase = New OracleConnection
                VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                             ";uid = " & _BaseCentral.Usuario & _
                             ";pwd = " & _BaseCentral.Contrasena
            Case TipoBase.SQLServer
                VLdbsBase = New SqlClient.SqlConnection
                If _BaseCentral.SI Then
                    VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                 ";Initial Catalog = " & _BaseCentral.Base & _
                                 ";Integrated Security=SSPI;"
                Else
                    VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                 ";Initial Catalog = " & _BaseCentral.Base & _
                                 ";uid = " & _BaseCentral.Usuario & _
                                 ";pwd = " & _BaseCentral.Contrasena
                End If
        End Select
        Try
            VLdbsBase.open()
            Select Case _BaseCentral.Tipo
                Case TipoBase.MySQL
                    VLcmdComando = New MySqlCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.SQLServer
                    VLcmdComando = New SqlClient.SqlCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.Oracle
                    VLcmdComando = New OracleCommand(VLstrSQL, VLdbsBase)
            End Select
            Try
                VLrcsFiltro = VLcmdComando.ExecuteReader()
                If VLrcsFiltro.hasrows Then
                    VLbytCont = 0
                    While VLrcsFiltro.Read()
                        'If VLintFilTemp <> VLrcsFiltro!IdFiltro Then
                        '    VLintFilTemp = VLrcsFiltro!IdFiltro

                        If VLbytCont = 0 Then
                            ReDim VMFIltros(0)
                        Else
                            ReDim Preserve VMFIltros(VLbytCont)
                        End If
                        VMFIltros(VLbytCont).Cabecera = Strings.Format(CInt(_layout), "00000") & Strings.Format(CInt(_formato), "000000")
                        VMFIltros(VLbytCont).Conector = VLrcsFiltro!Conector
                        VMFIltros(VLbytCont).IdCampo = VLrcsFiltro!Id_campo
                        VMFIltros(VLbytCont).IdFiltro = VLrcsFiltro!Num_Condicion
                        'If VLrcsFiltro!campoLongitud <> 0 Then
                        '    VMFIltros(VLbytCont).Longitud = VLrcsFiltro!campoLongitud
                        'Else
                        VMFIltros(VLbytCont).Longitud = VLrcsFiltro!Longitud
                        'End If
                        VMFIltros(VLbytCont).Negacion = VLrcsFiltro!Negada
                        'If VLrcsFiltro!Alias = "" Then
                        '    VMFIltros(VLbytCont).Nombre = VLrcsFiltro!NombreCampo
                        'Else
                        '    VMFIltros(VLbytCont).Nombre = VLrcsFiltro!Alias
                        'End If
                        VMFIltros(VLbytCont).Operador = VLrcsFiltro!Operador
                        VMFIltros(VLbytCont).Parentesis = VLrcsFiltro!Parentesis
                        VMFIltros(VLbytCont).Posicion = VLrcsFiltro!Posicion
                        'VMFIltros(VLbytCont).Query = VLrcsFiltro!Query

                        'Select Case Trim(VLrcsFiltro!Tipodato)
                        '    Case "ALFANUMERICO"
                        '        VMFIltros(VLbytCont).TipoDato = TipoCampo.Alfanumerico
                        '    Case "ENTERO32"
                        '        VMFIltros(VLbytCont).TipoDato = TipoCampo.Entero32
                        '    Case "FECHA"
                        '        VMFIltros(VLbytCont).TipoDato = TipoCampo.Fecha
                        '    Case "FIXED"
                        '        VMFIltros(VLbytCont).TipoDato = TipoCampo.Fixed
                        '    Case "NUMERICO"
                        '        VMFIltros(VLbytCont).TipoDato = TipoCampo.Numerico

                        'End Select

                        VMFIltros(VLbytCont).Valor1 = VLrcsFiltro!Valor
                        'VMFIltros(VLbytCont).Valor2 = VLrcsFiltro!Valor2
                        VLbytCont += 1

                        'End If
                    End While
                Else
                    VMFIltros = Nothing
                End If
                VLrcsFiltro.close()
                VLcmdComando.dispose()
                CargaFiltro = True
            Catch ex As Exception
                If _Log = enuDepuracion.Sencilla Then
                    escribeLog("  Error al cargar los Filtros")
                ElseIf _Log = enuDepuracion.Completa Then
                    escribeLog("  Error al cargar los Filtros: " & ex.Message)
                ElseIf _Log = enuDepuracion.Consola Then
                    VMGuiaHTTP.escribeConsola("  Error al cargar los Filtros: " & ex.Message)
                End If
            End Try
            VLdbsBase.close()
        Catch ex As Exception
            If _Log = enuDepuracion.Sencilla Then
                escribeLog("  Error al cargar los Filtros")
            ElseIf _Log = enuDepuracion.Completa Then
                escribeLog("  Error al cargar los Filtros: " & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                VMGuiaHTTP.escribeConsola("  Error al cargar los Filtros: " & ex.Message)
            End If
        End Try
        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
            escribeLog("<-- Finaliza Carga de Parametros del Filtro")
        ElseIf _Log = enuDepuracion.Consola Then
            VMGuiaHTTP.escribeConsola("<-- Finaliza Carga de Parametros del Filtro")
        End If
    End Function
    Public Function calculaposicion() As Boolean

        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
            escribeLog("--> Cálculo de la posicion de los campos")
        ElseIf _Log = enuDepuracion.Consola Then
            VMGuiaHTTP.escribeConsola("--> Cálculo de la posicion de los campos")
        End If

        Dim VLdbsBase As Object
        Dim VLrcsDatos As Object
        Dim VLcmdComando As Object
        Dim VLstrSQL As String
        Dim VLintI As Integer
        Dim VLstrValCampo As String
        Dim VLintRecorreMask As Integer
        Dim VLstrIDCampos() As String
        Dim VLstrRes As String = ""

        VLstrIDCampos = Nothing
        VLdbsBase = Nothing
        VLstrSQL = "SELECT * FROM tlaymsgt WHERE  Id_Formato='" & _formato & "'" & _
               " and ID_Lay='" & _layout & "'"
        Select Case _BaseCentral.Tipo
            Case TipoBase.MySQL
                VLdbsBase = New MySqlConnection
                VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                             ";database = " & _BaseCentral.Base & _
                                             ";uid = " & _BaseCentral.Usuario & _
                                             ";pwd = " & _BaseCentral.Contrasena
            Case TipoBase.Oracle
                VLdbsBase = New OracleConnection
                VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                             ";uid = " & _BaseCentral.Usuario & _
                             ";pwd = " & _BaseCentral.Contrasena
            Case TipoBase.SQLServer
                VLdbsBase = New SqlClient.SqlConnection
                If _BaseCentral.SI Then
                    VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                 ";Initial Catalog = " & _BaseCentral.Base & _
                                 ";Integrated Security=SSPI;"
                Else
                    VLdbsBase.ConnectionString = "Server = " & _BaseCentral.Servidor & _
                                 ";Initial Catalog = " & _BaseCentral.Base & _
                                 ";uid = " & _BaseCentral.Usuario & _
                                 ";pwd = " & _BaseCentral.Contrasena
                End If
        End Select



        Try
            VLdbsBase.open()
            VLcmdComando = Nothing
            Select Case _BaseCentral.Tipo
                Case TipoBase.MySQL
                    VLcmdComando = New MySqlCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.SQLServer
                    VLcmdComando = New SqlClient.SqlCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.Oracle
                    VLcmdComando = New OracleCommand(VLstrSQL, VLdbsBase)
            End Select
            VLrcsDatos = VLcmdComando.ExecuteReader

            If VLrcsDatos.HasRows Then
                While VLrcsDatos.Read
                    For VLintI = 0 To (Strings.Len(VLrcsDatos!Mascara) - 3) / 3
                        VLstrValCampo = Mid(VLrcsDatos!Mascara, VLintRecorreMask + 1, 3)
                        If VLintI = 0 Then
                            ReDim VLstrIDCampos(VLintI)
                        Else
                            ReDim Preserve VLstrIDCampos(VLintI)
                        End If
                        VLstrIDCampos(VLintI) = VLstrValCampo
                        VLintRecorreMask = VLintRecorreMask + 3
                    Next
                End While
                VLrcsDatos.Close()
            End If
            VMPosiciones = New Hashtable
            If Not VLstrIDCampos Is Nothing Then

                Dim VLintPosicion As Integer
                For VLintI = 0 To UBound(VLstrIDCampos)
                    VLstrSQL = "SELECT * FROM tlaylogc WHERE  Id_Lay='" & _layout & "'" & _
                                                     " and ID_Campo='" & VLstrIDCampos(VLintI) & "'"
                    Select Case _BaseCentral.Tipo
                        Case TipoBase.MySQL
                            VLcmdComando = New MySqlCommand(VLstrSQL, VLdbsBase)
                        Case TipoBase.SQLServer
                            VLcmdComando = New SqlClient.SqlCommand(VLstrSQL, VLdbsBase)
                        Case TipoBase.Oracle
                            VLcmdComando = New OracleCommand(VLstrSQL, VLdbsBase)
                    End Select
                    Try
                        VLrcsDatos = VLcmdComando.ExecuteReader
                        If VLrcsDatos.HasRows Then
                            While VLrcsDatos.Read
                                If VLintI = 0 Then
                                    VMPosiciones.Add(CInt(VLrcsDatos!id_campo), 1)
                                    If VLrcsDatos!camp_longitud <> 0 Then
                                        VLintPosicion = 1 + VLrcsDatos!camp_longitud
                                    Else
                                        VLintPosicion = 1 + VLrcsDatos!longitud
                                    End If
                                Else
                                    VMPosiciones.Add(CInt(VLrcsDatos!id_campo), VLintPosicion)
                                    If VLrcsDatos!camp_longitud <> 0 Then
                                        VLintPosicion = VLintPosicion + VLrcsDatos!camp_longitud
                                    Else
                                        VLintPosicion = VLintPosicion + VLrcsDatos!longitud
                                    End If
                                End If
                            End While
                        End If
                    Catch ex As MySqlException
                        VLstrRes = "PRC #02009 [" & obtieneErrorBD(ex) & "]"
                    Catch ex As SqlClient.SqlException
                        VLstrRes = "PRC #02010 [" & obtieneErrorBD(ex) & "]"
                    Catch ex As Exception
                        VLstrRes = "PRC #02011 [" & ex.Message & "]"
                    End Try
                    VLrcsDatos.Close()
                Next
            End If
            VLcmdComando.Dispose()
            VLcmdComando = Nothing
            VLdbsBase.close()
        Catch ex As MySqlException
            VLstrRes = "PRC #02009 [" & obtieneErrorBD(ex) & "]"
        Catch ex As SqlClient.SqlException
            VLstrRes = "PRC #02010 [" & obtieneErrorBD(ex) & "]"
        Catch ex As Exception
            VLstrRes = "PRC #02011 [" & ex.Message & "]"
        End Try

        Try
            If Not VMFIltros Is Nothing Then
                For VLintI = 0 To VMFIltros.Length - 1
                    If VMPosiciones.ContainsKey(VMFIltros(VLintI).IdCampo) Then
                        VMFIltros(VLintI).Posicion = VMPosiciones.Item(VMFIltros(VLintI).IdCampo)
                    End If
                Next
            End If
        Catch ex As Exception
            VLstrRes = VLstrRes & vbNewLine & "PRC #02011 [" & ex.Message & "]"
        End Try
        If VLstrRes <> "" Then
            If _Log = enuDepuracion.Sencilla Then
                escribeLog("   Error al calcular las posiciones")
            ElseIf _Log = enuDepuracion.Completa Then
                escribeLog("   Error al calcular las posiciones: " & VLstrRes)
            ElseIf _Log = enuDepuracion.Consola Then
                VMGuiaHTTP.escribeConsola("   Error al calcular las posiciones: " & VLstrRes)
            End If
        Else
            calculaposicion = True
            If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                escribeLog("   Carga de posiciones realizado con exito")
            ElseIf _Log = enuDepuracion.Consola Then
                VMGuiaHTTP.escribeConsola("   Carga de posiciones realizado con exito")
            End If
        End If


        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
            escribeLog("<-- Finaliza cálculo de la posicion de los campos")
        ElseIf _Log = enuDepuracion.Consola Then
            VMGuiaHTTP.escribeConsola("<-- Finaliza cálculo de la posicion de los campos")
        End If


    End Function
    '    Public Function fncRecibirRespuesta(ByVal PstrBúfer As String) As Boolean
    '        Dim VLintMásInfo As Integer
    '        Dim VLintTotBloque As Integer
    '        Dim VLintPos As Integer
    '        Dim VLintX As Integer
    '        Dim VLlngLongT As Long
    '        Dim VLlngLongR As Long
    '        Dim VLlngDesp As Long
    '        Dim VLstrCódigoServ As String = Nothing
    '        Dim VLstrEstacTrab As String
    '        Dim VLstrReferencia As String
    '        Dim VLstrServidor As String
    '        'Dim VLstrStratus As String
    '        Dim VLstrAux As String
    '        Dim VLstrCabecera As String
    '        Dim VLblnXSvr As Boolean
    '        Dim VLlngPosNva As Long
    '        Dim VLlngPos As Long
    '        'Dim VLlngLongitud As Long
    '        'Dim VLstrTemp As String
    '        'Dim VLstrTicket As String
    '        'Dim VLblnCoincidio As Boolean
    '        Dim VLstrEnviar As String
    '        'Dim VLstrSQL As String
    '        'Dim VLstrNoTicket As String
    '        'Dim VLrsoTicket As ADODB.Recordset
    '        'Dim VLdatFecha1 As Date
    '        'Dim VLdatFecha2 As Date
    '        'Dim VLstrTemp1

    '        VLintTotBloque = 0
    '        _Procesados += 1
    '        If (VMFIltros Is Nothing) Then
    '            fncRecibirRespuesta = True
    '            Exit Function
    '        End If

    '        Try


    '            VLblnXSvr = False
    '            VLstrEnviar = ""

    '            'Verifica que el bloque no tenga Retorno de carro o fin de página

    '            'VLstrTemp1 = PstrBúfer

    '            'Valida que no falten datos en el mensaje
    '            If VGblnFaltanDatos Then
    '                VGstrBúferGlobal = VGstrBúferGlobal & PstrBúfer
    '                If Len(VGstrBúferGlobal) >= Val(Mid(VGstrBúferGlobal, 32, 5)) Then       'Se ha completado
    '                    VGblnFaltanDatos = False
    '                    PstrBúfer = VGstrBúferGlobal
    '                    VGstrBúferGlobal = ""
    '                    VLstrServidor = Trim(Mid(PstrBúfer, 17, 8))
    '                    If Mid(VLstrServidor, 1, 8) <> "BFPROLAP" Then
    '                        VLstrCódigoServ = Val(Mid(PstrBúfer, 37, 5))
    '                        VLstrReferencia = Val(Mid(PstrBúfer, 42, 5))
    '                        VLstrEstacTrab = Trim(Mid(PstrBúfer, 67, 14))
    '                        VLblnXSvr = True
    '                    End If
    '                Else
    '                    fncRecibirRespuesta = True
    '                    Exit Function
    '                End If
    '            Else                                                                        'Todo normal
    '                'Valida el servidor que envia la respuesta
    '                VLstrServidor = Trim(Mid(PstrBúfer, 17, 8))
    '                If Mid(VLstrServidor, 1, 8) <> "BFPROLAP" And Mid(VLstrServidor, 1, 8) <> "BFPRTOTN" Then
    '                    'Respuesta de algun servidor involucrado en la carga inicial
    '                    Select Case Mid(VLstrServidor, 3, 6)                'Es un servidor válido
    '                        Case "ACTALA", "ACTUSR", "ALARMA", "APLICA", "APLGPO", "APLUSR", _
    '                             "CLVUSU", "COLDET", "COLGPO", "CONALA", "DEFLOG", "DETCON", _
    '                             "FMTGPO", "FMTUSR", "GPOCMD", "GPOLAY", "GPOUSR", "LAYDLG", _
    '                             "LAYEQC", "LAYLGC", "LAYLGT", "LAYMSG", "MTCHLG"
    '                            VLstrCódigoServ = Val(Mid(PstrBúfer, 55, 5))
    '                            VLstrReferencia = Val(Mid(PstrBúfer, 60, 5))
    '                            VLstrEstacTrab = Trim(Mid(PstrBúfer, 85, 14))
    '                            VLblnXSvr = True
    '                        Case Else                                       'Hay algún desfase en el mensaje
    '                            'Si no se trata de la solicitud de envio de las transacciones
    '                            If Not VLstrServidor = "SENDTRAN" Then
    '                                'Obtiene la posicion de la siguiente transaccion - si existe -
    '                                VLintPos = InStr(VLintPos + 1, PstrBúfer, "BFPROLAP", vbTextCompare)
    '                                'Continua con el proceso
    '                                VLstrServidor = Trim(Mid(PstrBúfer, 17, 8))

    '                            Else
    '                                fncRecibirRespuesta = True
    '                                Exit Function
    '                            End If
    '                    End Select
    '                End If
    '            End If
    '            'pendienteeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee
    '            'eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee
    '            If Val(VLstrCódigoServ) > 0 Then
    '                'DecodificarError(VLstrCódigoServ, VLstrReferencia, VLstrServidor)
    '                'EscribirError(VGstrMensajeError, "fncRecibirRespuesta")
    '                fncRecibirRespuesta = False
    '                'Terminar Val(VLstrCódigoServ)
    '                VGblnFaltanDatos = False
    '            Else
    '                If (Not Trim(Mid(PstrBúfer, 17, 8)) = "BFPROLAP" And Not Trim(Mid(PstrBúfer, 17, 8)) = "BFPRTOTN") And VLblnXSvr = False Then
    '                    VLintPos = 0
    '                    'Acumular las transacciones perdidas
    '                    VGlngNúmPerdidas = VGlngNúmPerdidas + 1
    '                    'Busca la siguiente transacción
    '                    VLintPos = InStr(VLintPos + 1, PstrBúfer, "BFPROLAP")
    '                    fncRecibirRespuesta = True
    '                    If VLintPos > 16 Then
    '                        VLstrAux = Mid(PstrBúfer, VLintPos - 16, Len(PstrBúfer))
    '                        PstrBúfer = VLstrAux
    '                    Else
    '                        Exit Function
    '                    End If
    '                    VLstrServidor = Trim(Mid(PstrBúfer, 17, 8))
    '                Else
    '                    VLintPos = 1
    '                End If

    '                If VLintPos Then
    '                    'If Val(Mid(PstrBúfer, 32, 5)) - Len(PstrBúfer) > 0 Then     'Faltan datos
    '                    If Len(PstrBúfer) < Val(Mid(PstrBúfer, 32, 5)) Then     'Faltan datos
    '                        VGblnFaltanDatos = True
    '                        VGstrBúferGlobal = PstrBúfer
    '                        fncRecibirRespuesta = True
    '                    Else
    '                        Select Case Mid(VLstrServidor, 1, 8)
    '                            Case "BFPROLAP"
    '                                VLlngLongT = Len(PstrBúfer)
    '                                fncRecibirRespuesta = True
    '                                Do
    '                                    If Mid(PstrBúfer, VLlngDesp + 17, 8) <> "BFPROLAP" Then
    '                                        VLlngPos = 0
    '                                        VLlngPos = Strings.InStr(VLlngDesp + 1, PstrBúfer, "BFPROLAP", CompareMethod.Text)
    '                                        If VLlngPos > 0 Then
    '                                            VLlngPosNva = VLlngPos - 17
    '                                            If Math.Abs(VLlngDesp - VLlngPosNva) > 2 Then
    '                                                'EscribirError("Burbuja:" & Format(Now, "YYYY/MM/DD hh:mm") & "Diferencia entre Posiciones: " & Str(VLlngDesp - VLlngPosNva), "Rec.Resp.")
    '                                            End If
    '                                            VLlngDesp = VLlngPosNva
    '                                        Else
    '                                            VGblnFaltanDatos = True
    '                                            VGstrBúferGlobal = Mid(PstrBúfer, VLlngDesp + 1)
    '                                            Exit Do
    '                                        End If
    '                                    End If
    '                                    VLlngLongR = Val(Mid(PstrBúfer, VLlngDesp + 32, 5))
    '                                    If VLlngLongR = 0 Then Exit Do
    '                                    LoadTran(Mid(PstrBúfer, VLlngDesp + 1, VLlngLongR))
    '                                    'Len(VGstrTransacción) + 54 < VLlngLongR hay que arreglar esta línea
    '                                    If Len(VGstrTransacción) < VLlngLongR Then 'Faltan datos
    '                                        VGblnFaltanDatos = True
    '                                        VGstrBúferGlobal = Mid(PstrBúfer, VLlngDesp + 1)
    '                                        Exit Do
    '                                    Else                                  'Bloque completo
    '                                        'Busca el formato correspondiente a la transaccion
    '                                        For VLintX = 1 To UBound(VMFIltros)
    '                                            VLstrCabecera = Mid(PstrBúfer, 6, 11)
    '                                            If VMFIltros(VLintX).Cabecera = VLstrCabecera Then
    '                                                'OJO, identificar automaticamente el formato de fecha de la PC, revisar WIN API, para tal efecto
    '                                                'Valida que la posicion 25 al 31 del buffer sea una fecha
    '                                                'If IsDate(Format(Mid(PstrBúfer, 25, 6), "YY/MM/DD")) Then
    '                                                '    VGfmtFormato(VLintX).Fecha_Arch = Mid(PstrBúfer, 25, 6)
    '                                                'Else
    '                                                '    If VGintFechaErronea > 32500 Then VGintFechaErronea = 1
    '                                                '    VGintFechaErronea = VGintFechaErronea + 1
    '                                                '    VGfmtFormato(VLintX).Fecha_Arch = "000000"

    '                                                'End If

    '                                                'Pasa la transaccion del mensaje a la variable de trabajo global
    '                                                If fncCargarTransaccion(Mid(PstrBúfer, VLlngDesp + 1, VLlngLongR), VLlngLongR) = True Then
    '                                                    'Variables para la carga de las Tablas Dinámicas TT
    '                                                    VGstrTransacción = Mid(PstrBúfer, VLlngDesp + _longitudEncabezado + 1, VLlngLongR - _longitudEncabezado)
    '                                                    If Not VGstrTransacción.Contains("BFPROLAP") Then
    '                                                        _TransaccionesTotales += 1
    '                                                        _UltimaTransaccion = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
    '                                                        If aplicaFiltro(VGstrTransacción, VMFIltros) Then
    '                                                            VGstrTransacción = Mid(PstrBúfer, VLlngDesp + 1, VLlngLongR)
    '                                                            VLstrEnviar = VLstrEnviar & VGstrTransacción
    '                                                            _TransaccionesFiltradas += 1
    '                                                            VLintTotBloque += 1
    '                                                        End If
    '                                                    Else
    '                                                        VLlngDesp += (VGstrTransacción.Substring(VGstrTransacción.IndexOf("BFPROLAP", CInt(VLlngDesp + _longitudEncabezado - 16))))
    '                                                    End If
    '                                                    'VGstrTransacción = Mid(VGstrTransacción, _longitudEncabezado + 1)

    '                                                    'VLintXTc = VLintX
    '                                                Else
    '                                                    Exit Function
    '                                                End If

    '                                                'Sale del ciclo de formatos
    '                                                Exit For
    '                                            End If
    '                                        Next VLintX
    '                                        'Siguiente Transacciòn
    'Siguiente:
    '                                        VLlngDesp = VLlngDesp + VLlngLongR
    '                                    End If
    '                                Loop Until VLlngDesp >= VLlngLongT
    '                                fncRecibirRespuesta = True
    '                        End Select

    '                        VGblnMasInform = VLintMásInfo
    '                        VGstrBufferEnv = PstrBúfer
    '                        VGstrNomServer = VLstrServidor
    '                    End If
    '                End If
    '            End If
    '            If VLstrEnviar <> "" Then
    '                If VMblnUsuarioAutentificado Then
    '                    Try
    '                        VMtcpServidor.EnviarDatos(VMipClientes(1), VLstrEnviar)
    '                        _TransaccionesEnviadas += 1
    '                    Catch ex As Exception
    '                        _Ignoradas += 1
    '                        _IgnoradasTotales += 1
    '                    End Try
    '                Else
    '                    _Ignoradas += 1
    '                    _IgnoradasTotales += 1
    '                End If
    '            End If
    '        Catch ex As Exception
    '            Dim i
    '            i = 1
    '        End Try
    '        'fncCargarTransaccion(PstrBúfer, PstrBúfer.Length)
    '    End Function
    'Public Function LoadTran(ByVal VPstrMensaje As String) As Boolean
    '    Try
    '        Dim VLstrMSGErr As String = ""
    '        VPstrMensaje = Replace(VPstrMensaje, "'", " ")
    '        VGstrTransacción = VPstrMensaje ' Mid(VPstrMensaje, _longitudEncabezado + 1)
    '        LoadTran = True
    '    Catch ex As Exception

    '    End Try
    'End Function
    Public Sub Enviar(ByVal Datos As String)
        '*****************************************************************************
        'Proceso de paso para recibir los datos y enviarlos a procesar
        'Parámetros:
        '           Datos=Contiene la transacción
        '*****************************************************************************
        Try
            _Recibidos += 1
            SyncLock VMobjCola
                VMobjCola.Enqueue(Datos)
            End SyncLock
            VMtmrColaEntrada.Enabled = False
            VMtmrColaEntrada.Enabled = True
            VMtmrColaSalida.Enabled = True
            If Not _Procesando Then
                VMtmrProcesa.Enabled = True
            End If
        Catch ex As Exception
            Notifica("No se pudo cargar el mensaje en la cola de salida en el puerto " & _puerto & vbNewLine & ex.Message, VMGuiaHTTP, enuEscribeLog.Completo, enuTipoAviso.MensajeError)
            _Ignoradas += 1
        End Try
    End Sub
    'Private Sub procesaDatos()
    '    Dim VLstrMensaje As String = ""
    '    '_Procesando = True
    '    Try

    '        SyncLock VMobjCola
    '            While VMobjCola.Count > 0
    '                VLstrMensaje = VMobjCola.Dequeue
    '                procesaMensaje(VLstrMensaje)
    '                'fncRecibirRespuesta(VLstrMensaje)
    '            End While
    '        End SyncLock
    '    Catch ex As Exception
    '        If _Log = enuDepuracion.Completa Then
    '            escribeLog(" No se pudo enviar el mensaje:" & VLstrMensaje & vbNewLine & " Error:" & ex.Message)
    '        ElseIf _Log = enuDepuracion.Consola Then
    '            VMGuiaHTTP.escribeConsola(" No se pudo enviar el mensaje:" & VLstrMensaje & vbNewLine & " Error:" & ex.Message)
    '        End If
    '        _Ignoradas += 1
    '    End Try

    '    '_Procesados = False
    'End Sub
    Private Sub procesaMensaje(ByVal VPstrMen As String)
        Static VLstrLinea As String
        Dim VlintLongitud As Integer
        Dim VLstrMensaje As String
        Dim VLstrMensajetemp As String

        Try

            VLstrLinea &= VPstrMen

            If VLstrLinea.Length > 0 Then
                While VLstrLinea.Contains("BFPROLAP")
                    If VLstrLinea.Length > (VLstrLinea.IndexOf("BFPROLAP") + 20) Then 'Verfica si la longitud viene en la cadena
                        VlintLongitud = VLstrLinea.Substring(VLstrLinea.IndexOf("BFPROLAP") + 15, 5)
                        If VLstrLinea.Length >= VlintLongitud + (VLstrLinea.IndexOf("BFPROLAP") + 20) Then
                            'If Not VLstrLinea.Substring(VLstrLinea.IndexOf("BFPROLAP") + 7, VlintLongitud - (VLstrLinea.IndexOf("BFPROLAP") + 7)).Contains("BFPROLAP") Then
                            If Not VLstrLinea.Substring(VLstrLinea.IndexOf("BFPROLAP") + 7, VlintLongitud).Contains("BFPROLAP") Then
                                VLstrMensaje = VLstrLinea.Substring((VLstrLinea.IndexOf("BFPROLAP") - 16 + _longitudEncabezado), VlintLongitud - _longitudEncabezado)
                                VLstrMensaje = Replace(VLstrMensaje, "'", " ")
                                VLstrMensajetemp = VLstrLinea.Substring((VLstrLinea.IndexOf("BFPROLAP") - 16), VlintLongitud)
                                If aplicaFiltro(VLstrMensaje, VMFIltros) Then
                                    If VMblnUsuarioAutentificado Then
                                        Try
                                            VMtcpServidor.EnviarDatos(VMipClientes(1), VLstrMensajetemp)
                                            If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                                                escribeLog(VLstrMensajetemp)
                                            End If
                                            _TransaccionesEnviadas += 1
                                        Catch ex As Exception
                                            _Ignoradas += 1
                                            _IgnoradasTotales += 1
                                        End Try
                                    Else
                                        _Ignoradas += 1
                                        _IgnoradasTotales += 1
                                    End If
                                    _TransaccionesFiltradas += 1
                                End If
                                VLstrLinea = VLstrLinea.Substring(VlintLongitud)
                                'VLstrLinea = VLstrLinea.Substring(VLstrLinea.IndexOf("BFPROLAP", 25) - 16)
                            Else
                                VLstrLinea = VLstrLinea.Substring(VLstrLinea.IndexOf("BFPROLAP", 25) - 16)
                            End If
                        Else
                            Exit While
                        End If
                    Else
                        Exit While
                    End If
                End While
            End If
        Catch ex As Exception
            If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                escribeLog("  Mensaje==" & ex.Message & vbNewLine & STX & encriptaValor(VLstrLinea) & ETX)
            ElseIf _Log = enuDepuracion.Consola Then
                VMGuiaHTTP.escribeConsola("  Mensaje==" & ex.Message & vbNewLine & encriptaValor(VLstrLinea))
            End If

            escribeVisorEventos("Error ==" & ex.Message & vbNewLine & STX & encriptaValor(VLstrLinea) & ETX, EventLogEntryType.Warning)
        End Try
    End Sub
    'Private Function fncCargarTransaccion(ByVal Pstrmensaje As String, ByVal PlngLongitud As Long) As Boolean
    '    'Dim VLstrDatos As String
    '    'Dim objfiltro() As Filtro
    '    'Dim VLinti As Integer

    '    'objfiltro = VMFIltros
    '    'If (objfiltro Is Nothing) Then
    '    '    fncCargarTransaccion = True
    '    '    Exit Function
    '    'End If
    '    'VLstrDatos = mensaje
    '    'VLstrDatos = Strings.Right(VLstrDatos, VLstrDatos.Length - 54)
    '    'If coincide(VLstrDatos, objfiltro) = False Then
    '    '    If VMipClientes.Count > 0 Then
    '    '        VMtcpServidor.EnviarDatos(VMipClientes(1), mensaje)
    '    '        For VLinti = 2 To VMipClientes.Count
    '    '            VMtcpServidor.EnviarDatos(VMipClientes(VLinti), "Solo se admite un cliente por servidor")
    '    '            VMipClientes.Remove(VLinti)
    '    '        Next
    '    '    End If
    '    'End If

    '    'Se sustituyen todos los ' x espacios
    '    Pstrmensaje = Replace(Pstrmensaje, "'", " ")

    '    'Valida que coincidan la longitud de la transaccion en el buffer con la marcada por la transaccion
    '    If PlngLongitud = Len(Pstrmensaje) And PlngLongitud >= (VGintLongitudTransaccion + _longitudEncabezado) Then
    '        'En caso de que el acumulador de transacciones sea mayor que un double se inicializa
    '        If VGlngNúmTotTran > 2147483600 Then VGlngNúmTotTran = 1

    '        'Se acumulan las transacciones
    '        VGlngNúmTotTran = VGlngNúmTotTran + 1

    '        'Se pasa la transaccion del mensaje a la variable global de trabajo
    '        'Se restan 36 por ser parte del header de comunicaciones
    '        'VGstrTransacción = Mid(PstrMensaje, 54, Val(Mid(PstrMensaje, 32, 5)) - 54)

    '        fncCargarTransaccion = True
    '    Else
    '        'Si no coincide las longitudes, entonces se almacena el mensaje en un arreglo
    '        'Incrementa el indice de transacciones incompletas - debe de ser igual al de completas al final del proceso -
    '        'VGintTransaccionesIncompletas = VGintTransaccionesIncompletas + 1
    '        'Limpia el arreglo en la posicion que marque el indice
    '        'ReDim VGstrTranIncompletas(1)
    '        'Mueve el mensaje incompleto de la transaccion al arreglo
    '        'VGstrTranIncompletas(1) = PstrMensaje

    '        fncCargarTransaccion = False
    '    End If

    '    Exit Function
    '    'fncCargarTransaccion = True
    'End Function
    Private Function aplicaFiltro(ByVal VPstrDatos As String, ByVal objFiltro() As Filtro) As Boolean
        '******************************************************************************
        'Se realizan las comparaciones para determinar si se cumplen con lo especificado 
        'en el filtro
        'Parámetros:
        '           VPstrDatos=Trama copn todos los datos(Transacción)
        '           objFiltro()=Contiene la información del Filtro
        'Valores devueltos:
        '              True='cumple con todas las condiciones
        '              False='NO cumple con todas las condiciones 
        '******************************************************************************
        Dim VLintCumple As Integer, VLintIdx As Integer, VLintIdy As Integer
        Dim VLintNivel As Integer
        Dim VLstrValor As String
        Dim VLblnParentesis() As Boolean

        ReDim VMmatNivelesXTran(0)
        ReDim VLblnParentesis(0)

        For VLintIdx = 0 To objFiltro.Length - 1
            With objFiltro(VLintIdx)
                Select Case Trim(.Parentesis)
                    Case 0 'Evalua la condición
                        If .Posicion > 0 Then
                            VLstrValor = Trim(VPstrDatos.Substring(.Posicion - 1, .Longitud))
                        Else
                            VLstrValor = Trim(VPstrDatos.Substring(.Posicion, .Longitud))
                        End If
                        VLintIdy = VLintIdy + 1
                        VLintCumple = 0

                        Select Case Trim(.Operador)
                            Case 0 'Igual a
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) = Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) = Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case 1 'Distinto de
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) <> Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) <> Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case 2 'Mayor que
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) > Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) > Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case 3 'Mayor o igual a
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) >= Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) >= Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case 4 'Menor que
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) < Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) < Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case 5 'Menor o igual a
                                If IsNumeric(Trim(.Valor1)) Then
                                    If Val(Trim(VLstrValor)) <= Val(Trim(.Valor1)) Then VLintCumple = 1
                                Else
                                    If Trim(VLstrValor) <= Trim(.Valor1) Then VLintCumple = 1
                                End If
                            Case "like" 'Como
                                ' If Trim(VLstrValor) Like .Valor & "*" Then VLintCumple = 1
                                If Trim(VLstrValor) Like .Valor1 Then VLintCumple = 1
                        End Select

                        If VLintIdy = 1 Then
                            If .Negacion Then
                                VMmatNivelesXTran(VLintNivel).cumpleTotal = Math.Abs(CInt(Not CBool(VLintCumple)))
                            Else
                                VMmatNivelesXTran(VLintNivel).cumpleTotal = Math.Abs(CInt(CBool(VLintCumple)))
                            End If
                        Else
                            If .Conector = 1 Then 'Y
                                If .Negacion Then
                                    If VLintCumple = 0 Then
                                        If VMmatNivelesXTran(VLintNivel).cumpleTotal = 0 Then
                                            VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                                        Else
                                            VMmatNivelesXTran(VLintNivel).cumpleTotal = 1
                                        End If
                                    Else
                                        If VMmatNivelesXTran(VLintNivel).cumpleTotal = 0 Then
                                            VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                                        Else
                                            ''''JK
                                            VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                                        End If
                                    End If
                                Else
                                    If VLintCumple > 0 And VMmatNivelesXTran(VLintNivel).cumpleTotal <= 0 Then
                                        VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                                    ElseIf VLintCumple = 0 Then
                                        VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                                    End If
                                End If
                            ElseIf .Conector = 0 Then  'O
                                If VMmatNivelesXTran(VLintNivel).cumpleTotal <= 0 Then
                                    If .Negacion Then
                                        If VLintCumple = 0 Then VMmatNivelesXTran(VLintNivel).cumpleTotal = 1
                                    Else
                                        If VLintCumple > 0 Then VMmatNivelesXTran(VLintNivel).cumpleTotal = 1
                                    End If
                                End If
                            End If
                        End If
                    Case 1 ' Abrir Parentesis
                        VLintNivel = VLintNivel + 1
                        ReDim Preserve VMmatNivelesXTran(VLintNivel)
                        VMmatNivelesXTran(VLintNivel).conector = .Conector
                        VLintIdy = 0
                        ReDim Preserve VLblnParentesis(VLblnParentesis.Length + 1)
                        If .Negacion Then
                            VLblnParentesis(UBound(VLblnParentesis)) = True
                            'Else
                            '   VLblnParentesis(UBound(VLblnParentesis)) = False
                        End If

                    Case 2 ' Cierra Parentesis
                        If VLblnParentesis(VLblnParentesis.Length - 1) Then
                            If (VMmatNivelesXTran(VLintNivel).cumpleTotal) > 0 Then
                                VMmatNivelesXTran(VLintNivel).cumpleTotal = 0
                            Else
                                VMmatNivelesXTran(VLintNivel).cumpleTotal = 1
                            End If
                        End If
                        ReDim Preserve VLblnParentesis(VLblnParentesis.Length - 1)
                        Select Case Trim(VMmatNivelesXTran(VLintNivel).conector)
                            Case -1 'Ninguno
                                VMmatNivelesXTran(VLintNivel - 1).cumpleTotal = VMmatNivelesXTran(VLintNivel).cumpleTotal
                            Case 1  'Y
                                VMmatNivelesXTran(VLintNivel - 1).cumpleTotal = VMmatNivelesXTran(VLintNivel - 1).cumpleTotal * VMmatNivelesXTran(VLintNivel).cumpleTotal
                            Case 0  'O
                                VMmatNivelesXTran(VLintNivel - 1).cumpleTotal = VMmatNivelesXTran(VLintNivel - 1).cumpleTotal + VMmatNivelesXTran(VLintNivel).cumpleTotal
                        End Select
                        VLintNivel = VLintNivel - 1
                        ReDim Preserve VMmatNivelesXTran(VLintNivel)
                End Select
            End With
        Next

        If VMmatNivelesXTran(VLintNivel).cumpleTotal > 0 Then
            aplicaFiltro = True
        Else
            aplicaFiltro = False
        End If

    End Function
    Private Sub VMtcpServidor_DatosRecibidos(ByVal IDTerminal As System.Net.IPEndPoint, ByVal Datos As String) Handles VMtcpServidor.DatosRecibidos
        If VMipClientes.ContainsValue(IDTerminal) Then
            If Datos.Contains("SENDTRAN") Then
                _NombrePC = Datos.Substring(51, 14)
                _Usuario = Datos.Substring(65, 5)
                _aplicacion = Datos.Substring(71, 2)
                _FechaConexion = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
                _EstadoPuerto = "Autentificado"
                VMtcpServidor.EnviarDatos(IDTerminal, Datos)
                Notifica("Se ha recibido la solicitud de envío de información en el puerto " & _puerto & " con los siguientes datos" & vbNewLine & "IP: " & IDTerminal.ToString & vbNewLine & "PC: " & _NombrePC & vbNewLine & "Usuario: " & _Usuario & "[" & _aplicacion & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
                VMblnUsuarioAutentificado = True
                VMtmrColaEntrada.Enabled = True
                'End If
            End If
        End If
        VMblnUsuarioAutentificado = True
    End Sub
    Private Sub VMtcpServidor_NuevaConexion(ByVal IDTerminal As System.Net.IPEndPoint) Handles VMtcpServidor.NuevaConexion
        If VMipClientes.Count = 0 Then
            VMipClientes.Add(1, IDTerminal)
            _UsuariosConectados += 1
Acepta:
            _IP = IDTerminal.ToString
            _Usuario = "Sin autentificar"
            _EstadoPuerto = "Sin autentificar"
        Else
            VMipClientes.Clear()
            _UsuariosConectados = 1
            VMipClientes.Add(1, IDTerminal)
            GoTo Acepta
            'If Clientes > 1 Then
            '    VMtcpServidor.EnviarDatos(IDTerminal, "Puerto Ocupado")
            '    Threading.Thread.Sleep(500)
            '    VMtcpServidor.CerrarCliente(IDTerminal)
            'Else
            '    GoTo Acepta
            'End If
        End If
    End Sub
    Private Sub VMtcpServidor_ConexionTerminada(ByVal IDTerminal As System.Net.IPEndPoint) Handles VMtcpServidor.ConexionTerminada
        If VMipClientes.ContainsValue(IDTerminal) Then
            VMipClientes.Clear()
            _UsuariosConectados -= 1
            _EstadoPuerto = "Libre"
            _Procesados = 0
            _Ignoradas = 0
            _Recibidos = 0
            _IP = ""
            _NombrePC = ""
            _aplicacion = 0
            _Usuario = "Libre"
            _UltimaTransaccion = ""
            VMblnUsuarioAutentificado = False
        End If
    End Sub
    Private Sub VMtcpServidor_ErrorServidor(ByVal Mensaje As String) Handles VMtcpServidor.ErrorServidor
        If Not _LevantandoSocket Then
            Notifica("Error en el socket " & Mensaje & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
            If Not _Reconectando Then
                VMtmrReconecta.Enabled = True
            End If
        End If
        _ultimoError = Mensaje
    End Sub
    Private Sub VMtmrReconecta_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrReconecta.Elapsed
        Dim VLstrRespuesta As String

        VMtmrReconecta.Enabled = False
        _Reconectando = True
        Notifica("Reiniciando servidor en puerto " & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
        VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "Reconectando servidor en puerto " & _puerto & ETX
        VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
        Try
            VMtcpServidor.Cerrar()
        Catch ex As Exception

        End Try

        VMtcpServidor = New Servidor
        If LevantaServidor() Then
            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "00" & FS & "Se ha restablecido el servidor en el puerto " & _puerto & ETX
            VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
            Notifica("Se ha restablecido el servidor en el puerto " & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
        Else
            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se pudo restablecer el servidor en el puerto " & _puerto & ". Se reintentará en un minuto" & ETX
            VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
            Notifica("No se pudo restablecer el servidor en el puerto " & _puerto & "  [" & _ultimoError & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
            VMtmrReconecta.Interval = 60000
            VMtmrReconecta.Enabled = True
        End If
        _Reconectando = False
    End Sub
    Private Sub VMtmrProcesa_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrProcesa.Elapsed
        Dim VLstrMensaje As String

        VMtmrProcesa.Enabled = False
        _Procesando = True
        Try
            While VMobjCola.Count > 0
                SyncLock VMobjCola
                    VLstrMensaje = VMobjCola.Dequeue
                End SyncLock
                VMtmrColaSalida.Enabled = False
                procesaMensaje(VLstrMensaje)
                VMtmrColaSalida.Enabled = True
            End While

        Catch ex As Exception
            If _Log = enuDepuracion.Completa Then
                escribeLog(" No se pudo enviar el mensaje [" & _puerto & "]" & vbNewLine & " Error:" & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                VMGuiaHTTP.escribeConsola(" No se pudo enviar el mensaje [" & _puerto & "]" & vbNewLine & " Error:" & ex.Message)
            End If
            _Ignoradas += 1
        End Try

        _Procesando = False
    End Sub
    Private Sub VMtmrColaEntrada_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrColaEntrada.Elapsed
        If Not VMblnValidaCola Then
            VMblnValidaCola = True
            VMtmrColaEntrada.Enabled = False
            VMGuiaHTTP.validaEnvio()
            VMtmrColaEntrada.Enabled = True
            VMblnValidaCola = False
        End If
    End Sub
    Private Sub VMtmrColaSalida_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrColaSalida.Elapsed
        If Not VMblnValidaCola Then
            VMblnValidaCola = True
            Try
                VMtmrColaSalida.Enabled = False

                'Ha ocurrido un minuto sin enviar información
                Notifica("Validando estado de envío [" & _puerto & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)

                If Not VMblnValidaCola Then
                    'VMblnValidaCola = True
                    'Verifica que la cola no tenga datos
                    If VMobjCola.Count > 0 And Not _Procesados Then
                        VMtmrProcesa.Enabled = True
                    Else
                        VMtmrColaEntrada.Enabled = False
                        VMGuiaHTTP.validaEnvio()
                    End If
                    'Else
                    '   VMblnEsperaCola = True
                End If

                VMtmrColaSalida.Enabled = True
            Catch ex As Exception
                VMtmrColaSalida.Enabled = True
            End Try
            VMblnValidaCola = False
        End If
    End Sub
End Class
