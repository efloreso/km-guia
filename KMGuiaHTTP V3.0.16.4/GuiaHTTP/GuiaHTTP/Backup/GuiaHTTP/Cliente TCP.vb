Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class Cliente_TCP
    '*******************************************************************************************************
    'FECHA DE MODIFICACIÓN: 23/ENERO/2009 
    'NOMBRE: SANDRA APARICIO
    'MODIFICACIONES:
    '1. Se agrego el Metodo EnviarArchivo, que permite enviar la información en bytes al servidor.
    'FECHA DE MODIFICACIÓN 16/FEBRERO/2009
    '1. Modificaciones en la codificación de los bytes enviados y recibidos. Se cambio de condificación ASCII, a la codificación por Default
    'FECHA DE MODIFICACIÓN 26/MAYO/2009
    '1. Modificaciones en la codificación de los bytes enviados y recibidos. Se cambio de condificación por Default, a codificación UTF8
    '2. Se agrego la función obtieneError para catalogar los errores que se pueden presentar en el Socket.
    '3. Se agrego la propiedad Retardo, la cual pone la propiedad .NoDelay de la variable TcpCliente en:, por default tiene un valor FALSE. Es 
    ' TRUE: Cuando se require que el envio de los datos sea de forma inmediata. 
    '4. El evento ErrorConexion ya no se genera unicamente al momento de realizar la conexión, ahora tambien se dispara al momento de generarse
    '  un error en el Socket al recibir o enviar los datos.
    '*******************************************************************************************************
#Region "Variables"
    Private VMstmDatos As Stream        'Utilizado para enviar datos al Servidor y recibir datos del mismo
    Private VMtcpCliente As TcpClient   'Permite la conexion con el servidor 

    Private VMhiloConexion As Thread    'Permite conectarme con el servidor en segundo plano
    Private VMhiloLeer As Thread        'Permite la lectura de datos del servidor en segundo plano

    Private _Estado As EstadoConexion   'Me indica el estado que se encuntra el cliente

    Private VMintEcosSinRespuesta As Integer    'Contiene el numero de Ecos sin responder

    Private _ActivoEco As Boolean       'Permite mandar los Ecos al administrador
    Private _Direccion As String        'Direccion del objeto de la clase Servidor
    Private _MensajeEco As String       'Es el mensaje que se le manda al servidor por el Eco

    Private _TiempoEco As Integer       'Es el tiempo para saber cuando hay que mandar un Eco en segundos
    Private _DesconexionEco As Integer  'Es el numero de Eco que no responde el Servidor para hacer la desconexion
    Private _Puerto As Integer          'Puerto donde escucha el objeto de la clase Servidor

    Private _Retardo As Boolean         'Indica si la propiedad NoDelay sera activada o no. 26-Mayo-09 Sandra Aparicio

    Private WithEvents VMtmrEco As Timers.Timer     'Permite mandar Ecos cuando no se recibe datos del servidor
#End Region
#Region "Estructuras y Enumeraciones"
    Public Enum EstadoConexion As Byte
        ''' <summary>
        ''' No se encuentra conectado con el servidor
        ''' </summary>
        Desconectado = 0
        ''' <summary>
        ''' Se encuentra conectado con el servidor
        ''' </summary>
        Conectado = 1
        ''' <summary>
        ''' Se esta conectando con el servidor
        ''' </summary>
        Conectando = 2
        ''' <summary>
        ''' Se esta desconectando del servidor
        ''' </summary>
        Desconectando = 3
        ''' <summary>
        ''' Esta determinando la direccion del servidor
        ''' </summary>
        EnError = 4
        ResolviendoHost = 5
    End Enum
#End Region
#Region "Eventos"
    ''' <summary>
    ''' Este evento se dispara cuando la clase cliente se desconecta con el servidor
    ''' </summary>
    Public Event ConexionTerminada(ByVal DesconexionEco As Boolean)
    ''' <summary>
    ''' Este evento se dispara cuando se reciben datos del servidor
    ''' </summary>
    Public Event DatosRecibidos(ByVal Datos As String)
    ''' <summary>
    ''' Este evento se dispara cuando existe un error en la conexion con el servidor
    ''' </summary>
    Public Event ErrorConexion(ByVal Mensaje As String)
    ''' <summary>
    ''' Este evento se dispara cuando se encuentra conectado con el servidor
    ''' </summary>
    Public Event ConexionEstablecida()
#End Region
#Region "Propiedades"
    ''' <summary>
    ''' Es el numero de ecos que se deben de mandar para realizar la desconexion con el servidor
    ''' </summary>
    Public Property DesconexionEco() As Integer
        Get
            Return _DesconexionEco
        End Get
        Set(ByVal value As Integer)
            _DesconexionEco = value
        End Set
    End Property
    ''' <summary>
    ''' Es el tiempo en segundos que va a esperar la clase cliente para mandar un Eco
    ''' </summary>
    Public Property TiempoEco() As Integer
        Get
            Return _TiempoEco
        End Get
        Set(ByVal value As Integer)
            If value <= 60 Then
                _TiempoEco = value
            Else
                Throw New System.Exception("El tiempo para mandar el Eco no puede ser mayor a 60 segundos")
            End If
        End Set
    End Property
    ''' <summary>
    ''' El mensaje que se va a mandar al servidor en caso de que se tenga que mandar un Eco
    ''' </summary>
    Public Property MensajeEco() As String
        Get
            Return _MensajeEco
        End Get
        Set(ByVal value As String)
            _MensajeEco = value
        End Set
    End Property
    ''' <summary>
    ''' Indica si se van a mandar ecos al servidor (valor por default es FALSE)
    ''' </summary>
    Public Property ActivarEco() As Boolean
        Get
            Return _ActivoEco
        End Get
        Set(ByVal value As Boolean)
            _ActivoEco = value
        End Set
    End Property
    ''' <summary>
    ''' Indica si me encuentro conectado con el servidor
    ''' </summary>
    Public ReadOnly Property Conectado() As Boolean
        Get
            Try
                Return VMtcpCliente.Connected
            Catch ex As Exception

            End Try
        End Get
    End Property

    ''' <summary>
    ''' La direccion IP donde se encuentra el servidor
    ''' </summary>
    Public Property DireccionIP() As String
        Get
            DireccionIP = _Direccion
        End Get
        Set(ByVal Value As String)
            _Direccion = Value
        End Set
    End Property

    ''' <summary>
    ''' El numero de puerto que se necesita para conectarse con el servidor
    ''' </summary>
    Public Property Puerto() As Integer
        Get
            Puerto = _Puerto
        End Get
        Set(ByVal Value As Integer)
            _Puerto = Value
        End Set
    End Property

    ''' <summary>
    ''' El estado en el que se encuentra la clase cliente
    ''' </summary>
    Public ReadOnly Property Estado() As EstadoConexion
        Get
            Return _Estado
        End Get
    End Property

    ''' <summary>
    ''' Indica si existira retardo al enviar datos, Valor por defaul FALSE
    ''' </summary>            
    Public Property Retardo() As Boolean
        'Agregado 26-Mayo-09 
        'Sandra Aparicio
        Get
            Return _Retardo
        End Get
        Set(ByVal value As Boolean)
            _Retardo = value
        End Set

    End Property
#End Region
#Region "Metodos"
    ''' <summary>
    ''' Instancia una nueva clase cliente
    ''' </summary>
    Public Sub New()
        VMhiloConexion = New Thread(AddressOf ConectarServidor)
        VMhiloLeer = New Thread(AddressOf LeerSocket)
        VMtcpCliente = New TcpClient()
        VMtmrEco = New Timers.Timer
        _ActivoEco = False
        _DesconexionEco = 2
        _Estado = EstadoConexion.Desconectado
        _MensajeEco = ChrW(2) & "9000" & ChrW(3)
        _TiempoEco = 60
        _Retardo = False '26-Mayo-09 Sandra Aparicio
    End Sub

    ''' <summary>
    ''' Se conecta con el servidor
    ''' </summary>
    Public Sub Conectar()
        NuevaConexion()
    End Sub

    ''' <summary>
    ''' Se conecta con el servidor
    ''' </summary>
    ''' <param name="Servidor">La direccion IP del servidor</param>
    Public Sub Conectar(ByVal Servidor As String)
        _Direccion = Servidor
        NuevaConexion()
    End Sub

    ''' <summary>
    ''' Se conecta con el servidor
    ''' </summary>
    ''' <param name="Servidor">La direccion IP del servidor</param>
    ''' <param name="Puerto">El puerto de conexion con el servidor</param>
    Public Sub Conectar(ByVal Servidor As String, ByVal Puerto As Integer)
        _Direccion = Servidor
        _Puerto = Puerto
        NuevaConexion()
    End Sub

    ''' <summary>
    ''' Permite enviar mensajes al servidor
    ''' </summary>
    ''' <param name="Datos">El mensaje que se le envia la servidor</param>
    Public Sub EnviarDatos(ByVal Datos As String)
        Dim VLobjCodificacion As New UTF8Encoding
        Try
            If Not VMtcpCliente.Connected Then
                Throw New System.Exception("No se encuentra conectado al servidor " & Me.DireccionIP & ":" & Me.Puerto)
            Else
                Dim VLbytEscritura() As Byte
                '------------------------------------------------------
                VLbytEscritura = VLobjCodificacion.GetBytes(Datos)
                'VLbytEscritura = Encoding.Default.GetBytes(Datos)
                '------------------------------------------------------
                If Not (VMstmDatos Is Nothing) Then
                    'Envio los datos al Servidor
                    VMstmDatos.Write(VLbytEscritura, 0, VLbytEscritura.Length)
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("Error al mandar los datos al servidor por el siguiente motivo: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Permite enviar mensajes al servidor en formato de bytes
    ''' </summary>
    ''' <param name="Datos">El mensaje en bytes que se le envia la servidor</param>
    Public Sub EnviarArchivo(ByVal Datos() As Byte)
        Try
            If Not VMtcpCliente.Connected Then
                Throw New System.Exception("No se encuentra conectado al servidor " & Me.DireccionIP & ":" & Me.Puerto)
            Else

                If Not (VMstmDatos Is Nothing) Then
                    'Envio los datos al Servidor
                    VMstmDatos.Write(Datos, 0, Datos.Length)
                End If
            End If
        Catch ex As Exception
            Throw New System.Exception("Error al mandar los datos al servidor por el siguiente motivo: " & ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Cierra la conexion con el servidor
    ''' </summary>
    Public Sub Cerrar()

        _Estado = EstadoConexion.Desconectando

        Try
            If VMhiloLeer.IsAlive Then
                VMhiloLeer.Abort()
            End If
        Catch ex As Exception

        End Try
        Try
            If VMtcpCliente.Connected Then
                VMtcpCliente.Close()
            End If
        Catch ex As Exception

        End Try
        Try
            If VMhiloConexion.IsAlive Then
                VMhiloConexion.Abort()
            End If
        Catch ex As Exception

        End Try
        Dispose()
        'VMstmDatos.Close()
        _Estado = EstadoConexion.Desconectado


    End Sub

    ''' <summary>
    ''' Libera todos los recurso que usa la clase
    ''' </summary>
    Public Sub Dispose()
        Try
            If VMtmrEco.Enabled Then
                VMtmrEco.Stop()
            End If
            VMtcpCliente = Nothing
            VMhiloLeer = Nothing
            VMhiloConexion = Nothing
            VMstmDatos = Nothing
            VMtmrEco.Dispose()
            VMtmrEco = Nothing
        Catch ex As Exception

        End Try

    End Sub
#End Region
#Region "Funciones privadas"
    ''' <summary>
    ''' Permite la conexion con el servidor en un hilo
    ''' </summary>
    Private Sub ConectarServidor()
        Try
            If _Estado = EstadoConexion.Conectado Then
                _Estado = EstadoConexion.EnError
                RaiseEvent ErrorConexion("Se encuentra conectado al servidor " & _Direccion & ":" & _Puerto & " si se quiere conectar primero debe cerrar la conexion existente")
                Exit Sub
            End If
            VMintEcosSinRespuesta = 0
            _Estado = EstadoConexion.Conectando
            VMtcpCliente.NoDelay = _Retardo  '26-Mayo-09 Sandra Aparicio
            'Me conecto al objeto de la clase Servidor, determinado por las propiedades IP y Puerto
            VMtcpCliente.Connect(_Direccion, _Puerto)
            If VMtcpCliente.Connected Then
                _Estado = EstadoConexion.Conectado
                VMstmDatos = VMtcpCliente.GetStream()
                'Inicio un hilo para que escuche los mensajes enviados por el Servidor
                VMhiloLeer.Start()
                ConfigurarEco()
                RaiseEvent ConexionEstablecida()
            Else
                _Estado = EstadoConexion.EnError
                RaiseEvent ErrorConexion("Error de conexion con el servidor: " & _Direccion & ":" & _Puerto)
            End If
        Catch ex As Exception
            _Estado = EstadoConexion.EnError
            RaiseEvent ErrorConexion("Error de conexion con el servidor: " & _Direccion & ":" & _Puerto & " por el siguiente motivo: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Permite configurar el timer para mandar el eco
    ''' </summary>
    Private Sub ConfigurarEco()
        VMtmrEco.AutoReset = True
        VMtmrEco.Interval = _TiempoEco * 1000
        If _ActivoEco Then
            VMtmrEco.Start()
        End If
    End Sub

    ''' <summary>
    ''' Permite escuchar al servidor para saber cuando a llegado datos
    ''' </summary>
    Private Sub LeerSocket()
        Dim VLbytLectura() As Byte
        Dim VLstrDatos As String
        Dim VLintLen As Integer

        '-------Variables para el Control de Error en el Socket-------
        Dim VLblnCorrecto As Boolean
        Dim VLstrError As String = ""
        '-------------------------------------------------------------
        '------Variable para la codificación de los Datos Recibidos----
        Dim VLobjCodificacion As New UTF8Encoding
        '--------------------------------------------------------------
        While True
            Try
                VLbytLectura = New Byte(0) {}
                'Me quedo esperando a que llegue algun mensaje
                VMstmDatos.Read(VLbytLectura, 0, VLbytLectura.Length)
                '----------------------------------------------------------------------------
                'VLstrDatos = Encoding.Default.GetString(VLbytLectura, 0, VLbytLectura.Length)
                VLstrDatos = VLobjCodificacion.GetString(VLbytLectura, 0, VLbytLectura.Length)
                '----------------------------------------------------------------------------
                VLintLen = VMtcpCliente.Available
                If VLintLen > 0 Then
                    VLbytLectura = New Byte(VLintLen) {}
                    VLintLen = VMstmDatos.Read(VLbytLectura, 0, VLbytLectura.Length)
                    '------------------------------------------------------------------------------
                    'VLstrDatos = VLstrDatos & Encoding.Default.GetString(VLbytLectura, 0, VLintLen)
                    VLstrDatos = VLstrDatos & VLobjCodificacion.GetString(VLbytLectura, 0, VLintLen)
                    '------------------------------------------------------------------------------
                End If
                'Limpio las variables del Eco
                If _ActivoEco Then
                    If Not VMtmrEco.Enabled Then
                        ConfigurarEco()
                    End If
                    VMintEcosSinRespuesta = 0
                    VMtmrEco.Stop()
                    VMtmrEco.Start()
                ElseIf Not _ActivoEco And VMtmrEco.Enabled Then
                    VMtmrEco.Stop()
                End If
                If VLstrDatos.Length = 1 Then
                    If Convert.ToChar(VLstrDatos) = Nothing Then
                        Exit While
                    End If
                End If
                'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                RaiseEvent DatosRecibidos(VLstrDatos)
            Catch ex As SocketException
                VLstrError = obtieneError(ex.SocketErrorCode)
            Catch ex As ThreadAbortException
                'Se supone que esta excepción nos indica que estamos cerrando la conexión
                'Se puede controlar mejor si se prende una bandera que indique que se 
                'desea cerrar la conexión
                VLblnCorrecto = True
                Exit While
            Catch ex As Exception
                Dim VLobjError As SocketException

                If Not IsNothing(ex.InnerException) Then
                    If ex.InnerException.GetType.ToString = "System.Net.Sockets.SocketException" Then
                        VLobjError = ex.InnerException
                        VLstrError = obtieneError(VLobjError.SocketErrorCode)
                    Else
                        VLstrError = "Error desconocido [" & ex.Message & "]"
                    End If
                End If
                Exit While
            End Try
        End While

        If VLblnCorrecto Then
            Me.Cerrar()
            RaiseEvent ConexionTerminada(False)
            _Estado = EstadoConexion.Desconectado
        Else
            RaiseEvent ErrorConexion(VLstrError)
            _Estado = EstadoConexion.EnError
        End If

    End Sub
    ''' <summary>
    ''' Permite mandar los ecos al servidor
    ''' </summary>
    Private Sub VMtmrEco_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrEco.Elapsed
        If VMintEcosSinRespuesta = _DesconexionEco Then
            Me.Cerrar()
            RaiseEvent ConexionTerminada(True)
        Else
            VMintEcosSinRespuesta += 1
            If Me.Estado = EstadoConexion.Conectado Then
                Me.EnviarDatos(_MensajeEco)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Permite crear una nueva conexion con el servidor
    ''' </summary>
    Private Sub NuevaConexion()
        VMhiloConexion = New Thread(AddressOf ConectarServidor)
        VMhiloLeer = New Thread(AddressOf LeerSocket)
        VMtcpCliente = New TcpClient()
        VMtmrEco = New Timers.Timer
        VMhiloConexion.Start()
    End Sub

    '''<summary>
    ''' Obtiene el tipo de Error que se presento en el Socket
    ''' </summary> 
    Private Function obtieneError(ByVal VPintError As SocketError) As String
        Select Case VPintError
            Case SocketError.AccessDenied
                Return "Se intentó obtener acceso a un Socket de una manera prohibida por sus permisos de acceso."
            Case SocketError.AddressAlreadyInUse
                Return "La dirección solicitada ya se encuentra en uso"
            Case SocketError.AddressFamilyNotSupported
                Return "No admite la familia de direcciones especificada. Se devuelve este error si se especificó la familia de direcciones IPv6 y la pila del IPv6 no está instalada en el equipo local. Se devuelve este error si se especificó la familia de direcciones IPv4 y la pila del IPv4 no está instalada en el equipo local."
            Case SocketError.AddressNotAvailable
                Return "La dirección IP seleccionada no es válida en este contexto."
            Case SocketError.AlreadyInProgress
                Return "El Socket de no bloqueo ya tiene una operación en curso."
            Case SocketError.ConnectionAborted
                Return ".NET Framework o el proveedor de sockets subyacentes anuló la conexión."
            Case SocketError.ConnectionRefused
                Return "La conexión ha sido rechazada o el host de destino no se encuentra disponible en este momento."
            Case SocketError.ConnectionReset
                Return "El interlocutor remoto restableció la conexión."
            Case SocketError.DestinationAddressRequired
                Return "Se ha omitido una dirección necesaria de una operación en un Socket."
            Case SocketError.Disconnecting
                Return "Se está realizando correctamente una desconexión."
            Case SocketError.Fault
                Return "El proveedor de sockets subyacentes detectó una dirección de puntero no válida."
            Case SocketError.HostDown
                Return "Se ha generado un error en la operación porque el host remoto está inactivo."
            Case SocketError.HostNotFound
                Return "Se desconoce el host. El nombre no es un nombre de host o alias oficial."
            Case SocketError.HostUnreachable
                Return "No hay ninguna ruta de red al host especificado."
            Case SocketError.InProgress
                Return "Hay una operación de bloqueo en curso."
            Case SocketError.Interrupted
                Return "Se canceló una llamada Socket de bloqueo."
            Case SocketError.InvalidArgument
                Return "Se ha proporcionado un argumento no válido a un miembro de Socket."
            Case SocketError.IOPending
                Return "La aplicación ha iniciado una operación superpuesta que no se puede finalizar inmediatamente."
            Case SocketError.IsConnected
                Return "El Socket ya está conectado."
            Case SocketError.MessageSize
                Return "El datagrama es demasiado largo."
            Case SocketError.NetworkDown
                Return "La red no está disponible."
            Case SocketError.NetworkReset
                Return "La aplicación intentó establecer KeepAlive en una conexión cuyo tiempo de espera ya está agotado."
            Case SocketError.NetworkUnreachable
                Return "No existe ninguna ruta al host remoto."
            Case SocketError.NoBufferSpaceAvailable
                Return "No hay espacio en búfer disponible para una operación de Socket."
            Case SocketError.NoData
                Return "No se encontró el nombre o la dirección IP solicitada en el servidor de nombres."
            Case SocketError.NoRecovery
                Return "El error es irrecuperable o no se encuentra la base de datos solicitada."
            Case SocketError.NotConnected
                Return "La aplicación intentó enviar o recibir datos y el Socket no está conectado."
            Case SocketError.NotInitialized
                Return "No se ha inicializado el proveedor de sockets subyacentes."
            Case SocketError.NotSocket
                Return "Se intentó realizar una operación de Socket en algo que no es un socket."
            Case SocketError.OperationAborted
                Return "La operación superpuesta se anuló debido al cierre del Socket."
            Case SocketError.OperationNotSupported
                Return "La familia de protocolos no admite la familia de direcciones."
            Case SocketError.ProcessLimit
                Return "Demasiados procesos están utilizando el proveedor de sockets subyacentes."
            Case SocketError.ProtocolFamilyNotSupported
                Return "La familia de protocolos no está implementada o no está configurada."
            Case SocketError.ProtocolNotSupported
                Return "El protocolo no está implementado o no está configurado."
            Case SocketError.ProtocolOption
                Return "Se ha utilizado una opción o un nivel desconocido, no válido o incompatible con un Socket."
            Case SocketError.ProtocolType
                Return "El tipo de protocolo es incorrecto para este Socket."
            Case SocketError.Shutdown
                Return "Se denegó una solicitud de envío o recepción de datos porque ya se ha cerrado el Socket."
            Case SocketError.SocketError
                Return "Se ha producido un error de Socket no especificado."
            Case SocketError.SocketNotSupported
                Return "Esta familia de direcciones no es compatible con el tipo de socket especificado."
            Case SocketError.Success
                Return "La operación de Socket se ha realizado correctamente."
            Case SocketError.SystemNotReady
                Return "El subsistema de red no está disponible."
            Case SocketError.TimedOut
                Return "El intento de conexión ha sobrepasado el tiempo de espera o el host conectado no ha respondido."
            Case SocketError.TooManyOpenSockets
                Return "Hay demasiados sockets abiertos en el proveedor de sockets subyacentes."
            Case SocketError.TryAgain
                Return "No se pudo resolver el nombre del host. Inténtelo de nuevo más tarde."
            Case SocketError.TypeNotFound
                Return "No se encontró la clase especificada."
            Case SocketError.VersionNotSupported
                Return "La versión del proveedor de sockets subyacentes está fuera del intervalo."
            Case SocketError.WouldBlock
                Return "No se puede finalizar inmediatamente una operación en un socket de no bloqueo."
            Case Else
                Return "Error desconocido"
        End Select
    End Function
#End Region
End Class
