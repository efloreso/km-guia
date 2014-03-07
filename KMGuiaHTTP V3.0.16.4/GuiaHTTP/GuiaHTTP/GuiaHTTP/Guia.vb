Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient

Public Class Guia

#Region "Numeraciones"
    Private Enum TipoCampo
        Numerico = 10
        Alfanumerico = 1
        Entero32 = 11
        Fixed = 20
        Fecha = 30
        NoAplica = 99
    End Enum
#End Region
#Region "Variables de módulo"
#Region "Objeto"
    Private VMobjCola As Queue                          'Se almacena la cadena recibida del Servidor
    Private VMGuiaHTTP As GuiaHTTP                      'Clase local que hace referencia a la clase principal
    Public VMipClientes As New Hashtable                'Detecta los clientes conectados 
    Private VMobjHilo As Threading.Thread               'Hilo que ejecuta el procesamiento de la Transacción
    Public WithEvents VMtcpServidor As New Servidor     'Clase servidor que envia los bloques
#End Region
#Region "Boleanas"
    Private _Procesando As Boolean                      'Variable que controla la ejecucion y creación del Hilo
    Private _Reconectando As Boolean                    'Está en proceso de reconectar el socket
    Private VMblnValidaCola As Boolean                  'Si es true uno de los timers de la cola está validando y el otro debe esperar
    Private VMblnEsperaCola As Boolean                  'Si es true al final de un timer significa que el otro está esperando que termine para validar el otro extremo de la cola
    Private _LevantandoSocket As Boolean                'True si se está llevando a cabo el proceso de listen, 
    Private VMblnUsuarioAutentificado As Boolean        'Determina si el usuario se autentifico y mando la cadena SENDTRAN
#End Region
#Region "Temporizadores"
    Private WithEvents VMtmrReconecta As Timers.Timer   'En caso de que se caiga el socket lo reinicia
    'Private WithEvents VMtmrColaSalida As Timers.Timer  'Verifica si se están enviando datos correctamente
    'Private WithEvents VMtmrColaEntrada As Timers.Timer 'Verifica si se están recibiendo datos correctamente
    Private WithEvents VMtmrProcesa As Timers.Timer     'Envía los datos al cliente
#End Region
#End Region
#Region "Propiedades privadas"
    Private _ultimoError As String
    Private _Restaurando As Boolean
    Private _Duplicando As Boolean
    Private _aplicacion As String           'Nombre de la aplicacion que levantara sus cliente
    Private _servidor As String             'Nombre del Servidor
    Private _puerto As String               'Puerto en el que se levanta el servidor
    Private _echo As Boolean                'Determina si se envian echos
    Private _formato As String              'Formato
    Private _layout As String               'LayOut
    Private _VMobjBase As tipBD             'Propiedades de la conexión a la base
    Private Shared _UsuariosConectados As Integer   'Presenta el numero de usuarios conectados
    Private Shared _IgnoradasTotales As Integer     'Numero de Bloques ignorados
    Private _IP As String                   'Contiene la IP del cliente conectado 
    Private _NombrePC As String = ""        'Contiene el Nombre de la PC del cliente conectado 
    Private _Usuario As String = "Libre"    'Contiene el usuario con el que se esta conectando
    Private _Ignoradas As Integer           'Cuenta el numero de bloques ignoradas
    Private _Procesados As Integer          'Cuenta el numero de bloques procesadas
    Private _Recibidos As Integer           'Cuenta el numero de bloques recibidas
    Private _FechaConexion As String        'La fecha cuando se conecto el cliente
    Private _EstadoPuerto As String = "Libre"       'El estado en el que se encuentra el puerto de conexion con el cliente
    Private _UltimaTransaccion As String    'La fecha cuando se recibio el último bloque
    Private _Replicador As Boolean          'Indica si el puero sólo recibirá conexiones del replicador
    'Private _recuperacion As Boolean
#End Region
#Region "Propiedades públicas"
    Public ReadOnly Property Duplicando() As Boolean
        Get
            Return _Duplicando
        End Get
    End Property
    Public ReadOnly Property Restaurando() As Boolean
        Get
            Return _Restaurando
        End Get
    End Property
    Public ReadOnly Property UltimoError() As String
        Get
            Return _ultimoError
        End Get
    End Property
    'Public Property Recuperaciones() As Boolean
    '    Get
    '        Return _Recuperacion
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _Recuperacion = value
    '    End Set
    'End Property
    Public Property Replicador() As Boolean
        Get
            Return _Replicador
        End Get
        Set(ByVal value As Boolean)
            _Replicador = value
        End Set
    End Property
    Public Property VMobjBase() As Object
        Get
            VMobjBase = _VMobjBase
        End Get
        Set(ByVal value As Object)
            _VMobjBase = value
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
    Public Property Aplicacion() As String
        Get
            Aplicacion = _aplicacion
        End Get
        Set(ByVal value As String)
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
    Public ReadOnly Property Ignoradas()
        Get
            Return _Ignoradas
        End Get
    End Property
    Public Shared ReadOnly Property UsuariosConectados() As Integer
        Get
            Return _UsuariosConectados
        End Get
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
    Public ReadOnly Property UltimaTransaccion() As String
        Get
            Return _UltimaTransaccion
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
    Public Sub New(ByVal VPrincipal As GuiaHTTP)
        Try
            VMobjCola = New Queue
            VMGuiaHTTP = VPrincipal

            VMtmrReconecta = New Timers.Timer
            VMtmrReconecta.Enabled = False
            VMtmrReconecta.Interval = 4000 '4 Segundos

            'VMtmrReconecta.Interval = 100

            'VMtmrColaEntrada = New Timers.Timer
            'VMtmrColaEntrada.Interval = 60000
            'VMtmrColaEntrada.Enabled = False

            'VMtmrColaSalida = New Timers.Timer
            'VMtmrColaSalida.Interval = 60000
            'VMtmrColaSalida.Enabled = False

            VMtmrProcesa = New Timers.Timer
            VMtmrProcesa.Interval = 100
            VMtmrProcesa.Enabled = False
        Catch ex As Exception
            Notifica("No se pudo inicializar la clase para el envío de transacciones " & vbNewLine & ex.Message, VPrincipal, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
        End Try
    End Sub
    Public Sub Enviar(ByVal Datos As String)
        '*****************************************************************************
        'Proceso de paso para recibir los datos y enviarlos a procesar
        'Parámetros:
        '           Datos=Contiene la transacción
        '*****************************************************************************
        If _Usuario <> "Libre" Then
            _UltimaTransaccion = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
            Try
                _Recibidos += 1
                SyncLock VMobjCola
                    VMobjCola.Enqueue(Datos)
                End SyncLock
                'VMtmrColaEntrada.Enabled = False
                'VMtmrColaEntrada.Enabled = True
                'VMtmrColaSalida.Enabled = True
                If Not _Procesando Then
                    VMtmrProcesa.Enabled = True
                End If
            Catch ex As Exception
                Notifica("No se pudo cargar el mensaje en la cola de salida en el puerto " & _puerto & vbNewLine & ex.Message, VMGuiaHTTP, enuEscribeLog.Completo, enuTipoAviso.MensajeError)
                _Ignoradas += 1
            End Try
        End If
    End Sub
    Public Function LevantaServidor() As Boolean
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
        Catch ex As Exception
            _LevantandoSocket = False
            _ultimoError = "Error al abrir el servidor: " & ex.Message
            Return False
        End Try
    End Function
    'Private Sub procesaDatos()
    '    Dim VLstrMensaje As String = ""
    '    Try
    '        SyncLock VMobjCola
    '            While VMobjCola.Count > 0

    '                'If VMobjCola.Count > 0 Then
    '                VLstrMensaje = VMobjCola.Dequeue
    '                'End If

    '                If VLstrMensaje <> "" Then
    '                    If VMblnUsuarioAutentificado Then
    '                        VMtcpServidor.EnviarDatos(VMipClientes(1), VLstrMensaje)
    '                        _Procesados += 1
    '                    Else
    '                        _Ignoradas += 1
    '                    End If
    '                End If
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
    '
    '
    'End Sub
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
    Private Sub VMtcpServidor_DatosRecibidos(ByVal IDTerminal As System.Net.IPEndPoint, ByVal Datos As String) Handles VMtcpServidor.DatosRecibidos
        Try
            If VMipClientes.ContainsValue(IDTerminal) Then
                If Datos.Contains("SENDTRAN") Then
                    _NombrePC = Datos.Substring(51, 14)
                    _Usuario = Datos.Substring(65, 5)
                    _aplicacion = Datos.Substring(71, 2)
                    _FechaConexion = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
                    _EstadoPuerto = "Autentificado"

                    Notifica("Se ha recibido la solicitud de envío de información en el puerto " & _puerto & " con los siguientes datos" & vbNewLine & "IP: " & IDTerminal.ToString & vbNewLine & "PC: " & _NombrePC & vbNewLine & "Usuario: " & _Usuario & "[" & _aplicacion & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
                    'Actica los timers de recepción y envío de datos
                    'VMtmrColaEntrada.Enabled = True

                    VMtcpServidor.EnviarDatos(IDTerminal, Datos)
                    VMblnUsuarioAutentificado = True
                End If
                'VMblnUsuarioAutentificado = True
            End If
        Catch ex As Exception
            Notifica("Error al procesar la petición del cliente [" & IDTerminal.ToString & "]" & "[" & Datos & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.MensajeError, ex.Message)
            _NombrePC = "N/R"
            _Usuario = "N/R"
            _aplicacion = "N/R"

            '------AGREGADO 15-11-11----------------
            Try
                VMtcpServidor.Cerrar()
                '-----------------------------------
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
                '------------------------------------
                VMtmrReconecta.Enabled = True
            Catch ex1 As Exception
                Notifica("1. Conflicto al intentar reanudar comunicación  en el puerto [" & _puerto & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.MensajeError, ex.Message)
            End Try
            '-----------------------------------------------

        End Try

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
    Private Sub VMtcpServidor_NuevaConexion(ByVal IDTerminal As System.Net.IPEndPoint) Handles VMtcpServidor.NuevaConexion
        If VMipClientes.Count = 0 Then
            _UsuariosConectados += 1
Acepta:
            VMipClientes.Add(1, IDTerminal)
            _IP = IDTerminal.ToString
            _Usuario = "Sin autentificar"
            _EstadoPuerto = "Sin autentificar"
            'guardaBitacora("Se conecto el cliente " & IDTerminal.ToString, 1)
        Else
            'guardaBitacora("El puerto esta ocupado cuando se intento conctar " & IDTerminal.ToString, 8)
            'If Clientes > 1 Then
            '    'VMtcpServidor.EnviarDatos(IDTerminal, "Puerto Ocupado")
            '    Threading.Thread.Sleep(500)
            '    VMtcpServidor.CerrarCliente(IDTerminal)
            'Else
            '    VMipClientes.Clear()
            '    GoTo Acepta
            'End If
        End If
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
            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "Se ha restablecido el servidor en el puerto " & _puerto & ETX
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
    'Private Sub VMtmrColaSalida_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrColaSalida.Elapsed
    '    If Not VMblnValidaCola Then
    '        VMblnValidaCola = True
    '        Try
    '            VMtmrColaSalida.Enabled = False

    '            'Ha ocurrido un minuto sin enviar información
    '            Notifica("Validando estado de envío [" & _puerto & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)

    '            If Not VMblnValidaCola Then
    '                'Verifica que la cola no contenga datos
    '                If VMobjCola.Count > 0 And Not _Procesados Then
    '                    VMtmrProcesa.Enabled = True
    '                Else
    '                    'VMtmrColaEntrada.Enabled = False
    '                    VMGuiaHTTP.validaEnvio()
    '                End If
    '            End If

    '            VMtmrColaSalida.Enabled = True
    '        Catch ex As Exception
    '            VMtmrColaSalida.Enabled = True
    '        End Try
    '        VMblnValidaCola = False
    '    End If
    'End Sub
    Private Sub VMtmrProcesa_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrProcesa.Elapsed
        Dim VLstrMensaje As String
        _Procesando = True
        VMtmrProcesa.Enabled = False
        Try
            While VMobjCola.Count > 0
                Try
                    SyncLock VMobjCola
                        VLstrMensaje = VMobjCola.Dequeue
                    End SyncLock
                Catch ex As InvalidOperationException
                    'La cola está vacía
                    _Procesando = False
                    Return
                Catch ex As Exception
                    'Si la cola no está vacía vuelve a prender el timer
                    If VMobjCola.Count > 0 Then
                        _Procesando = False
                        VMtmrProcesa.Enabled = True
                    End If
                    Return
                End Try

                'VMtmrColaSalida.Enabled = False
                If VLstrMensaje <> "" Then
                    If VMblnUsuarioAutentificado Then
                        Try
                            VMtcpServidor.EnviarDatos(VMipClientes(1), VLstrMensaje)
                        Catch ex As Exception
                            Notifica("No se pudo enviar el mensaje al cliente en el puerto " & _puerto & vbNewLine & "Error: " & ex.Message, VMGuiaHTTP, enuEscribeLog.Completo, enuTipoAviso.MensajeError)
                        End Try
                        _Procesados += 1
                    Else
                        _Ignoradas += 1
                    End If
                End If
                'VMtmrColaSalida.Enabled = True
            End While
            'VMtmrColaSalida.Enabled = False
        Catch ex As Exception
            Notifica("Ocurrió una excepción al procesar la cola de datos en el puerto " & _puerto & vbNewLine & " Error:" & ex.Message, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
            'Verifica que la cola no quede con datos
            If VMobjCola.Count > 0 Then
                VMtmrProcesa.Enabled = True
                'VMtmrColaSalida.Enabled = True
            End If
        End Try

        _Procesando = False
    End Sub
    'Private Sub VMtmrColaEntrada_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrColaEntrada.Elapsed
    '    If Not VMblnValidaCola Then
    '        VMblnValidaCola = True
    '        VMtmrColaEntrada.Enabled = False
    '        VMGuiaHTTP.validaEnvio()
    '        VMtmrColaEntrada.Enabled = True
    '        VMblnValidaCola = False
    '    End If
    'End Sub

#Region "RESTAURAR FILTROS"
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
                VLstrRespuesta = STX & 8014 & FS & ACK & FS & "RC" & FS & _puerto & ETX
                VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                Notifica("Se ha restablecido el servidor en el puerto " & _puerto, VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
            Else
                VLstrRespuesta = STX & 8014 & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & _puerto & ETX
                VMGuiaHTTP.VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                Notifica("No se pudo restablecer el servidor en el puerto " & _puerto & "  [" & _ultimoError & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
            End If
            _Reconectando = False
        Catch ex As Exception
            VMGuiaHTTP.VMtcpAdmin.EnviarDatos(STX & 8014 & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & _puerto & "[" & ex.Message & "]" & ETX)
            Notifica("No se pudo restablecer el servidor en el puerto " & _puerto & "  [" & ex.Message & "]", VMGuiaHTTP, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
        End Try
    End Sub
#End Region

End Class
