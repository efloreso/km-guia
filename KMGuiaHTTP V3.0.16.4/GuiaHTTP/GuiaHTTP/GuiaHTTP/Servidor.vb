Option Explicit On

Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.Net.Sockets

Public Class Servidor
#Region "NUMERACIONES"
    Public Enum enuEstado As Byte
        Cerrado = 0
        Abierto = 1
        Escuchando = 2
        ConexionPendiente = 3
        ResolviendoHost = 4
        HostResuelto = 5
        Conectando = 6
        Conectado = 7
        Cerrando = 8
        EnError = 9
    End Enum
#End Region
#Region "VARIABLES"
    Private tcpThd As Thread                        'Se ocupa para levantar el hilo  que espera la conexión de un cliente remoto
    Private _IPLocal As String                      'Guarda la dirección de la primer tarjeta de red IPV4
    Private _Estado As enuEstado                    'Estado de la clase
    Private VMstrDatos As String                    'Son los datos leídos del cliente
    Private tcpLsn As TcpListener                   'Es la variable que recibe la conexión de un nuevo cliente
    Private VMsckCliente As Socket                  'Guarda el socket cliente
    Private VMthdLectura As Thread                  'Hilo que está en espera de datos enviados por el cliente
    Private _PuertoDeEscucha As Integer

    Private WithEvents VMtmrLimpiar As Timers.Timer 'Es el hilo que limpia las variables cuando se desconecta el cliente
#End Region
#Region "EVENTOS"
    Public Event NuevaConexion(ByVal IDTerminal As Net.IPEndPoint)
    Public Event DatosRecibidos(ByVal IDTerminal As Net.IPEndPoint, ByVal Datos As String)
    Public Event ConexionTerminada(ByVal IDTerminal As Net.IPEndPoint)
    Public Event ErrorServidor(ByVal Mensaje As String)
#End Region
#Region "PROPIEDADES"
    Property PuertoEscucha() As Integer
        Get
            PuertoEscucha = _PuertoDeEscucha
        End Get

        Set(ByVal Value As Integer)
            _PuertoDeEscucha = Value
        End Set
    End Property
    ReadOnly Property Estado() As enuEstado
        Get
            Estado = _Estado
        End Get
    End Property
#End Region
#Region "METODOS"
    Public Sub New()
        Dim hostInfo As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(My.Computer.Name)

        For Each ip As System.Net.IPAddress In hostInfo.AddressList
            If ip.AddressFamily = AddressFamily.InterNetwork Then
                _IPLocal = ip.ToString
            End If
        Next
        VMtmrLimpiar = New Timers.Timer
        _Estado = enuEstado.Cerrado
    End Sub
    Public Sub Escuchar(ByVal Puerto As Integer)
        _PuertoDeEscucha = Puerto
        Escuchar()
    End Sub
    Public Sub Escuchar()
        Dim localAddr As IPAddress = IPAddress.Parse("0.0.0.0")

        tcpLsn = New TcpListener(localAddr, _PuertoDeEscucha)
        Try
            tcpLsn.Start()

            _Estado = enuEstado.Escuchando

            tcpThd = New Thread(AddressOf EsperarCliente)
            tcpThd.IsBackground = False
            tcpThd.Start()

        Catch ex As Exception
            RaiseEvent ErrorServidor("No se pudo realizar la conexión el el puerto " & _PuertoDeEscucha & " por el siguiente motivo" & vbNewLine & ex.Message)
            _Estado = enuEstado.Cerrado
        End Try
    End Sub
    Public Sub Cerrar()
        _Estado = enuEstado.Cerrando
        'Detiene el hilo
        If Not IsNothing(tcpThd) Then
            If tcpThd.IsAlive Then
                tcpThd.Abort()
                tcpLsn.Stop()
            End If
        End If

        tcpThd = Nothing
        tcpLsn = Nothing
        Try
            If Not IsNothing(VMsckCliente) Then
                If VMsckCliente.Connected Then
                    VMsckCliente.Shutdown(SocketShutdown.Both)
                    VMsckCliente.Close()
                End If
            End If
        Catch ex As Exception
        Finally
            VMsckCliente = Nothing
        End Try


        Try
            VMthdLectura.Abort()
        Catch ex As ThreadAbortException
            'Debug.Print(ex.Message)
            VMthdLectura = Nothing
        Catch ex As Exception
            'Debug.Print(ex.Message)
            VMthdLectura = Nothing
        Finally
            VMthdLectura = Nothing
        End Try

        _Estado = enuEstado.Cerrado
    End Sub
    Public Sub EnviarDatos(ByVal IDCliente As Net.IPEndPoint, ByVal Datos As String)
        Dim VLobjCodifica As New ASCIIEncoding

        'Le envio el mensaje 
        VMsckCliente.Send(VLobjCodifica.GetBytes(Datos))
    End Sub
#End Region
#Region "FUNCIONES PRIVADAS"
    Private Sub EsperarCliente()

        'Guardo el Socket que utilizo para mantener la conexion con el cliente 
        VMsckCliente = tcpLsn.AcceptSocket() 'Se queda esperando la conexion de un cliente 
        VMsckCliente.NoDelay = True

        tcpLsn.Stop()
        tcpLsn = Nothing

        'Guardo el el RemoteEndPoint, que utilizo para identificar al cliente 
        'VMtcpCliente = VMsckCliente.RemoteEndPoint


        'Creo un Thread para que se encargue de escuchar los mensaje del cliente 
        VMthdLectura = New Thread(AddressOf LeerSocket)

        'Genero el evento Nueva conexion 
        RaiseEvent NuevaConexion(VMsckCliente.RemoteEndPoint)

        'Inicio el thread encargado de escuchar los mensajes del cliente 
        VMthdLectura.Start()
    End Sub
    Private Sub LeerSocket()
        Dim Recibir() As Byte 'Array utilizado para recibir los datos que llegan 
        Dim VLstrAux As String
        Dim VLobjDecodifica As New ASCIIEncoding
        Dim Ret As Integer = 0

        'With InfoClienteActual
        While True
            If VMsckCliente.Connected Then
                'Recibir = New Byte(100) {}
                Try
                    'Me quedo esperando a que llegue un mensaje desde el cliente 
                    'Ret = .Socket.Receive(Recibir, Recibir.Length, SocketFlags.None)   
                    Recibir = New Byte(0) {}
                    Ret = VMsckCliente.Receive(Recibir, Recibir.Length, SocketFlags.None)
                    VLstrAux = Mid(VLobjDecodifica.GetString(Recibir), 1, Ret) ' Mid(Encoding.Unicode.GetString(Recibir), 1, Ret)
                    'Mid(Encoding.ASCII.GetString(Recibir), 1, Ret)
                    Ret = VMsckCliente.Available
                    If _Estado = enuEstado.Cerrado Or Estado = enuEstado.Cerrando Then
                        Exit While
                    End If
                    If Ret > 0 Then
                        Recibir = New Byte(Ret - 1) {}
                        Ret = VMsckCliente.Receive(Recibir, Recibir.Length, SocketFlags.None)
                    End If
                    'Verifica cuanta información está pendiente y la procesa
                    If Ret > 0 Or VLstrAux.Length > 0 Then
                        'Guardo el mensaje recibido 
                        If Ret > 0 Then
                            VMstrDatos = VLstrAux & Mid(VLobjDecodifica.GetString(Recibir), 1, Ret)
                        Else
                            VMstrDatos = VLstrAux
                        End If

                        'Genero el evento de la recepcion del mensaje 
                        RaiseEvent DatosRecibidos(VMsckCliente.RemoteEndPoint, VMstrDatos)
                    Else
                        'Genero el evento de la finalizacion de la conexion 
                        RaiseEvent ConexionTerminada(VMsckCliente.RemoteEndPoint)
                        Exit While
                    End If
                Catch e As Exception
                    If Not _Estado = enuEstado.Cerrando And Not _Estado = enuEstado.Cerrado Then
                        If Not VMsckCliente.Connected Then
                            'Genero el evento de la finalizacion de la conexion 
                            RaiseEvent ConexionTerminada(VMsckCliente.RemoteEndPoint)
                            Exit While
                        End If
                    End If
                End Try
            End If
        End While

        If Not _Estado = enuEstado.Cerrando And Not _Estado = enuEstado.Cerrado Then
            VMtmrLimpiar.Interval = 150
            VMtmrLimpiar.Enabled = True
        End If
        'CerrarThread()
        'End With
    End Sub
    Private Sub VMtmrLimpiar_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrLimpiar.Elapsed
        VMtmrLimpiar.Enabled = False

        _Estado = enuEstado.Cerrando
        Try
            VMthdLectura.Abort()
        Catch ex As ThreadAbortException
            'Debug.Print(ex.Message)
            VMthdLectura = Nothing
        Catch ex As Exception
            'Debug.Print(ex.Message)
            VMthdLectura = Nothing
        Finally
            VMthdLectura = Nothing
        End Try
        _Estado = enuEstado.Cerrado

        'Vuelve a llamar a la función escuchar para que se abra de nuevo el listen 
        Escuchar(_PuertoDeEscucha)
    End Sub
#End Region
End Class
