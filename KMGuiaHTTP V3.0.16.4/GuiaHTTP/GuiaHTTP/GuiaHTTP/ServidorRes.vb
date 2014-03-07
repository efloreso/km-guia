Option Explicit On
'*******************************************************************************************************
'FECHA DE MODIFICACIÓN: 15/DICIEMBRE/2008 
'NOMBRE: SANDRA APARICIO
'MODIFICACIONES:
'1. Modificaciones en el evento de datos recibidos(Se envian los datos en formato cadena y en bytes)
'2. Función LeerSocket
'FECHA DE MODIFICACIÓ0N 16/FEBRERO/2009
'1. Modificaciones en la codificación de los bytes enviados y recibidos. Se cambio de condificación ASCII, a la codificación por Default
'FECHA DE MODIFICACIÓ0N 26/MAYO/2009
'1. Modificaciones en la codificación de los bytes enviados y recibidos. Se cambio de condificación por Default, a la codificación UTF8
'FECHA DE MODIFICACIÓN 08/JULIO/2009
'1. Se agrego el método ConexionUnica, de tipo Boolean, para establecer que el Socket solo permitira la conexión de un cliente.  
'FECHA DE MODIFICACION 4/AGOSTO/2009
'1. Se agrego validación de Socket <> Nothing al momento de enviar y recibir datos. 
'FECHA DE MODIFICACION 6/OCTUBRE/2009
'1. Se comento el código referente al método de ConexionUnica
'*******************************************************************************************************
Imports System
Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Public Class Servidor

#Region "NUMERACIONES"
    Public Enum sckEstado As Byte
        sckCerrado = 0
        sckAbierto = 1
        sckEscuchando = 2
        sckConexionPendiente = 3
        sckResolviendoHost = 4
        sckHostResuelto = 5
        sckConectando = 6
        sckConectado = 7
        sckCerrando = 8
        sckError = 9
    End Enum
#End Region
#Region "ESTRUCTURAS"
    Private Structure InfoCliente
        'Esta estructura permite guardar la información sobre un cliente 
        Public Socket As Socket 'Socket utilizado para mantener la conexion con el cliente 
        Public Thread As Thread 'Thread utilizado para escuchar al cliente 
        Public UltimosDatosRecibidos As String 'Ultimos datos enviados por el cliente 
    End Structure
#End Region
#Region "VARIABLES"
    Private tcpLsn As TcpListener
    Private Clientes As New Hashtable() 'Aqui se guarda la informacion de todos los clientes conectados 
    Private tcpThd As Thread
    Private IDClienteActual As Net.IPEndPoint 'Ultimo cliente conectado 
    Private _PuertoDeEscucha As Integer
    Private _Estado As sckEstado
    Private _IPLocal As String
    Private VMblnCerrar As Boolean
    'Private _ConexionUnica As Boolean
#End Region
#Region "EVENTOS"
    Public Event NuevaConexion(ByVal IDTerminal As Net.IPEndPoint, ByVal Clientes As Integer)
    Public Event DatosRecibidos(ByVal IDTerminal As Net.IPEndPoint, ByVal Datos As String, ByVal DatosBytes() As Byte)
    Public Event ConexionTerminada(ByVal IDTerminal As Net.IPEndPoint)
    Public Event ErrorServidor(ByVal Mensaje As String)
    Public Event ErrorLog(ByVal Mensaje As String)
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

    ReadOnly Property Estado() As sckEstado
        Get
            Estado = _Estado
        End Get
    End Property

    'Property ConexionUnica() As Boolean
    '    Get
    '        Return _ConexionUnica
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _ConexionUnica = value
    '    End Set
    'End Property
#End Region
#Region "METODOS"

    Public Sub New()
        Dim hostInfo As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(My.Computer.Name)

        For Each ip As System.Net.IPAddress In hostInfo.AddressList
            _IPLocal = ip.ToString
        Next
        _Estado = sckEstado.sckCerrado
        VMblnCerrar = False
        '_ConexionUnica = False
    End Sub
    Public Sub Escuchar()
        Dim VLipDirLocal As IPAddress
        Try
            If _Estado = sckEstado.sckCerrado Then
                _Estado = sckEstado.sckConectando
                VLipDirLocal = IPAddress.Parse(_IPLocal)

                tcpLsn = New TcpListener(VLipDirLocal, _PuertoDeEscucha)


                'Inicio la escucha 
                tcpLsn.Start()
                _Estado = sckEstado.sckEscuchando

                'Creo un thread para que se quede escuchando la llegada de un cliente 
                tcpThd = New Thread(AddressOf EsperarCliente)
                tcpThd.IsBackground = False
                tcpThd.Start()
            Else
                RaiseEvent ErrorServidor("No se puede iniciar una nueva conexión por que el Socket ya está conectado")
            End If
        Catch ex As Exception
            RaiseEvent ErrorServidor(ex.Message)
        End Try

    End Sub
    Public Sub Escuchar(ByVal Puerto As Integer)
        'Dim VLipDirLocal As IPAddress
        'Dim VLbytIP() As Byte

        _PuertoDeEscucha = Puerto
        'VLbytIP = Net.IPAddress.Loopback.GetAddressBytes
        'VLipDirLocal = IPAddress.Parse(_IPLocal)
        'tcpLsn = New TcpListener(VLipDirLocal, _PuertoDeEscucha)
        Dim localAddr As IPAddress = IPAddress.Parse("0.0.0.0")

        tcpLsn = New TcpListener(localAddr, _PuertoDeEscucha)

        Try
            tcpLsn.Start()

            _Estado = sckEstado.sckEscuchando

            tcpThd = New Thread(AddressOf EsperarCliente)
            tcpThd.IsBackground = False
            tcpThd.Start()

        Catch ex As Exception
            'MsgBox(ex.Message)
            RaiseEvent ErrorServidor(ex.Message)
        End Try
    End Sub
    Public Function ObtenerDatos(ByVal IDCliente As Net.IPEndPoint) As String
        Dim InfoClienteSolicitado As InfoCliente
        'Obtengo la informacion del cliente solicitado 
        Try
            InfoClienteSolicitado = Clientes(IDCliente)

            ObtenerDatos = InfoClienteSolicitado.UltimosDatosRecibidos
        Catch ex As Exception
            RaiseEvent ErrorLog("Error al obtener Datos" & vbTab & ex.Message)
            Call CerrarThread(IDCliente)
            ObtenerDatos = ""
        End Try
    End Function
    Public Sub CerrarCliente(ByVal IDCliente As Net.IPEndPoint)
        Dim InfoClienteActual As InfoCliente
        'Obtengo la informacion del cliente solicitado 
        Try
            InfoClienteActual = Clientes(IDCliente)
            'Cierro la conexion con el cliente 
            InfoClienteActual.Socket.Close()
        Catch ex As Exception
            RaiseEvent ErrorLog("Error al cerrar cliente" & vbTab & ex.Message)
            Call CerrarThread(IDCliente)
        End Try
        VMblnCerrar = True
    End Sub
    Public Sub CerrarClientes()
        Dim InfoClienteActual As InfoCliente
        'Recorro todos los clientes y voy cerrando las conexiones 
        For Each InfoClienteActual In Clientes.Values
            Call CerrarCliente(InfoClienteActual.Socket.RemoteEndPoint)
        Next
    End Sub
    Public Sub Cerrar()
        _Estado = sckEstado.sckCerrando
        'Detiene el hilo
        tcpThd.Abort()

        tcpLsn.Stop()
        _Estado = sckEstado.sckCerrado

    End Sub
    Public Sub EnviarDatos(ByVal IDCliente As Net.IPEndPoint, ByVal Datos As String)
        Dim Cliente As InfoCliente
        Dim VLobjCodificacion As New UTF8Encoding
        'Obtengo la informacion del cliente al que se le quiere enviar el mensaje 
        Cliente = Clientes(IDCliente)

        Try
            If Not IsNothing(Cliente.Socket) Then
                'Le envio el mensaje 
                '-------------------------------------------
                'Cliente.Socket.Send(Encoding.Default.GetBytes(Datos))
                Cliente.Socket.Send(VLobjCodificacion.GetBytes(Datos))
                '-------------------------------------------
            End If
        Catch ex As Exception
            RaiseEvent ErrorLog("Error al Enviar Datos" & vbTab & ex.Message)
            Call CerrarCliente(IDCliente)
        End Try

    End Sub
    Public Sub EnviarDatos(ByVal Datos As String)
        Dim Cliente As InfoCliente

        'Recorro todos los clientes conectados, y les envio el mensaje recibido 
        'en el parametro Datos 
        For Each Cliente In Clientes.Values
            EnviarDatos(Cliente.Socket.RemoteEndPoint, Datos)
        Next

    End Sub
    Public Sub EnviarArchivo(ByVal IDCliente As Net.IPEndPoint, ByVal Datos() As Byte)
        'SANDRA APARICIO 19-ENERO-09
        Dim Cliente As InfoCliente
        Cliente = Clientes(IDCliente)
        Try
            If Not IsNothing(Cliente.Socket) Then
                Cliente.Socket.Send(Datos)
            End If
        Catch ex As Exception
            RaiseEvent ErrorLog("Error al enviar archivo" & vbTab & ex.Message)
            Call CerrarCliente(IDCliente)
        End Try
    End Sub
#End Region
#Region "FUNCIONES PRIVADAS"
    Private Sub EsperarCliente()
        Dim InfoClienteActual As InfoCliente = New InfoCliente

        With InfoClienteActual
            While True

                'Cuando se recibe la conexion, guardo la informacion del cliente 
                'Guardo el Socket que utilizo para mantener la conexion con el cliente 

                .Socket = tcpLsn.AcceptSocket() 'Se queda esperando la conexion de un cliente 

                'If _ConexionUnica And Clientes.Count > 0 Then
                '    .Socket.Shutdown(SocketShutdown.Both)
                '    .Socket.Close()
                'Else

                'Guardo el el RemoteEndPoint, que utilizo para identificar al cliente 
                IDClienteActual = .Socket.RemoteEndPoint


                'Creo un Thread para que se encargue de escuchar los mensaje del cliente 
                .Thread = New Thread(AddressOf LeerSocket)


                'Agrego la informacion del cliente al HashArray Clientes, donde esta la 
                'informacion de todos estos 
                SyncLock Me
                    Clientes.Add(IDClienteActual, InfoClienteActual)
                End SyncLock

                'Genero el evento Nueva conexion 
                RaiseEvent NuevaConexion(IDClienteActual, Clientes.Count)

                'Inicio el thread encargado de escuchar los mensajes del cliente 
                If Not VMblnCerrar Then
                    .Thread.Start()
                Else
                    SyncLock Me
                        Clientes.Remove(IDClienteActual)
                    End SyncLock

                    VMblnCerrar = False
                End If
            End While
        End With
    End Sub
    Private Sub LeerSocket()
        Dim IDReal As Net.IPEndPoint 'ID del cliente que se va a escuchar 
        Dim Recibir() As Byte 'Array utilizado para recibir los datos que llegan 
        Dim InfoClienteActual As InfoCliente 'Informacion del cliente que se va escuchar 
        Dim VLstrAux As String
        Dim VLobjCodificacion As New UTF8Encoding

        Dim Ret As Integer = 0

        IDReal = IDClienteActual
        InfoClienteActual = Clientes(IDReal)

        With InfoClienteActual
            While True
                Try
                    If Not IsNothing(.Socket) Then
                        If .Socket.Connected Then
                            Try
                                'Me quedo esperando a que llegue un mensaje desde el cliente 
                                'Ret = .Socket.Receive(Recibir, Recibir.Length, SocketFlags.None)
                                Recibir = New Byte(0) {}
                                Ret = .Socket.Receive(Recibir, Recibir.Length, SocketFlags.None)
                                '----------------------------------------------------------
                                VLstrAux = Mid(VLobjCodificacion.GetString(Recibir), 1, Ret)
                                '----------------------------------------------------------
                                Ret = .Socket.Available
                                If Ret > 0 Then
                                    ReDim Preserve Recibir(Ret)
                                    Ret = .Socket.Receive(Recibir, 1, Recibir.Length - 1, SocketFlags.None)
                                End If
                                'Verifica cuanta información está pendiente y la procesa
                                If Ret > 0 Or VLstrAux.Length > 0 Then
                                    'Guardo el mensaje recibido 
                                    If Ret > 0 Then
                                        '-------------------------------------------------------------------------------
                                        .UltimosDatosRecibidos = VLobjCodificacion.GetString(Recibir, 0, Recibir.Length)
                                        '-------------------------------------------------------------------------------
                                    Else
                                        .UltimosDatosRecibidos = VLstrAux
                                    End If
                                    Clientes(IDReal) = InfoClienteActual

                                    '**********************************************************************************************
                                    'Se genera el evento de datos recibidos enviando la información en formato de Cadena y en Bytes
                                    RaiseEvent DatosRecibidos(IDReal, .UltimosDatosRecibidos, Recibir)
                                    '**********************************************************************************************
                                Else
                                    'Genero el evento de la finalizacion de la conexion 
                                    RaiseEvent ConexionTerminada(IDReal)
                                    Exit While
                                End If
                            Catch e As Exception
                                If Not .Socket.Connected Then
                                    'Genero el evento de la finalizacion de la conexion 
                                    RaiseEvent ConexionTerminada(IDReal)
                                    'Espera un segundo para cerrar el socket para que se cierre del lado del cliente
                                    'Thread.Sleep(1000)
                                    Exit While
                                End If
                            End Try
                        Else
                            'Se sale de la rutina
                            RaiseEvent ConexionTerminada(IDReal)
                            Exit While
                        End If
                    Else
                        RaiseEvent ErrorLog("Conexión terminada por Socket Nulo")
                        RaiseEvent ConexionTerminada(IDReal)
                        Exit While
                    End If
                Catch ex As SocketException
                    RaiseEvent ErrorLog("Conexión terminada por excepción de Socket" & vbTab & ex.Message)
                    RaiseEvent ConexionTerminada(IDReal)
                    Exit While
                Catch ex As Exception
                    'Se genera el evento de cerrar conexión
                    RaiseEvent ErrorLog("Conexión terminada por excepción" & vbTab & ex.Message)
                    RaiseEvent ConexionTerminada(IDReal)
                    Exit While
                End Try

            End While
            Call CerrarThread(IDReal)
        End With
    End Sub
    Private Sub CerrarThread(ByVal IDCliente As Net.IPEndPoint)
        Dim InfoClienteActual As InfoCliente
        'Cierro el thread que se encargaba de escuchar al cliente especificado 
        InfoClienteActual = Clientes(IDCliente)
        Try
            InfoClienteActual.Thread.Abort()
        Catch e As Exception
            RaiseEvent ErrorLog("Cerrar Hilo" & vbTab & e.Message)
            SyncLock Me
                'Elimino el cliente del HashArray que guarda la informacion de los clientes 
                Clientes.Remove(IDCliente)
            End SyncLock
        End Try
    End Sub
#End Region
End Class
