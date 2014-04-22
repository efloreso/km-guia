Imports MySql.Data.MySqlClient
Imports Microsoft.Win32
Imports System.Text
Imports System.Data.OracleClient
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Data.SqlClient

Public Class GuiaHTTP
    Private Enum enuMensajesAdmin
        EstadoServicio = 7004
        DescribeError = 7014
        EstadoPuertoTxn = 9004
        EstadoCliente = 8004
        Depurar = 5004
    End Enum
    Private WithEvents VMtmrInicia As Timers.Timer                  'Temporizador que inicializa el servicio - En caso de existir un error se reactivara
    Private WithEvents VMtmrProcesa As Timers.Timer                 'Temporizador para enviar la información de la cola de datos
    Public WithEvents VMtcpCliente As Cliente_TCP                   'Cliente que recibe las transaciones del SERVIDOR 
    Private VMhasPuertos As New Hashtable                           'Guarda los puertos que se han levantado
    Private _Procesando As Boolean                                  'Variable que controla la ejecucion y creación del Hilo
    Private _timeOut As Boolean
    Private _Clientes As Integer                                    'Contador del numero de clientes que se estan procesando
    Public VMobjBloques As New Queue                                'Estructura de datos con las tramas recibidas del servidor transaccional
    Private VMblnRevbloques As Boolean                              'Controla la inicialización del Timer(VMtmrRevBloques)
    Private VMstrError As String                                    'Controla los mensajes de error
    Private VMobjClientes As Hashtable                              'Tabla virtual que contiene la información de los Clientes(replicadores y receptores)
    Dim VMmsgEnviaTransacciones As tipHdrComunicaciones             'Contiene Informacion de los mensajes que se envian al servidor Transaccional
    Private WithEvents VMtmrTiempoSinTran As System.Timers.Timer    'Se activa en caso de que no se esten Recibiendo transacciones y hace un Reinicio
    Private WithEvents VMtmrCambiaTmr As System.Timers.Timer        'Revisa cuando es de dia o de Noche, Para poder cambiar el valor de VMtmrTiempoSinTran
    Dim VMintTiempoCambio As Integer                                'Tiene el horario el valor de VMtmrCambiaTmr(Dia/Noche)
    Private VMblnSigue As Boolean                                   'Espera a recibir una respuesta de conexion
    Dim VMobjRc4 As New RC4                                         'Se ocupa para desencriptar las contraseñas de configuración
    Dim VMintBloques As Integer                                     'Cuenta el numero de bloques recibidos
    Dim VMConfiguracion As String                                   'Almacena el tipo de configuración
    Private VMbytIntentaReconexion                                  'Indica el número de reintentos para reconectarse al servidor transaccional
    Private VMintSinTransaccion                                     'Minutos que han transcurrido sin una transacción
    'Variables del Administrador
    Public WithEvents VMtcpAdmin As Cliente_TCP                     'Conexion con el Administrador
    Private VMintProcesados As Integer                              'Numero de bloques procesados
    Private VMintIgnorados As Integer                               'Número de bloques ignorados
    Private VMstrUltimoBloque As String                             'Fecha del ultimo bloque procesado
    Private VMblnListo1 As Boolean                                  'Si es true se ha confirmado el sendtran
    Private WithEvents VMtmrReconecta As Timers.Timer               'Timer para reconectar a la 
    Private VMblnReconectando As Boolean                            'Si es true se encuentra reconectando con el servidor transaccional
    Private VMstrErrorConsola As String                             'Almacena el último error para que sea enviado a la consola
    Private VMblnValidaEnvio As Boolean
    Private VMobjDepuracion As Depuracion

    Protected Overrides Sub OnStart(ByVal args() As String)
        VMobjDepuracion = Depuracion.Instancia
        Threading.Thread.Sleep(15000)
        If args.GetLength(0) > 0 Then
            Select Case args(0).Trim.ToLower
                Case "/ds"
                    _Log = enuDepuracion.Sencilla
                    abreArchivoLog()
                Case "/dt"
                    _Log = enuDepuracion.Completa
                    abreArchivoLog()
                Case "/dc"
                    _Log = enuDepuracion.Consola
                Case "/dl"
                    VMobjDepuracion.Depuracion = Depuracion.TipoDepuracion.Completa
                Case Else
                    _Log = enuDepuracion.Ninguna
            End Select
        Else
            _Log = enuDepuracion.Ninguna
        End If

        VMtmrProcesa = New Timers.Timer
        VMtmrProcesa.Interval = 150
        VMtmrProcesa.Enabled = False

        VMtmrInicia = New System.Timers.Timer
        VMtmrInicia.AutoReset = False
        VMtmrInicia.Interval = 150
        VMtmrInicia.Enabled = True


        VMtmrReconecta = New Timers.Timer
        VMtmrReconecta.Enabled = False
        VMtmrReconecta.Interval = 100
    End Sub
    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        Try
            If Not VMtcpCliente Is Nothing Then
                If VMtcpCliente.Estado = Cliente_TCP.EstadoConexion.Conectado Then
                    VMtcpCliente.Cerrar()
                End If
            End If

        Catch ex As Exception

        End Try
        VMtmrProcesa.Stop()
        VMtmrProcesa.Dispose()
        VMtmrProcesa = Nothing
        Try
            'If Not IsNothing(VMobjHilo) Then
            '    If VMobjHilo.IsAlive Then
            '        VMobjHilo.Abort()
            '    End If
            'End If
            If Not VMhasPuertos Is Nothing Then
                VMhasPuertos.Clear()
            End If
        Catch ex As Exception

        End Try


        'RequestAdditionalTime(20000)
        If ExitCode = 0 Then
            enviaCorreo("La aplicación se ha cerrado satisfactoriamente.||2")
        Else
            enviaCorreo("Se ha forzado el cierre de la aplicación posiblemente por un evento grave." & vbNewLine & "Revise el visor de eventos para mayor detalle||3")
        End If


        VMhasPuertos = Nothing
        escribeVisorEventos("El servicio se detuvo Correctamente", EventLogEntryType.Information)
        End
    End Sub
    Private Sub VMtmrInicia_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrInicia.Elapsed
        Dim VLioTipo As StreamReader
        Dim VLstrTipo As String
        Dim VLclsDatos As New RC4
        Dim VLstrDatos() As String
        Dim VLstrConfiguracion As String = ""
        Dim VLstrSeparador() As String = {"@|"}
        Dim VLstrLlave As String
        Dim VLstrRespuesta As String

        VMtmrInicia.Stop()
        Try
            VLioTipo = New StreamReader(My.Application.Info.DirectoryPath & "\KSGuia.dat")
            VLstrTipo = VLioTipo.ReadLine()
            VLioTipo.Close()
            VLioTipo = Nothing

            Try
                VLstrLlave = My.Resources.Tipo & My.Application.Info.Version.Major '& My.Application.Info.Version.Minor & My.Application.Info.Version.Build
                VLstrTipo = VLclsDatos.RC4Decrypt(VLstrTipo, VLstrLlave)
                If VLstrTipo.Length = 128 Then
                    VLstrDatos = VLstrTipo.Split(VLstrSeparador, StringSplitOptions.None)
                    If VLstrDatos.Length >= 6 Then
                        _Configuracion = VLstrDatos(4) & " [S/C]"
                        'If VLstrDatos(0) = My.Application.Info.AssemblyName Then
                        If VLstrDatos(0) = "KMGuiaHTTP" Then ' "KM" & My.Application.Info.AssemblyName Then
                            If IsNumeric(VLstrDatos(1)) Then
                                '1=Replicacion,2=Filtros,3=Recuperaciones,4=Linea
                                Select Case CType(VLstrDatos(1), TipoGuia)
                                    Case TipoGuia.Filtros
                                        VGtipGeneral.Tipo = TipoGuia.Filtros
                                        _TipoGuia = 2
                                    Case TipoGuia.Linea
                                        VGtipGeneral.Tipo = TipoGuia.Linea
                                        _TipoGuia = 4
                                    Case TipoGuia.Recuperacion
                                        VGtipGeneral.Tipo = TipoGuia.Recuperacion
                                        _TipoGuia = 3
                                    Case TipoGuia.Replicacion
                                        VGtipGeneral.Tipo = TipoGuia.Replicacion
                                        _TipoGuia = 1
                                    Case Else
                                        escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [1]" & vbNewLine & "[Aplicación incorrecta]", EventLogEntryType.Error)
                                        Me.ExitCode = 1
                                        Me.OnStop()
                                End Select
                                If VLstrDatos(2).Length = 4 Then
                                    If IsNumeric(VLstrDatos(2)) Then
                                        VGtipGeneral.Licencias = CInt(VLstrDatos(2))
                                        If IsNumeric(VLstrDatos(3)) Then
                                            If CInt(VLstrDatos(3)) = 1 Then
                                                VGtipGeneral.Licencias = VGtipGeneral.Licencias + 1
                                            End If
                                            VLstrConfiguracion = VLstrDatos(4).Trim & "\\"
                                            VGstrUsuario = VLstrDatos(5).Trim
                                            'escribeVisorEventos("Configuracion=" & VLstrConfiguracion & ", Usuario=" & VGstrUsuario, EventLogEntryType.Information)
                                        Else
                                            escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [2]", EventLogEntryType.Error)
                                            ExitCode = 1
                                            OnStop()
                                        End If
                                    Else
                                        escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [3]", EventLogEntryType.Error)
                                        ExitCode = 1
                                        OnStop()
                                    End If
                                Else
                                    escribeVisorEventos("Se encontró el archivo de datos    pero es incorrecto o ha sido modificado [4]", EventLogEntryType.Error)
                                    ExitCode = 1
                                    OnStop()
                                End If
                            Else
                                escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [5]", EventLogEntryType.Error)
                                ExitCode = 1
                                OnStop()
                            End If
                        Else
                            escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [6]", EventLogEntryType.Error)
                            ExitCode = 1
                            OnStop()
                        End If
                    Else
                        escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [7]", EventLogEntryType.Error)
                        ExitCode = 1
                        OnStop()
                    End If
                Else
                    escribeVisorEventos("Se encontró el archivo de datos pero es incorrecto o ha sido modificado [8]", EventLogEntryType.Error)
                    ExitCode = 1
                    OnStop()
                End If

                _Clientes = 0
                If ObtenConfig() Then
                    If VLstrConfiguracion = VMConfiguracion Then
                        _Configuracion = VLstrConfiguracion
                        If ConfigurarAdmin() Then
                            If cargarGeneral() Then
                                If cargaBDCentral() Then
                                    'Si no encuentra información de la bitácora continúa operando
                                    If Not cargaBDBitacora() Then
                                        VGtipBDBitacora = Nothing
                                        VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "Error al leer la información de la base de datos de bitácora " & VMstrError & ETX
                                        Notifica("Error al leer la información de la base de datos de bitácora " & VMstrError, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                        VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                    End If

                                    'Se conecta con la base de datos central
                                    If abrirBase(TipoBD.Central) Then
                                        If Not IsNothing(VGtipBDBitacora) Then
                                            If Not abrirBase(TipoBD.Bitacora) Then
                                                VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se pudo conectar con la base de datos de bitácora" & VMstrError & ETX
                                                VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                            End If
                                        End If
                                        VMobjClientes = New Hashtable
                                        If VGtipGeneral.Tipo = TipoGuia.Filtros Then
                                            If cargarPuertos() Then
                                                Dim VLhasPuerto As System.Collections.DictionaryEntry
                                                For Each VLhasPuerto In VMhasPuertos
                                                    Dim VLobjGuia As GuiaFiltros
                                                    Notifica("--> Levantando Servidor en el puerto " & VGtipGeneral.Puerto, Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
                                                    VLobjGuia = New GuiaFiltros(Me)
                                                    VLobjGuia.Layout = VGtipLayout.LayOut ' VLobjLicencias(0)(VLintI)
                                                    VLobjGuia.Formato = VGtipLayout.Formato  'VLobjLicencias(1)(VLintI)
                                                    VLobjGuia.Puerto = VLhasPuerto.Key ' VLobjLicencias(2)(VLintI)
                                                    VGtipGeneral.PuertoLocal = VLobjGuia.Puerto
                                                    VLobjGuia.NumFiltro = VLhasPuerto.Value ' VLobjLicencias(3)(VLintI)
                                                    VLobjGuia.BaseCentral = VGtipBDCentral
                                                    VLobjGuia.LongitudEncabezado = VGtipGeneral.LongitudE

                                                    Dim VLblnServidorOK As Boolean
                                                    VLblnServidorOK = VLobjGuia.LevantaServidor
                                                    If VLblnServidorOK Then
                                                        Notifica("   Escuchando en el puerto " & VLhasPuerto.Key, Me, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
                                                        VMobjClientes.Add(VMobjClientes.Count + 1, VLobjGuia)
                                                    Else
                                                        Notifica("   No se pudo escuchar en el Puerto " & VLhasPuerto.Key, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                        VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se pudo escuchar en el Puerto " & VGtipGeneral.PuertoLocal & ETX
                                                        VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                                    End If
                                                Next
                                            Else
                                                VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & VMstrError & ETX
                                                VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                                Threading.Thread.Sleep(100)
                                                ExitCode = 1
                                                OnStop()
                                            End If
                                        Else
                                            Dim VLintI As Integer

                                            For VLintI = 0 To VGtipGeneral.Licencias - 1
                                                Dim VLobjGuia As Guia
                                                Dim VLblnServidorOK As Boolean

                                                VLobjGuia = New Guia(Me)

                                                Notifica("--> Levantando Servidor en el puerto " & VGtipGeneral.PuertoLocal, Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)

                                                VLobjGuia.Puerto = VGtipGeneral.PuertoLocal
                                                VLblnServidorOK = VLobjGuia.LevantaServidor
                                                If VLblnServidorOK Then
                                                    Notifica("   Escuchando en el puerto " & VGtipGeneral.PuertoLocal, Me, enuEscribeLog.Completo, enuTipoAviso.Informativo)
                                                    VMobjClientes.Add(VMobjClientes.Count + 1, VLobjGuia)
                                                Else
                                                    Notifica("   No se pudo escuchar en el Puerto " & VGtipGeneral.PuertoLocal & "  [" & VLobjGuia.UltimoError & "]", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se pudo escuchar en el Puerto " & VGtipGeneral.PuertoLocal & "[" & VLobjGuia.UltimoError & "]" & ETX
                                                    VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                                End If
                                                VGtipGeneral.PuertoLocal += 1
                                            Next
                                        End If

                                        Notifica("--> Conectando con servidor transaccional", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno, "  Servidor: " & VGtipGeneral.IP & ": " & VGtipGeneral.Puerto)
                                        VMtcpCliente = New Cliente_TCP
                                        VMtcpCliente.Puerto = VGtipGeneral.Puerto
                                        VMtcpCliente.DireccionIP = VGtipGeneral.IP
                                        VMtcpCliente.Conectar()
                                        VMblnSigue = False
                                        Do
                                            Thread.Sleep(1000)
                                        Loop While Not VMblnSigue
                                        If VMtcpCliente.Estado = Cliente_TCP.EstadoConexion.Conectado Then
                                            Inicia()
                                            VMtmrTiempoSinTran = New Timers.Timer
                                            VMtmrTiempoSinTran.Interval = 60000
                                            VMtmrTiempoSinTran.Enabled = True
                                        Else
                                            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & "Conexión con servidor remoto errónea" & ETX
                                            VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                            Notifica("  Conexión con servidor remoto errónea ", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError, vbNewLine & "<-- Finaliza conexión con servidor remoto")
                                            ExitCode = 1
                                            OnStop()
                                        End If
                                    Else
                                        VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & "Error conectar con la base de datos central " & VMstrError & ETX
                                        VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                        Notifica("Error conectar con la base de datos central " & VMstrError, Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
                                        Threading.Thread.Sleep(100)
                                        ExitCode = 1
                                        OnStop()
                                    End If
                                Else
                                    VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & "Error al leer la información de la base de datos central " & VMstrError & ETX
                                    VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                    Notifica("Error al leer la información de la base de datos central " & VMstrError, Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
                                    Threading.Thread.Sleep(100)
                                    ExitCode = 1
                                    OnStop()
                                End If
                            Else
                                'Manda el error a la consola
                                VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & "Error al leer la información de configuración " & VMstrError & ETX
                                VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                                Notifica("Error al leer la información de configuración " & VMstrError, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                Threading.Thread.Sleep(100)
                                ExitCode = 1
                                OnStop()
                            End If
                        Else
                            ExitCode = 1
                            OnStop()
                        End If
                    Else
                        Notifica("La configuración del archivo KSGuia no corresponde con el KMConfig", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
                        ExitCode = 1
                        OnStop()
                    End If
                Else
                    'La notificación se envió en la función ObtenConfig()
                    ExitCode = 1
                    OnStop()
                End If
            Catch ex As Exception
                escribeVisorEventos("Archivo de datos incorrecto" & ex.Message, EventLogEntryType.Error)
                ExitCode = 1
                OnStop()
            End Try
        Catch ex As Exception
            escribeVisorEventos("No se encontro el Archivo que especifica el tipo de Guia", EventLogEntryType.Error)
            ExitCode = 1
            OnStop()
        End Try
    End Sub
    Private Function ObtenConfig() As Boolean
        Dim VLstrRuta As String
        Dim VMioArchivo As StreamReader
        VLstrRuta = My.Application.Info.DirectoryPath

        Try
leeArchivo:

            VLstrRuta = StrReverse(VLstrRuta)
            VLstrRuta = VLstrRuta.Substring(VLstrRuta.IndexOf("\") + 1, VLstrRuta.Length - VLstrRuta.IndexOf("\") - 1)
            VLstrRuta = StrReverse(VLstrRuta)

            Dim VLstrLinea As String
            Dim VLintRenglon As Integer = 1
            VMioArchivo = New StreamReader(VLstrRuta & "\KMConfig.dat")
            While Not VMioArchivo.EndOfStream
                VLstrLinea = VMioArchivo.ReadLine
                If VLintRenglon = 2 Then
                    Dim VLstrConfiguracion() As String
                    VLstrConfiguracion = Split(VMobjRc4.RC4Decrypt(VLstrLinea, "KS%Soluciones%Configuracion"), "@|")
                    VMConfiguracion = VLstrConfiguracion(1)
                    VMConfiguracion = VMConfiguracion & "\\"
                    'VGtipGeneral.Licencias = VLstrConfiguracion(3)
                    ObtenConfig = True
                    Exit While
                End If
                VLintRenglon += 1
            End While
            VMioArchivo.Close()
        Catch ex As Exception
            ObtenConfig = False
            If ex.Message.Contains("No se pudo encontrar el") Then
                GoTo leeArchivo
            Else
                Notifica("No se pudo encontrar el Archivo de Configuracion", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
            End If
        End Try
    End Function
    Private Function ValidaLicencias(ByRef VLobjLicencias() As Hashtable) As String

        '***********************************************************
        'Aqui se van a comparar las licencias, 
        'y asignar puertos por lay
        '**********************************************************

        ValidaLicencias = ""
        'VMstrLicencias = "3"
        Dim VLdbsBase As Object
        Dim VLcmdComando As Object
        Dim VLrcsDatos As Object
        Dim VLstrSQL As String
        VLcmdComando = Nothing

        VLstrSQL = "select * from tconalarfiltros where id_lay = " & VGtipLayout.LayOut

        Select Case VGtipBDCentral.Tipo
            Case TipoBase.MySQL
                VLdbsBase = New MySqlConnection
                VLcmdComando = New MySqlCommand
            Case TipoBase.SQLServer
                VLdbsBase = New SqlConnection
                VLcmdComando = New SqlCommand
            Case TipoBase.Oracle
                VLdbsBase = New OracleConnection
                VLcmdComando = New OracleCommand
            Case Else
                VLdbsBase = Nothing
                VLcmdComando = Nothing
        End Select

        Try
            VLdbsBase.ConnectionString = cadenaConexion(VGtipBDCentral)
            VLdbsBase.Open()
            VLcmdComando.Connection = VLdbsBase
            VLcmdComando.CommandText = VLstrSQL
            Try
                VLrcsDatos = VLcmdComando.ExecuteReader
                If VLrcsDatos.HasRows Then
                    Dim VLinti As Integer = 1
                    While VLrcsDatos.read()
                        VLobjLicencias(0).Add(VLinti, VLrcsDatos!Id_Lay)
                        VLobjLicencias(1).Add(VLinti, VLrcsDatos!Id_Formato)
                        VLobjLicencias(2).Add(VLinti, VLrcsDatos!Puerto)
                        VLobjLicencias(3).Add(VLinti, VLrcsDatos!Num_Filtro)
                        VLinti += 1
                    End While
                End If
            Catch ex As MySqlException
                ValidaLicencias = ex.Message
            Catch ex As OracleException
                ValidaLicencias = ex.Message
            Catch ex As SqlClient.SqlException
                ValidaLicencias = ex.Message
            End Try
            VLdbsBase.Close()
        Catch ex As MySqlException
            ValidaLicencias = ex.Message
        Catch ex As OracleException
            ValidaLicencias = ex.Message
        Catch ex As SqlClient.SqlException
            ValidaLicencias = ex.Message
        End Try
    End Function
    Private Sub VMtcpCliente_ConexionTerminada(ByVal DesconexionEco As Boolean) Handles VMtcpCliente.ConexionTerminada
        If DesconexionEco = True Then
            If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                escribeLog("---------------------------" & vbNewLine & "Se perdio conexión con el servidor transaccional" & vbNewLine & "---------------------------")
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("---------------------------" & vbNewLine & "Se perdio conexión con el servidor transaccional" & vbNewLine & "---------------------------")
            End If
            escribeVisorEventos("---------------------------" & vbNewLine & "Se perdio conexión con el servidor transaccional" & vbNewLine & "---------------------------", EventLogEntryType.Error)
        End If

    End Sub
    Private Sub VMtcpCliente_ErrorConexion(ByVal Mensaje As String) Handles VMtcpCliente.ErrorConexion
        Dim VLstrRespuesta As String

        If VMblnListo1 Then
            If VMbytIntentaReconexion = 0 Then
                Notifica("Se ha perdido la comunicación con el servidor transaccional " & vbNewLine & Mensaje, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True)

                VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & "Se perdió conexión con servidor transaccional" & Mensaje & ETX
                Try
                    VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                Catch ex As Exception
                    Notifica("No se pudo enviar el mensaje a la consola central " & vbNewLine & "Mensaje original: " & Mensaje, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True)
                End Try

                'Inicia el timer de reconexión
                If Not VMblnReconectando Then
                    VMtmrReconecta.Interval = 10000
                    VMtmrReconecta.Enabled = True
                End If
            End If
            VMblnSigue = True
        Else
            If Not IsNothing(VMtcpAdmin) Then
                If VMtcpAdmin.Estado = Cliente_TCP.EstadoConexion.Conectado Then
                    VLstrRespuesta = STX & 7014 & FS & ACK & FS & "02" & FS & Mensaje & ETX
                    Try
                        VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                        Threading.Thread.Sleep(100)
                    Catch ex As Exception
                        Notifica("No se pudo enviar el mensaje a la consola central" & vbNewLine & "Mensaje original: " & Mensaje, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True)
                    End Try
                End If
            End If
            Notifica("  Conexión con servidor remoto errónea (" & Mensaje.Replace(vbNewLine, "  " & vbNewLine) & ")" & vbNewLine & "<-- Finaliza conexión con servidor remoto", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
            VMtmrReconecta.Interval = 60000
            VMtmrReconecta.Enabled = True
            'End
        End If
    End Sub
    Private Sub VMtcpCliente_ConexionEstablecida() Handles VMtcpCliente.ConexionEstablecida
        VMblnSigue = True
        If _Log = enuDepuracion.Completa Then
            escribeLog("  Respuesta recibida")
        ElseIf _Log = enuDepuracion.Consola Then
            escribeConsola("  Respuesta recibida")
        End If
    End Sub
    Private Sub VMtcpCliente_DatosRecibidos(ByVal Mensaje As String) Handles VMtcpCliente.DatosRecibidos
        VMintBloques += 1
        VMstrUltimoBloque = Format(Date.Now, "dd/MM/yyyy HH:mm:ss")
        If VPblnLogtran Then
            Try
                VPioLogTran.Write(Mensaje)
            Catch ex As Exception
                VMtcpAdmin.EnviarDatos(STX & "7014" & FS & ACK & FS & "01" & FS & "No se pudo insertar en el archivo transaccional " & ex.Message & ETX)
            End Try
        End If
        If Not VMtmrTiempoSinTran Is Nothing Then
            VMtmrTiempoSinTran.Enabled = False
            VMtmrTiempoSinTran.Enabled = True
            VMintSinTransaccion = 0
        End If
        If VGtipGeneral.Acknowled = True Then
            VMtcpCliente.EnviarDatos("#")
        End If
        SyncLock VMobjBloques
            VMobjBloques.Enqueue(Mensaje)
        End SyncLock

        If Not _Procesando Then
            VMtmrProcesa.Enabled = True
            'VMobjHilo = New Threading.Thread(AddressOf procesadatos)
            'VMobjHilo.Start()
        Else
            'Hay que validar correctamente este evento
            VMtmrProcesa.Enabled = True
            '_Procesando = True
            'If Not VMobjHilo.IsAlive Then
            '    VMobjHilo = New Threading.Thread(AddressOf procesadatos)
            '    VMobjHilo.Start()
            'End If
        End If
        'If Not VMtmrTiempoSinTran Is Nothing Then
        '    VMtmrTiempoSinTran.Enabled = False
        'End If
        VMblnListo1 = True
    End Sub
    Private Sub procesadatos()
        Dim VLstrMensaje As String = ""
        _Procesando = True
        Try

            SyncLock VMobjBloques
                While VMobjBloques.Count > 0
                    VLstrMensaje = VMobjBloques.Dequeue
                    fncCargarBloque(VLstrMensaje)
                End While
            End SyncLock
        Catch ex As Exception
            If _Log = enuDepuracion.Completa Then
                escribeLog("  Error al procesar los bloques " & vbNewLine & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("  Error al procesar los bloques " & vbNewLine & ex.Message)
            End If
            escribeVisorEventos("Error al procesar los bloques" & vbNewLine & ex.Message, EventLogEntryType.Error)
        End Try
        _Procesando = False
    End Sub
    Private Sub fncCargarBloque(ByVal mensaje As String)
        Dim VLblnEnvio As Boolean = False
        Try
            Dim VLobjClientes As DictionaryEntry
            For Each VLobjClientes In VMobjClientes
                If VLobjClientes.Value.usuariosconectados > 0 Then
                    VLobjClientes.Value.enviar(mensaje)
                    VLblnEnvio = True
                End If
            Next

            If VLblnEnvio Then
                VMintProcesados += 1
            Else
                VMintIgnorados += 1
            End If
        Catch ex As Exception
            Notifica("  Error al entregar los bloques a las clases " & vbNewLine & ex.Message, Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
            'Informar a la consola y detallar mejor el puerto en el que no se pudo entregar la información
        End Try
    End Sub
    Private Sub Inicia()
        Dim VLstrRespuesta As String

        Try
            Notifica("--> Enviando solicitud de transacciones", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
            solicitaEnvioTransacciones("SENDTRAN")
        Catch ex As Exception
            Notifica("<-- Error al enviar [SENDTRAN]:" & ex.Message, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
        End Try

        Try
            If VGtipGeneral.MinDia = 0 Or VGtipGeneral.MinNoche = 0 Then
                VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se reiniciará por ausencia de transacciones, ya que en la configuración no se indicó un tiempo específico sin transacciones" & ETX
                VMtcpAdmin.EnviarDatos(VLstrRespuesta)

                Notifica("No se reiniciará por ausencia de transacciones, ya que en la configuración no se indicó un tiempo específico sin transacciones", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
            Else
                'VMtmrCambiaTmr = New Timers.Timer
                'VMtmrCambiaTmr.Enabled = True
            End If
        Catch ex As Exception
            VGtipGeneral.MinDia = 0
            VGtipGeneral.MinNoche = 0
            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "No se reiniciará por ausencia de transacciones, ya que no se pudo determinar el tiempo de reinicio sin transacciones: " & ex.Message & ETX
            VMtcpAdmin.EnviarDatos(VLstrRespuesta)
            Notifica("No se reiniciará por ausencia de transacciones, ya que no se pudo determinar el tiempo de reinicio sin transacciones: " & ex.Message, Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
        End Try
    End Sub
    Private Sub solicitaEnvioTransacciones(ByVal PstrServicio As String, Optional ByVal VPstrRecPtrRP As String = "")
        Dim Log As Integer

        Try

            VGstrLocalHostName = My.Computer.Name
            If VGstrLocalHostName.Length > 14 Then
                If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                    escribeLog("   Se recortara el nombre de la PC=" & VGstrLocalHostName)
                ElseIf _Log = enuDepuracion.Consola Then
                    escribeLog("   Se recortara el nombre de la PC=" & VGstrLocalHostName)
                End If
                VGstrLocalHostName = VGstrLocalHostName.Substring(0, 14)
            End If
            If VGstrUsuario.Length > 5 Then
                If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                    escribeLog("   Se recortara el Usuario=" & VGstrUsuario)
                ElseIf _Log = enuDepuracion.Consola Then
                    escribeLog("   Se recortara el Usuario=" & VGstrUsuario)
                End If
                VGstrUsuario = VGstrUsuario.Substring(0, 5)
            End If
            VGstrLocalIPAddress = RetIPAddress("localhost")
            If VGtipGeneral.Tipo = TipoGuia.Linea Or VGtipGeneral.Tipo = TipoGuia.Filtros Then
                VGintAplicacion = 10
            ElseIf VGtipGeneral.Tipo = TipoGuia.Recuperacion Then
                VGintAplicacion = 12
            End If
            Dim VGintTipoCnx As Integer = 0
            Dim VLstrBuferDeSolicitud As String = ""
            Dim VLstrIPAddress As String = Space(15)
            Dim VLstrPCName As String = Space(14)
            Dim VLstrUsr As String = Space(6)
            Dim VLstrUsuario As String = Space(55)
            Dim VLstrWSE As String
            Dim VLstrApp As String = ""
            Dim VLEnviaLongitud As String
            VMmsgEnviaTransacciones.Pathway = Space(16)
            VMmsgEnviaTransacciones.Servidor = Space(15)
            VMmsgEnviaTransacciones.Longitud = Space(5)
            Log = 6
            If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                escribeLog("   Creando Cadena a Enviar")
                escribeLog("   Usuario =" & VGstrUsuario)
                escribeLog("   Nombre PC =" & VGstrLocalHostName)
                escribeLog("   Dirección IP =" & VGstrLocalIPAddress)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeLog("   Creando Cadena a Enviar")
                escribeLog("   Usuario =" & VGstrUsuario)
                escribeLog("   Nombre PC =" & VGstrLocalHostName)
                escribeLog("   Dirección IP =" & VGstrLocalIPAddress)
            End If
            If VGtipGeneral.MensajeOrigen = enuServOrigen.Stratus Then
                Log = 8
                VMmsgEnviaTransacciones.Pathway = "$ENVIAR" & Space(Len(VMmsgEnviaTransacciones.Pathway) - Len("$ENVIAR"))
                VMmsgEnviaTransacciones.Servidor = PstrServicio & Space(Len(VMmsgEnviaTransacciones.Servidor) - Len(PstrServicio))
                VLstrIPAddress = VGstrLocalIPAddress & Space(VLstrIPAddress.Length - VGstrLocalIPAddress.Length)
                Log = 9
                VLstrPCName = VGstrLocalHostName & Space(VLstrPCName.Length - VGstrLocalHostName.Length)
                Log = 10
                VLstrUsr = Format(CInt(VGstrUsuario), "00000")
                Log = 11
                If VGtipGeneral.Tipo = TipoGuia.Linea Then
                    VLstrApp = 10
                ElseIf VGtipGeneral.Tipo = TipoGuia.Recuperacion Then
                    VLstrApp = 12
                ElseIf VGtipGeneral.Tipo = TipoGuia.Filtros Then
                    VLstrApp = 14
                End If
                VLEnviaLongitud = Len(VMmsgEnviaTransacciones.Pathway) + Len(VMmsgEnviaTransacciones.Servidor) + _
                                  Len(VMmsgEnviaTransacciones.Longitud) + Len(VLstrIPAddress) + _
                                  Len(VLstrPCName) + Len(VLstrUsr) + 8
                Log = 13
                If Not VGtipGeneral.Tipo = TipoGuia.Linea Then
                    VLEnviaLongitud = VLEnviaLongitud + 68
                End If

                VMmsgEnviaTransacciones.Longitud = VLEnviaLongitud
                Log = 14
                VLstrBuferDeSolicitud = VMmsgEnviaTransacciones.Pathway & VMmsgEnviaTransacciones.Servidor & Format(CInt(VMmsgEnviaTransacciones.Longitud), "00000") & _
                                        VLstrIPAddress & VLstrPCName & VLstrUsr & " " & VLstrApp
                If VGtipGeneral.Tipo = TipoGuia.Linea Then
                    VLstrBuferDeSolicitud = VLstrBuferDeSolicitud & Format(VGtipGeneral.Confirmaciones, "00000")
                End If
                Log = 15
            ElseIf VGtipGeneral.MensajeOrigen = enuServOrigen.Estandar Then
                VMmsgEnviaTransacciones.Pathway = "$ENVIAR" & Space(Len(VMmsgEnviaTransacciones.Pathway) - Len("$ENVIAR"))
                VMmsgEnviaTransacciones.Servidor = PstrServicio & Space(Len(VMmsgEnviaTransacciones.Servidor) - Len(PstrServicio))
                Log = 16
                VLEnviaLongitud = (Len(VMmsgEnviaTransacciones.Longitud + VMmsgEnviaTransacciones.Pathway + VMmsgEnviaTransacciones.Servidor)) + Len(VLstrIPAddress) + Len(VLstrPCName) + Len(VLstrUsr) + 2
                VMmsgEnviaTransacciones.Longitud = VLEnviaLongitud & Space(Len(VMmsgEnviaTransacciones.Longitud) - VLEnviaLongitud.Length)
                Log = 17
                VLstrIPAddress = VGstrLocalIPAddress & Space(VLstrIPAddress.Length - VGstrLocalIPAddress.Length)
                Log = 18
                VLstrPCName = VGstrLocalHostName & Space(VLstrPCName.Length - VGstrLocalHostName.Length)
                Log = 19
                VLstrUsr = VGstrUsuario & Space(VLstrUsr.Length - VGstrUsuario.Length)
                Log = 20
                If VGtipGeneral.Mensaje = enuTipoMensaje.Corto Then
                    VLEnviaLongitud = (Len(VMmsgEnviaTransacciones.Longitud + VMmsgEnviaTransacciones.Pathway + VMmsgEnviaTransacciones.Servidor)) + Len(VLstrIPAddress) + Len(VLstrPCName)
                    VMmsgEnviaTransacciones.Longitud = VLEnviaLongitud & Space(Len(VMmsgEnviaTransacciones.Longitud) - VLEnviaLongitud.Length)
                    VLstrBuferDeSolicitud = VMmsgEnviaTransacciones.Pathway & VMmsgEnviaTransacciones.Servidor _
                                        & VMmsgEnviaTransacciones.Longitud & VLstrIPAddress & VLstrPCName & VLstrUsr & Strings.Format(CInt(VGintAplicacion), "00")
                    Log = 21
                ElseIf VGtipGeneral.Mensaje = enuTipoMensaje.Normal Then
                    VLstrWSE = VLstrIPAddress & VLstrPCName
                    VLstrWSE = VLstrWSE & Space(218 - Len(VLstrWSE)) & VLstrUsr
                    Log = 22
                    If VGintTipoCnx = 158 Then
                        VMmsgEnviaTransacciones.Pathway = "$REENVIAR" & Space(Len(VMmsgEnviaTransacciones.Pathway) - Len("$REENVIAR"))
                    End If
                    VLstrBuferDeSolicitud = VMmsgEnviaTransacciones.Pathway & VMmsgEnviaTransacciones.Servidor _
                                        & VMmsgEnviaTransacciones.Longitud & VLstrWSE
                    Log = 23
                End If
            End If
            If _Log = enuDepuracion.Completa Then
                escribeLog("   Se esta enviando solucitud de transacciones " & vbNewLine & "   Cadena:" & VLstrBuferDeSolicitud)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("   Se esta enviando solucitud de transacciones " & vbNewLine & "   Cadena:" & VLstrBuferDeSolicitud)
            End If
            SolicitudEnviada(PstrServicio, VLstrBuferDeSolicitud, VLstrBuferDeSolicitud.Length, 12)
        Catch ex As Exception
            If _Log = enuDepuracion.Completa Then
                escribeLog("   Error al Crear la cadena de envio: " & Log & "  Mensaje" & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("   Error al Crear la cadena de envio: " & Log & "  Mensaje" & ex.Message)
            End If
        End Try
    End Sub
    Public Function SolicitudEnviada(ByVal PstrServidor As String, ByVal PstrBufer As String, ByVal PintTamaño As Integer, ByVal PvarId As Object) As Integer
        Try
            VMtcpCliente.EnviarDatos(PstrBufer)
            If _Log = enuDepuracion.Completa Then
                escribeLog("   Se envio solucitud de transacciones con Exito")
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("   Se envio solucitud de transacciones con Exito")
            End If
        Catch ex As Exception
            If _Log = enuDepuracion.Completa Then
                escribeLog("   Error al Enviar: " & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("   Error al Enviar: " & ex.Message)
            End If
        End Try

    End Function
    Private Sub CalculaTiempo(ByVal Hora1 As DateTime, ByVal Hora2 As DateTime, Optional ByVal VPCambioDia As Boolean = False)
        'Hora2 es la mas alta
        Dim VLtmTiempo As TimeSpan
        Dim VLdblTmilis As Double

        Dim VLintHora As Integer
        Dim VLintMin As Integer
        Dim VLintSeg As Integer
        Dim VLTiempoFinDia As Integer
        Try
            If VPCambioDia = True Then
                'Se calcula primero el fin del dia, para despues agregar lo que falta
                VLintHora = 23 - Hora1.Hour
                VLintMin = 59 - Hora1.Minute
                VLintSeg = 59 - Hora1.Second
                VLTiempoFinDia = ((VLintHora * 60 + VLintMin) * 60 + VLintSeg) * 1000
                Hora1 = "00:00"
            End If
            'VLtmTiempo = Hora2.Subtract(Hora1).Duration
            VLtmTiempo = Hora2 - Hora1
            VLdblTmilis = ((VLtmTiempo.Hours * 60) * 60) + (VLtmTiempo.Minutes * 60) + VLtmTiempo.Seconds
            VLdblTmilis = (VLdblTmilis * 1000) + VLTiempoFinDia
            VMtmrCambiaTmr.Interval = VLdblTmilis
            VMtmrCambiaTmr.Enabled = True
            If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                escribeLog("  Se Realiza el cambio de Horario en: " & VLtmTiempo.Hours & ":" & VLtmTiempo.Minutes & ":" & VLtmTiempo.Seconds)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("  Se Realiza el cambio de Horario en: " & VLtmTiempo.Hours & ":" & VLtmTiempo.Minutes & ":" & VLtmTiempo.Seconds)
            End If
            escribeVisorEventos("  Se Realiza el cambio de Horario en: " & VLtmTiempo.Hours & ":" & VLtmTiempo.Minutes & ":" & VLtmTiempo.Seconds, EventLogEntryType.Information)
        Catch ex As Exception
            If _Log = enuDepuracion.Sencilla Or _Log = enuDepuracion.Completa Then
                escribeLog("  Error al obtrener la información del cambio de Horario " & vbNewLine & ex.Message)
            ElseIf _Log = enuDepuracion.Consola Then
                escribeConsola("  Error al obtrener la información del cambio de Horario " & vbNewLine & ex.Message)
            End If
            escribeVisorEventos("  Error al obtrener la información del cambio de Horario " & vbNewLine & ex.Message, EventLogEntryType.Information)
        End Try
    End Sub
    Private Sub VMtmrTiempoSinTran_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrTiempoSinTran.Elapsed
        Dim VLdatHoraActual As Date
        Dim VLdatInicioDia As Date
        Dim VLdatInicioNoche As Date

        'Se conecta con el servidor transaccional

        'If Not VMtmrTiempoSinTran Is Nothing Then
        VMtmrTiempoSinTran.Enabled = False

        VMintSinTransaccion += 1

        VLdatHoraActual = TimeOfDay
        VLdatInicioDia = CDate(VGtipGeneral.InicioDia)
        VLdatInicioNoche = CDate(VGtipGeneral.InicioNoche)

        'Verifica el horario en el que se encuentra
        If VLdatHoraActual >= VLdatInicioDia And VLdatHoraActual < VLdatInicioNoche Then
            'Horario matutino
            If VMintSinTransaccion >= VGtipGeneral.MinDia Then
                'reinicia
                ReinciaSvrTxn("Matutino")
            Else
                VMtmrTiempoSinTran.Enabled = True
            End If
        Else
            'Horario vespertino
            If VMintSinTransaccion >= VGtipGeneral.MinNoche Then
                'reinicia
                ReinciaSvrTxn("Vespertino")
            Else
                VMtmrTiempoSinTran.Enabled = True
            End If
        End If
    End Sub
#Region "Configuracion"
    Private Function cargarGeneral() As Boolean
        '***********************************************************************************'
        '   Carga todos los datos necesarios para la configuracion general                  '
        '                                                                                   '
        '   Devuelve: TRUE si carga todos los datos que se encuntran en el registro         '
        '             FALSE si no carga todos los datos que se encuntran en el registro     '
        '***********************************************************************************'

        Dim VLkeyRegistro As RegistryKey
        Dim VLstrClave As String
        Dim VLstrRuta As String

        Select Case VGtipGeneral.Tipo
            Case TipoGuia.Filtros
                VLstrClave = "\\GuiaFx"
            Case TipoGuia.Linea
                VLstrClave = "\\GuiaEnLinea"
            Case TipoGuia.Recuperacion
                VLstrClave = "\\GuiaRecuperacion"
            Case Else
                VLstrClave = "\\GuiaReplicacion"
        End Select

        VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\Layout"
        Try
            Notifica("--> Obteniendo información de configuración general", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
            VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)

            If Not (VLkeyRegistro Is Nothing) Then
                Try
                    VGtipLayout.LayOut = VLkeyRegistro.GetValue("Layout")
                    VGtipLayout.Formato = VLkeyRegistro.GetValue("Formato")
                    VLkeyRegistro.Close()

                    VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\General"
                    Try
                        VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)
                        If Not (VLkeyRegistro Is Nothing) Then
                            Try
                                VGtipGeneral.MensajeOrigen = VLkeyRegistro.GetValue("Mensaje")
                                VGtipGeneral.Mensaje = VLkeyRegistro.GetValue("TipoMsj")
                                Try
                                    VGtipGeneral.InicioDia = VLkeyRegistro.GetValue("TM")
                                    VGtipGeneral.InicioNoche = VLkeyRegistro.GetValue("TV")
                                Catch ex As Exception

                                End Try
                                VGtipGeneral.MinDia = VLkeyRegistro.GetValue("MinDia")
                                VGtipGeneral.MinNoche = VLkeyRegistro.GetValue("MinNoche")
                                VGtipGeneral.Acknowled = VLkeyRegistro.GetValue("ECHO")
                                VGtipGeneral.LongitudE = VLkeyRegistro.GetValue("LongitudE")
                                VLkeyRegistro.Close()

                                VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\Guias" & VLstrClave
                                Try
                                    VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)
                                    If Not (VLkeyRegistro Is Nothing) Then
                                        Try
                                            VGtipGeneral.Puerto = VLkeyRegistro.GetValue("PuertoTran")
                                            VGtipGeneral.IP = VLkeyRegistro.GetValue("ServidorTran")
                                            VGtipGeneral.PuertoLocal = VLkeyRegistro.GetValue("PuertoEscucha")
                                            VLkeyRegistro.Close()

                                            With VGtipGeneral
                                                Notifica("  Layout: " & VGtipLayout.LayOut, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Formato: " & VGtipLayout.Formato, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Longitud encabezado: " & .LongitudE, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Tiempo Sin Transacciones M: " & .MinDia, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Inicio Vespertino: " & .InicioNoche, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Tiempo Sin Transacciones V: " & .MinNoche, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Tipo Guia: " & .IP, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Servidor Transaccional: " & .Tipo, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Puerto: " & .Puerto, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Notifica("  Puerto Escucha: " & .PuertoLocal, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                If .Mensaje = enuTipoMensaje.Corto Then
                                                    Notifica("  Tipo Mensaje: " & "Corto", Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Else
                                                    Notifica("  Tipo Mensaje: " & "Normal", Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                End If
                                                If .MensajeOrigen = enuServOrigen.Estandar Then
                                                    Notifica("  Servidor Origen: " & "Estandar", Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                Else
                                                    Notifica("  Servidor Origen: " & "Stratus", Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                                                End If
                                            End With

                                            'Carga la configuración del correo
                                            VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\Correo"
                                            Try
                                                VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)
                                                If Not IsNothing(VLkeyRegistro) Then
                                                    Try
                                                        With VGtipCorreo
                                                            .Activo = CBool(VLkeyRegistro.GetValue("Activo"))
                                                            If .Activo Then
                                                                .Servidor = VLkeyRegistro.GetValue("Servidor")
                                                                .Cuenta = VLkeyRegistro.GetValue("Correo")

                                                                If Not IsNothing(.Servidor) And Not IsNothing(.Cuenta) Then
                                                                    If .Servidor.Trim.Length > 0 And .Cuenta.Trim.Length > 0 Then
                                                                        .Autenticacion = CBool(VLkeyRegistro.GetValue("Autenticacion"))
                                                                        If .Autenticacion Then
                                                                            .Usuario = VLkeyRegistro.GetValue("Usuario")
                                                                            .Contrasena = VLkeyRegistro.GetValue("Contrasena")

                                                                            If Not IsNothing(.Usuario) And Not IsNothing(.Contrasena) Then
                                                                                If .Contrasena.Trim.Length > 0 Then
                                                                                    Try
                                                                                        .Contrasena = desencriptaValor(.Contrasena)
                                                                                    Catch ex As Exception
                                                                                        .Activo = False
                                                                                        Notifica("     El envío de correo está configurado para autenticarse con el servidor de salida pero la información está incompleta o errónea [Contraseña], vuelva a configurar el envío de correo para solucionar el problema. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                                                    End Try
                                                                                End If
                                                                            Else
                                                                                .Activo = False
                                                                                Notifica("     El envío de correo está configurado para autenticarse con el servidor de salida pero la información está incompleta o errónea, vuelva a configurar el envío de correo para solucionar el problema. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                                            End If
                                                                        End If

                                                                        If .Activo Then 'Si no hubo error en la autenticación
                                                                            'Carga las cuentas destino
                                                                            .Para = VLkeyRegistro.GetValue("Para")
                                                                            .CC = VLkeyRegistro.GetValue("CC")
                                                                            .CO = VLkeyRegistro.GetValue("CCO")

                                                                            If Not IsNothing(.Para) Then
                                                                                If .Para.Trim.Length > 0 Then
                                                                                    .Para = .Para.Substring(1)
                                                                                End If
                                                                            Else
                                                                                .Para = ""
                                                                            End If

                                                                            If Not IsNothing(.CC) Then
                                                                                If .CC.Trim.Length > 0 Then
                                                                                    .CC = .CC.Substring(1)
                                                                                End If
                                                                            Else
                                                                                .CC = ""
                                                                            End If

                                                                            If Not IsNothing(.CO) Then
                                                                                If .CO.Trim.Length > 0 Then
                                                                                    .CO = .CO.Substring(1)
                                                                                End If
                                                                            Else
                                                                                .CO = ""
                                                                            End If

                                                                            Notifica("     Envío de correos activo", Me, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
                                                                        End If
                                                                    Else
                                                                        Notifica("     El envío de correo está activo pero la configuración está incompleta o errónea, vuelva a configurar el envío de correo para solucionar el problema. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                                    End If
                                                                Else
                                                                    Notifica("     El envío de correo está activo pero la configuración está incompleta o errónea, vuelva a configurar el envío de correo para solucionar el problema. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                                End If
                                                            Else
                                                                Notifica("     El envío de correo está desctivado. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                            End If
                                                        End With
                                                    Catch ex As System.Security.SecurityException
                                                        Notifica("     No se pudo obtener la conffiguración del correo [El usuario no tiene permisos suficientes para leer la clave del registro]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    Catch ex As ObjectDisposedException
                                                        Notifica("     No se pudo obtener la conffiguración del correo [El objeto no ha sido abierto o se ha cerrado anticipadamente]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    Catch ex As IOException
                                                        Notifica("     No se pudo obtener la conffiguración del correo [El objeto Registry que contiene el valor especificado se ha marcado para su eliminación]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    Catch ex As UnauthorizedAccessException
                                                        Notifica("     No se pudo obtener la conffiguración del correo [El usuario no tiene los derechos necesarios en el Registro]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    Catch ex As Exception
                                                        Notifica("     No se pudo obtener la conffiguración del correo [" & ex.Message & "]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                    End Try
                                                Else
                                                    Notifica("     No se encontró la configuración para el envío de correo. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                                End If
                                            Catch ex As ArgumentNullException
                                                Notifica("     No se pudo obtener la conffiguración del correo [No se ha especificado la llave a consultar]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                            Catch ex As ArgumentException
                                                Notifica("     No se pudo obtener la conffiguración del correo [El nombre de la llave excede el número de caracteres permitidos]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                            Catch ex As ObjectDisposedException
                                                Notifica("     No se pudo obtener la conffiguración del correo [El objeto no ha sido abierto o se ha cerrado anticipadamente]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                            Catch ex As Exception
                                                Notifica("     No se pudo obtener la conffiguración del correo [" & ex.Message & "]. No se recibirán notificaciones de desconexión", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia)
                                            End Try

                                            Notifica("<-- Finaliza lectura configuración general", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
                                            Return True
                                        Catch ex As System.Security.SecurityException
                                            VMstrError = "El usuario no tiene permisos suficientes para leer la clave del registro [Guias]"
                                        Catch ex As ObjectDisposedException
                                            VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [Guias]"
                                        Catch ex As IOException
                                            VMstrError = "El objeto Registry que contiene el valor especificado se ha marcado para su eliminación [Guias]"
                                        Catch ex As UnauthorizedAccessException
                                            VMstrError = "El usuario no tiene los derechos necesarios en el Registro [Guias]"
                                        Catch ex As Exception
                                            VMstrError = ex.Message & " [Guias]"
                                        End Try
                                    Else
                                        VMstrError = "No se ha encontrado la configuración [Guias]"
                                    End If
                                Catch ex As ArgumentNullException
                                    VMstrError = "No se ha especificado la llave a consultar [Guias]"
                                Catch ex As ArgumentException
                                    VMstrError = "El nombre de la llave excede el número de caracteres permitidos [Guias]"
                                Catch ex As ObjectDisposedException
                                    VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [Guias]"
                                Catch ex As Exception
                                    VMstrError = ex.Message & " [Guias]"
                                End Try
                            Catch ex As System.Security.SecurityException
                                VMstrError = "El usuario no tiene permisos suficientes para leer la clave del registro [General]"
                            Catch ex As ObjectDisposedException
                                VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [General]"
                            Catch ex As IOException
                                VMstrError = "El objeto Registry que contiene el valor especificado se ha marcado para su eliminación [General]"
                            Catch ex As UnauthorizedAccessException
                                VMstrError = "El usuario no tiene los derechos necesarios en el Registro [General]"
                            Catch ex As Exception
                                VMstrError = ex.Message & " [General]"
                            End Try
                        Else
                            VMstrError = "No se ha encontrado la configuración [General]"
                        End If
                    Catch ex As ArgumentNullException
                        VMstrError = "No se ha especificado la llave a consultar [General]"
                    Catch ex As ArgumentException
                        VMstrError = "El nombre de la llave excede el número de caracteres permitidos [General]"
                    Catch ex As ObjectDisposedException
                        VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [General]"
                    Catch ex As Exception
                        VMstrError = ex.Message & " [General]"
                    End Try
                Catch ex As System.Security.SecurityException
                    VMstrError = "El usuario no tiene permisos suficientes para leer la clave del registro [LayOut]"
                Catch ex As ObjectDisposedException
                    VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [LayOut]"
                Catch ex As IOException
                    VMstrError = "El objeto Registry que contiene el valor especificado se ha marcado para su eliminación [LayOut]"
                Catch ex As UnauthorizedAccessException
                    VMstrError = "El usuario no tiene los derechos necesarios en el Registro [LayOut]"
                Catch ex As Exception
                    VMstrError = ex.Message & " [LayOut]"
                End Try
            Else
                VMstrError = "No se ha encontrado la configuración [LayOut]"
            End If
        Catch ex As ArgumentNullException
            VMstrError = "No se ha especificado la llave a consultar [LayOut]"
        Catch ex As ArgumentException
            VMstrError = "El nombre de la llave excede el número de caracteres permitidos [LayOut]"
        Catch ex As ObjectDisposedException
            VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [LayOut]"
        Catch ex As Exception
            VMstrError = ex.Message & " [LayOut]"
        End Try
    End Function
    Private Function cargaBDCentral() As Boolean
        Dim VLkeyRegistro As RegistryKey
        Dim VLclsEncripta As New RC4
        Dim VLstrTemp As String
        Dim VLstrRuta

        VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\BD Central"
        Notifica("--> Obteniendo información de la base de datos central", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno, " " & VLstrRuta)

        Try
            VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)

            If Not IsNothing(VLkeyRegistro) Then
                Try
                    With VGtipBDCentral
                        .Base = VLkeyRegistro.GetValue("Base")
                        VLstrTemp = VLkeyRegistro.GetValue("Contrasena")
                        .Instancia = VLkeyRegistro.GetValue("Esquema")
                        .Puerto = VLkeyRegistro.GetValue("Puerto")
                        .Servidor = VLkeyRegistro.GetValue("Servidor")
                        .SI = Convert.ToBoolean(VLkeyRegistro.GetValue("SI"))
                        .Tipo = VLkeyRegistro.GetValue("Tipo")
                        .Usuario = VLkeyRegistro.GetValue("Usuario")
                        If VLstrTemp.Length > 0 Then
                            Try
                                .Contrasena = desencriptaValor(VLstrTemp)
                            Catch ex As Exception
                                VMstrError = "No se pudo obtener la contraseña de la base de datos central " & ex.Message
                                Return False
                            End Try
                        Else
                            .Contrasena = ""
                        End If

                        Notifica("  Tipo: " & .Tipo, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Servidor: " & .Servidor, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Puerto: " & .Puerto, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Base: " & .Base, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Usuario: " & .Usuario, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Contraseña: " & VLstrTemp, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Seguridad integrada: " & .SI, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Instancia: " & .Instancia, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                    End With
                    VLkeyRegistro.Close()
                    Notifica("<-- Finaliza lectura configuración base central", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
                    Return True
                Catch ex As System.Security.SecurityException
                    VMstrError = "El usuario no tiene permisos suficientes para leer la clave del registro [BD Central]"
                Catch ex As ObjectDisposedException
                    VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [BD Central]"
                Catch ex As IOException
                    VMstrError = "El objeto Registry que contiene el valor especificado se ha marcado para su eliminación [BD Central]"
                Catch ex As UnauthorizedAccessException
                    VMstrError = "El usuario no tiene los derechos necesarios en el Registro [BD Central]"
                Catch ex As Exception
                    VMstrError = ex.Message & " [BD Central]"
                End Try
            Else
                Notifica("  No se ha encontrado la configuración de la base de datos central", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError, vbNewLine & "<-- Finaliza lectura configuración base central")
            End If
        Catch ex As ArgumentNullException
            VMstrError = "No se ha especificado la llave a consultar [BD Central]"
        Catch ex As ArgumentException
            VMstrError = "El nombre de la llave excede el número de caracteres permitidos [BD Central]"
        Catch ex As ObjectDisposedException
            VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [BD Central]"
        Catch ex As Exception
            VMstrError = ex.Message & " [BD Central]"
        End Try
    End Function
    Private Function cargaBDBitacora() As Boolean
        Dim VLkeyRegistro As RegistryKey
        Dim VLclsEncripta As New RC4
        Dim VLstrTemp As String
        Dim VLstrRuta

        VLstrRuta = My.Resources.Registro & VMConfiguracion & "\\BD Bitacora"
        Notifica("--> Obteniendo información de la base de datos central", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno, " " & VLstrRuta)

        Try
            VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VLstrRuta, False)

            If Not IsNothing(VLkeyRegistro) Then
                Try
                    With VGtipBDBitacora
                        .Base = VLkeyRegistro.GetValue("Base")
                        VLstrTemp = VLkeyRegistro.GetValue("Contrasena")
                        .Instancia = VLkeyRegistro.GetValue("Esquema")
                        .Puerto = VLkeyRegistro.GetValue("Puerto")
                        .Servidor = VLkeyRegistro.GetValue("Servidor")
                        .SI = Convert.ToBoolean(VLkeyRegistro.GetValue("SI"))
                        .Tipo = VLkeyRegistro.GetValue("Tipo")
                        .Usuario = VLkeyRegistro.GetValue("Usuario")
                        If VLstrTemp.Length > 0 Then
                            Try
                                .Contrasena = desencriptaValor(VLstrTemp)
                            Catch ex As Exception
                                VMstrError = "No se pudo obtener la contraseña de la base de datos central " & ex.Message
                                Return False
                            End Try
                        Else
                            .Contrasena = ""
                        End If

                        Notifica("  Tipo: " & .Tipo, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Servidor: " & .Servidor, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Puerto: " & .Puerto, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Base: " & .Base, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Usuario: " & .Usuario, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Contraseña: " & VLstrTemp, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Seguridad integrada: " & .SI, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                        Notifica("  Instancia: " & .Instancia, Me, enuEscribeLog.Completo, enuTipoAviso.Ninguno)
                    End With
                    VLkeyRegistro.Close()
                    Notifica("<-- Finaliza lectura configuración base central", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
                    Return True
                Catch ex As System.Security.SecurityException
                    VMstrError = "El usuario no tiene permisos suficientes para leer la clave del registro [BD Bitácora]"
                Catch ex As ObjectDisposedException
                    VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [BD Bitácora]"
                Catch ex As IOException
                    VMstrError = "El objeto Registry que contiene el valor especificado se ha marcado para su eliminación [BD Bitácora]"
                Catch ex As UnauthorizedAccessException
                    VMstrError = "El usuario no tiene los derechos necesarios en el Registro [BD Bitácora]"
                Catch ex As Exception
                    VMstrError = ex.Message & " [BD Bitácora]"
                End Try
            Else
                Notifica("  No se ha encontrado la configuración de la base de datos central", Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError, vbNewLine & "<-- Finaliza lectura configuración base de bitácora")
            End If
        Catch ex As ArgumentNullException
            VMstrError = "No se ha especificado la llave a consultar [BD Bitácora]"
        Catch ex As ArgumentException
            VMstrError = "El nombre de la llave excede el número de caracteres permitidos [BD Bitácora]"
        Catch ex As ObjectDisposedException
            VMstrError = "El objeto no ha sido abierto o se ha cerrado anticipadamente [BD Bitácora]"
        Catch ex As Exception
            VMstrError = ex.Message & " [BD Bitácora]"
        End Try
    End Function
    Private Function abrirBase(ByVal Tipo As TipoBD) As Boolean
        '***********************************************************************************************'
        '       Prueba las conexiones con las bases de datos central, transaccional y bitacora          '
        '                                                                                               '
        '       Devuelve: TRUE si las conexiones a las bases de datos estan correctos                   '
        '                 FALSE si alguna conexion falla al conectarse                                  '
        '***********************************************************************************************'
        Dim VLdbsBase As Object
        Dim VLtipBase As tipBD
        Dim VLblnRes As Boolean

        If Tipo = TipoBD.Bitacora Then
            VLtipBase = VGtipBDBitacora
        ElseIf Tipo = TipoBD.Central Then
            VLtipBase = VGtipBDCentral
        Else
            VMstrError = "  Tipo de base de datos desconocido " & Tipo
            Return False
        End If
        Notifica(" --> Abirendo base de Datos " & Tipo, Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)

        Select Case VLtipBase.Tipo
            Case TipoBase.MySQL
                VLdbsBase = New MySqlConnection
            Case TipoBase.SQLServer
                VLdbsBase = New SqlConnection
            Case TipoBase.Oracle
                VLdbsBase = New OracleConnection
            Case Else
                VLdbsBase = Nothing
                VMstrError = "Manejador de base de datos desconocido"
                Return False
        End Select
        Try
            VLdbsBase.ConnectionString = cadenaConexion(VLtipBase)
            VLdbsBase.Open()
            VLdbsBase.Close()
            VLblnRes = True
            Notifica(" Conexión Exitosa ", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
        Catch ex As Exception
            VMstrError = " Error al abrir la base: " & ex.Message
            VLblnRes = False
        End Try
        Notifica(" <-- Finaliza Apertura la apertura de la base ", Me, enuEscribeLog.Ambos, enuTipoAviso.Ninguno)
        VLdbsBase.Dispose()
        VLdbsBase = Nothing
        Return VLblnRes
    End Function
    Private Sub Reintentar()
        If Not VMtmrInicia.Enabled Then
            VMtmrInicia.Interval = 60000
            VMtmrInicia.AutoReset = False
            VMtmrInicia.Start()
        End If
    End Sub
    Private Function cargarPuertos() As Boolean

        Dim VLdbsBase As Object
        Dim VLcmdComando As Object
        Dim VLrcsDatos As Object
        Dim VLstrSQL As String
        Dim VLblnRespuesta As Boolean

        Select Case VGtipBDCentral.Tipo
            Case TipoBase.MySQL
                VLdbsBase = New MySqlConnection
                VLcmdComando = New MySqlCommand
            Case TipoBase.SQLServer
                VLdbsBase = New SqlConnection
                VLcmdComando = New SqlCommand
            Case TipoBase.Oracle
                VLdbsBase = New OracleConnection
                VLcmdComando = New OracleCommand
            Case Else
                Return False
        End Select
        VLdbsBase.connectionstring = cadenaConexion(VGtipBDCentral)

        VLstrSQL = "SELECT * FROM TCONALARFILTROS"
        Try
            VLdbsBase.open()
            VLcmdComando.commandtext = VLstrSQL
            VLcmdComando.connection = VLdbsBase

            Try
                VLrcsDatos = VLcmdComando.executereader
                If VLrcsDatos.hasrows Then
                    While VLrcsDatos.read
                        VMhasPuertos.Add(VLrcsDatos!Puerto, VLrcsDatos!Num_Filtro)
                    End While
                    VLblnRespuesta = True
                Else
                    VMstrError = "No se han encontrado puertos configurados para el envío de transacciones"
                    ''escribeVisorEventos("No se han encontrado puertos configurados para el envío de transacciones", EventLogEntryType.Error)
                End If
                VLrcsDatos.close()
            Catch ex As Exception

            End Try
            VLcmdComando.dispose()
            VLdbsBase.close()
        Catch ex As Exception
            escribeVisorEventos("No se pudo abrir la base de datos central por el siguiente motivo" & vbNewLine & ex.Message, EventLogEntryType.Error)
        End Try
        VLrcsDatos = Nothing
        VLcmdComando = Nothing
        VLdbsBase = Nothing

        Return VLblnRespuesta
    End Function
#End Region
#Region "Administrador"
    Private Function ConfigurarAdmin() As Boolean
        '****************************************************************************************
        '              Permite la configuracion y la conexion con el administrador
        '****************************************************************************************
        Dim VLkeyRegistro As RegistryKey

        Try
            VLkeyRegistro = Registry.LocalMachine.OpenSubKey("SOFTWARE\\KS Soluciones\\Key Monitor\\V3\\" & VMConfiguracion)
            If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                escribeLog(vbNewLine & "-->Inicia Lectura Configuracion de los parametros del Administrador" & vbNewLine & "  [" & VLkeyRegistro.ToString & "]")
            End If
            If Not (VLkeyRegistro Is Nothing) Then
                VMtcpAdmin = New Cliente_TCP
                VMtcpAdmin.DireccionIP = "localhost"
                VMtcpAdmin.Puerto = VLkeyRegistro.GetValue("Puerto")
                VMtcpAdmin.ActivarEco = True
                VMtcpAdmin.Conectar()
                VMblnSigue = False
                Do
                    Thread.Sleep(1000)
                Loop While Not VMblnSigue
                ConfigurarAdmin = True
                If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                    escribeLog("--> Conexion Satisfactoria con el Administrador")
                    escribeLog("    Servidor= " & VMtcpAdmin.DireccionIP)
                    escribeLog("    Puerto= " & VMtcpAdmin.Puerto)
                ElseIf _Log = enuDepuracion.Consola Then
                    escribeConsola("--> Conexion Satisfactoria con el Administrador")
                    escribeConsola("    Servidor= " & VMtcpAdmin.DireccionIP)
                    escribeConsola("    Puerto= " & VMtcpAdmin.Puerto)
                End If
            Else
                ConfigurarAdmin = False
                If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                    escribeLog("No se encuentra el puerto de conexion para el Administrador")
                End If
                escribeVisorEventos("No se encuentra el puerto de conexion para el Administrador", EventLogEntryType.Warning)
            End If
        Catch ex As Exception
            ConfigurarAdmin = False
            If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                escribeLog("Error de configuracion del Administrador: " & ex.Message)
            End If
            escribeVisorEventos("Error de configuracion del Administrador: " & ex.Message, EventLogEntryType.Warning)
        End Try
        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
            escribeLog("<--Fin Lectura Configuracion de los parametros del Administrador")
        ElseIf _Log = enuDepuracion.Consola Then
            escribeConsola("<--Fin Lectura Configuracion del Administrador")
        End If
    End Function
    Private Sub VMtcpAdmin_ConexionEstablecida() Handles VMtcpAdmin.ConexionEstablecida
        Try
            VMblnSigue = True
            If VGtipGeneral.Tipo = TipoGuia.Linea Then
                VMtcpAdmin.EnviarDatos(STX & "04" & ETX)
            ElseIf VGtipGeneral.Tipo = TipoGuia.Filtros Then
                VMtcpAdmin.EnviarDatos(STX & "02" & ETX)
            End If
        Catch ex As Exception
            escribeVisorEventos("Error al identificarse con el administrador: " & ex.Message, EventLogEntryType.Warning)
        End Try
    End Sub
    Private Sub VMtcpAdmin_ConexionTerminada(ByVal DesconexionEco As Boolean) Handles VMtcpAdmin.ConexionTerminada
        If DesconexionEco = True Then
            ConfigurarAdmin()
        End If
    End Sub
    Private Sub VMtcpAdmin_DatosRecibidos(ByVal Mensaje As String) Handles VMtcpAdmin.DatosRecibidos
        Dim VLintInicio As Integer
        Dim VLintLen As Integer
        Dim VLstrMensaje As String
        Dim VLstrDatos() As String
        Dim VLstrRespuesta As String = ""
        Dim VLstrEstado As String = "00"
        Dim VLobjGuia As Guia
        Dim VLobjFiltros As GuiaFiltros
        Try
            If Mensaje.Contains(STX) And Mensaje.Contains(ETX) Then
                VLintInicio = InStr(Mensaje, STX)
                VLintLen = InStr(Mensaje, ETX) - VLintInicio - 1
                VLstrMensaje = Mensaje.Substring(VLintInicio, VLintLen)
                VLstrDatos = Split(VLstrMensaje, FS)
                Select Case VLstrDatos(0)
                    Case "7001"
                        Try
                            If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                VLstrEstado = "02"
                            End If
                            '<STX>7001<FS>ACK<FS>No. De Identificador de Estatus<FS>Bloques Recibidos<FS>Bloques Procesados<FS>Tamaño  Memoria<FS>No. Clientes Conectados <FS>Puertos Disponibles <FS> Ultima Transacción Recibida (Hora Y Fecha) <FS> Bloques(Ignorados) <ETX>
                            VLstrRespuesta = STX & "7001" & FS & ACK & FS & VLstrEstado & FS & VMintBloques & FS & VMintProcesados & FS & VMobjBloques.Count & FS & Guia.UsuariosConectados & FS & VMobjClientes.Count - Guia.UsuariosConectados & FS & VMstrUltimoBloque & ETX
                        Catch ex As Exception
                            '<STX>7001 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7001" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "7002"
                        Try
                            If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                VLstrEstado = "02"
                            End If
                            '<STX>7002<FS>ACK<FS>No. De Identificador de Estatus<FS>Bloques Recibidos<FS>Bloques Procesados<FS>Tamaño en Memoria <FS>No. Clientes Conectados<FS>No. Puertos Disponibles <FS> Ultima Transacción Recibida (Hora Y Fecha) <FS> Bloques(Ignorados) <ETX>
                            VLstrRespuesta = STX & "7002" & FS & ACK & FS & VLstrEstado & FS & VMintBloques & FS & VMintProcesados & FS & VMobjBloques.Count & FS & GuiaFiltros.UsuariosConectados & FS & VMobjClientes.Count - GuiaFiltros.UsuariosConectados & FS & VMstrUltimoBloque & FS & VMintIgnorados & ETX
                        Catch ex As Exception
                            '<STX>7002 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7002" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "7004"
                        Try
                            If VMtcpCliente Is Nothing Then
                                VLstrEstado = "02"
                                '<STX>7002 <FS> NAK <FS> MENSAJE ERROR <ETX>
                                VLstrRespuesta = STX & "7004" & FS & NAK & FS & "No se ha conectado con el Servidor Tranasaccional" & ETX
                            Else
                                If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                    VLstrEstado = "02"
                                End If
                            End If
                            '<STX>7002<FS>ACK<FS>No. De Identificador de Estatus<FS>Bloques Recibidos<FS>Bloques Procesados<FS>Tamaño en Memoria <FS>No. Clientes Conectados<FS>No. Puertos Disponibles <FS> Ultima Transacción Recibida (Hora Y Fecha) <FS> Bloques(Ignorados) <ETX>
                            VLstrRespuesta = STX & "7004" & FS & ACK & FS & VLstrEstado & FS & VMintBloques & FS & VMintProcesados & FS & VMobjBloques.Count & FS & Guia.UsuariosConectados & FS & VMobjClientes.Count - Guia.UsuariosConectados & FS & VMstrUltimoBloque & FS & VMintIgnorados & ETX
                        Catch ex As Exception
                            '<STX>7002 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7004" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "7011"
                        Try
                            If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                VLstrEstado = "02"
                                VLstrRespuesta = "No hay conexion con el servidor transaccional"
                            End If
                            '<STX>7011<FS> ACK <FS>No. Identificador Estatus<FS>Motivo del Error o Alerta <ETX>
                            VLstrRespuesta = STX & "7011" & FS & ACK & FS & VLstrEstado & FS & VLstrRespuesta & ETX
                        Catch ex As Exception
                            '<STX>7011 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7011" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "7012"
                        Try
                            If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                VLstrEstado = "02"
                                VLstrRespuesta = "No hay conexion con el servidor transaccional"
                            End If
                            '<STX>7012<FS>ACK <FS>No. Identificador  de Estatus<FS>Motivo del Error o Alerta <ETX> 
                            VLstrRespuesta = STX & "7012" & FS & ACK & FS & VLstrEstado & FS & VLstrRespuesta & ETX
                        Catch ex As Exception
                            '<STX>7012 <FS> NAK <FS>MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7012" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "7014"
                        Try
                            If VMtcpCliente.Estado <> Cliente_TCP.EstadoConexion.Conectado Then
                                VLstrEstado = "02"
                                VLstrRespuesta = "No hay conexion con el servidor transaccional"
                            End If
                            '<STX>7012<FS>ACK <FS>No. Identificador  de Estatus<FS>Motivo del Error o Alerta <ETX> 
                            VLstrRespuesta = STX & "7014" & FS & ACK & FS & VLstrEstado & FS & VLstrRespuesta & ETX
                        Catch ex As Exception
                            '<STX>7012 <FS> NAK <FS>MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "7014" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "8011" 'Guia Replicación
                        'Solicitud de restaurar un cliente
                        Try
                            VLobjGuia = VMobjClientes(CInt(VLstrDatos(1)))
                            'Verifica que el cliente solicitado no esté en proceso de restauración
                            If Not VLobjGuia.Restaurando Then
                                VMtcpAdmin.EnviarDatos(ACK)
                                'Le indica al cliente que reinicie
                                'VLobjGuia.Reiniciar()
                                VLstrRespuesta = ""
                            Else
                                'Le indica que ya se está en proceso de restauración
                                VLstrRespuesta = STX & "8001" & FS & NAK & FS & VLstrDatos(1) & FS & VLobjGuia.Usuario & FS & VLobjGuia.NombrePC & FS & VLobjGuia.IP & FS & VLobjGuia.Puerto & FS & VLobjGuia.Aplicacion & FS & VLobjGuia.Recibidos & FS & VLobjGuia.Procesados & FS & VLobjGuia.Memoria & ETX
                            End If
                            '<STX>8001<FS>ACK <FS>No. Cliente<FS>Usuario <FS>Nombre PC<FS>Dirección IP<FS>Puerto  <FS>Aplicación <FS>Bloques Recibidos <FS>Bloques Procesados<FS>Tamaño Memoria<ETX>
                            VLstrRespuesta = STX & "8001" & FS & ACK & FS & VLstrDatos(1) & FS & VLobjGuia.Usuario & FS & VLobjGuia.NombrePC & FS & VLobjGuia.IP & FS & VLobjGuia.Puerto & FS & VLobjGuia.Aplicacion & FS & VLobjGuia.Recibidos & FS & VLobjGuia.Procesados & FS & VLobjGuia.Memoria & ETX
                        Catch ex As Exception
                            '<STX> 8001 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "8001" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "8012" 'Guia Filtros
                        'Solicitud de restaurar un cliente
                        'Verifica que el cliente solicitado no esté en proceso de restauración
                        Try
                            Select Case VLstrDatos(1) 'Indica el tipo de petición realizada
                                Case "RC" 'Restablecer Conexión
                                    For Each VLobjFiltros In VMobjClientes.Values
                                        If VLobjFiltros.Puerto = CInt(VLstrDatos(2)) Then
                                            If Not VLobjFiltros.Reconectando Then
                                                VLobjFiltros.RestaurarConexion()
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    'VLobjFiltros = VMobjClientes(CInt(VLstrDatos(2)))
                                Case "VE" 'Validar Envió
                                Case "DE" 'Detener Envió
                                Case "DC" 'Diagnosticar Conflicto
                                Case Else
                            End Select
                        Catch ex As Exception
                            VLstrRespuesta = STX & "8012" & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & VLstrDatos(2) & "( " & ex.Message & ") " & ETX
                        End Try
                    Case "8014" 'Guia En Linea
                        Try
                            Select Case VLstrDatos(1) 'Indica el tipo de petición realizada
                                Case "RC" 'Restablecer Conexión
                                    For Each VLobjGuia In VMobjClientes.Values
                                        If VLobjGuia.Puerto = CInt(VLstrDatos(2)) Then
                                            If Not VLobjGuia.Reconectando Then
                                                VLobjGuia.RestaurarConexion()
                                            End If
                                            Exit For
                                        End If
                                    Next
                                    'VLobjFiltros = VMobjClientes(CInt(VLstrDatos(2)))
                                Case "VE" 'Validar Envió
                                Case "DE" 'Detener Envió
                                Case "DC" 'Diagnosticar Conflicto
                                Case Else
                            End Select
                        Catch ex As Exception
                            VLstrRespuesta = STX & "8014" & FS & NAK & FS & "No se pudo restablecer el servidor en el puerto " & VLstrDatos(2) & "( " & ex.Message & ") " & ETX
                        End Try

                    Case "7024" 'Guia Linea
                        'Solicitud de restaurar un cliente
                        'Verifica que el cliente solicitado no esté en proceso de restauración

                    Case "7031"
                        'Solicitud de duplicar la información
                    Case "7032"
                        'Solicitud de duplicar la información
                    Case "7034"
                        'Solicitud de duplicar la información
                    Case "8001"
                        Try
                            VLobjGuia = VMobjClientes(CInt(VLstrDatos(1)))
                            '<STX>8001<FS>ACK <FS>No. Cliente<FS>Usuario <FS>Nombre PC<FS>Dirección IP<FS>Puerto  <FS>Aplicación <FS>Bloques Recibidos <FS>Bloques Procesados<FS>Tamaño Memoria<ETX>
                            VLstrRespuesta = STX & "8001" & FS & ACK & FS & VLstrDatos(1) & FS & VLobjGuia.Usuario & FS & VLobjGuia.NombrePC & FS & VLobjGuia.IP & FS & VLobjGuia.Puerto & FS & VLobjGuia.Aplicacion & FS & VLobjGuia.Recibidos & FS & VLobjGuia.Procesados & FS & VLobjGuia.Memoria & ETX
                        Catch ex As Exception
                            '<STX> 8001 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "8001" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "8002"
                        Try
                            VLobjFiltros = VMobjClientes(CInt(VLstrDatos(1)))
                            '<STX>8002<FS> ACK <FS>
                            VLstrRespuesta = STX & "8002" & FS & ACK & FS
                            'No. Cliente <FS> 
                            VLstrRespuesta &= VLstrDatos(1) & FS
                            'Usuario <FS>
                            VLstrRespuesta &= VLobjFiltros.Usuario & FS
                            'Nombre PC <FS>
                            VLstrRespuesta &= VLobjFiltros.NombrePC & FS
                            'Dirección IP <FS>
                            VLstrRespuesta &= VLobjFiltros.IP & FS
                            'Puerto  <FS>
                            VLstrRespuesta &= VLobjFiltros.Puerto & FS
                            'Aplicación <FS>
                            VLstrRespuesta &= VLobjFiltros.Aplicacion & FS
                            'Bloques Recibidos <FS> 
                            VLstrRespuesta &= VLobjFiltros.Recibidos & FS
                            'Bloques Procesados <FS>
                            VLstrRespuesta &= VLobjFiltros.Procesados & FS
                            'Tamaño de Memoria <FS>
                            VLstrRespuesta &= VLobjFiltros.Memoria & FS
                            'Transacciones Totales <FS>
                            VLstrRespuesta &= VLobjFiltros.TransaccionesTotales & FS
                            'Transacciones Filtradas <FS>
                            VLstrRespuesta &= VLobjFiltros.TransaccionesFiltradas & FS
                            'Filtro <FS>
                            VLstrRespuesta &= VLobjFiltros.NumFiltro & FS
                            'Fecha Conexión <FS>
                            VLstrRespuesta &= VLobjFiltros.FechaConexion & FS
                            'Estado del Puerto <FS>
                            VLstrRespuesta &= VLobjFiltros.EstadoPuerto & FS
                            'Bloques Ignorados <FS>
                            VLstrRespuesta &= VLobjFiltros.Ignoradas & FS
                            'Ultima Transacción Recibida(Hora y Fecha)<FS>  
                            VLstrRespuesta &= VLobjFiltros.UltimaTransaccion & FS
                            'Transacciones Enviadas
                            VLstrRespuesta &= VLobjFiltros.TransaccionesEnviadas & ETX

                        Catch ex As Exception
                            '<STX> 8002 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "8002" & FS & NAK & FS & ex.Message & ETX
                        End Try
                    Case "8004"
                        Try
                            VLobjGuia = VMobjClientes(CInt(VLstrDatos(1)))
                            With VLobjGuia
                                '<STX>8002<FS> ACK <FS>
                                VLstrRespuesta = STX & "8004" & FS & ACK & FS
                                'No. Cliente <FS> 
                                VLstrRespuesta &= VLstrDatos(1) & FS
                                'Usuario <FS>
                                VLstrRespuesta &= .Usuario & FS
                                'Nombre PC <FS>
                                VLstrRespuesta &= .NombrePC & FS
                                'Dirección IP <FS>
                                VLstrRespuesta &= .IP & FS
                                'Puerto  <FS>
                                VLstrRespuesta &= .Puerto & FS
                                'Aplicación <FS>
                                VLstrRespuesta &= .Aplicacion & FS
                                'Bloques Recibidos <FS> 
                                VLstrRespuesta &= .Recibidos & FS
                                'Bloques Procesados <FS>
                                VLstrRespuesta &= .Procesados & FS
                                'Tamaño de Memoria <FS>
                                VLstrRespuesta &= .Memoria & FS
                                'Fecha Conexión <FS>
                                VLstrRespuesta &= .FechaConexion & FS
                                'Estado del Puerto <FS>
                                VLstrRespuesta &= .EstadoPuerto & FS
                                'Bloques Ignorados <FS>
                                VLstrRespuesta &= .Ignoradas & FS
                                'Ultima Transacción Recibida(Hora y Fecha)<FS>  
                                VLstrRespuesta &= .UltimaTransaccion & FS & ETX
                            End With

                        Catch ex As Exception
                            '<STX> 8002 <FS> NAK <FS> MENSAJE ERROR <ETX>
                            VLstrRespuesta = STX & "8004" & FS & NAK & FS & ex.Message & " mensaje=" & VLstrRespuesta & ETX
                        End Try
                    Case "5004"
                        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                            VGioArchivoLog.Close()
                        End If
                        Select Case VLstrDatos(1).Trim.ToLower
                            Case "/ds"
                                _Log = enuDepuracion.Sencilla
                                abreArchivoLog()
                                VLstrRespuesta = STX & "5004" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                            Case "/dt"
                                _Log = enuDepuracion.Completa
                                abreArchivoLog()
                                VLstrRespuesta = STX & "5004" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                            Case "/dc"
                                _Log = enuDepuracion.Consola
                                VLstrRespuesta = STX & "5004" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                            Case "/dtran"
                                If Not VPblnLogtran Then
                                    'Prende el log
                                    Try
                                        VPioLogTran = Nothing
                                        VPioLogTran = New StreamWriter(My.Application.Info.DirectoryPath & "\RastroKM" & Format(Now, "ddMMyyyyHHmmss") & ".dat", True)
                                        VPioLogTran.AutoFlush = True
                                        VLstrRespuesta = STX & "5004" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                                        VPblnLogtran = True
                                    Catch ex As System.IO.IOException
                                        VLstrRespuesta = STX & "5004" & FS & NAK & FS & "No se pudo crear el archivo " & ex.Message & ETX
                                    Catch ex As Exception
                                        VLstrRespuesta = STX & "5004" & FS & NAK & FS & "No se pudo crear el archivo " & ex.Message & ETX
                                    End Try
                                Else
                                    'Apaga el log
                                    Try
                                        VPioLogTran.Close()
                                        VPblnLogtran = False
                                    Catch ex As IOException
                                        VLstrRespuesta = STX & "5004" & FS & NAK & FS & "No se pudo cerrar el archivo " & ex.Message & ETX
                                    Catch ex As Exception
                                        VLstrRespuesta = STX & "5004" & FS & NAK & FS & "No se pudo cerrar el archivo " & ex.Message & ETX
                                    End Try
                                End If
                            Case Else
                                _Log = enuDepuracion.Ninguna
                                VLstrRespuesta = STX & "5004" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                        End Select
                    Case "5002"
                        If _Log = enuDepuracion.Completa Or _Log = enuDepuracion.Sencilla Then
                            VGioArchivoLog.Close()
                        End If
                        Select Case VLstrDatos(1).Trim.ToLower
                            Case "/ds"
                                _Log = enuDepuracion.Sencilla
                                abreArchivoLog()
                            Case "/dt"
                                _Log = enuDepuracion.Completa
                                abreArchivoLog()
                            Case "/dc"
                                _Log = enuDepuracion.Consola
                            Case Else
                                _Log = enuDepuracion.Ninguna
                        End Select
                        VLstrRespuesta = STX & "5002" & FS & ACK & FS & VLstrDatos(1).Trim & ETX
                    Case Else
                        VLstrRespuesta = ""
                End Select
                If VLstrRespuesta.Trim.Length > 0 Then
                    VMtcpAdmin.EnviarDatos(VLstrRespuesta)
                End If
            ElseIf Not Mensaje = ACK Then
                VMtcpAdmin.EnviarDatos(NAK)
            End If
        Catch ex As Exception
            VMtcpAdmin.EnviarDatos(ex.Message)
        End Try
    End Sub
    Private Sub VMtcpAdmin_ErrorConexion(ByVal Mensaje As String) Handles VMtcpAdmin.ErrorConexion
        escribeVisorEventos("Error al conectarse con el Administrador: " & Mensaje, EventLogEntryType.Warning)
        VMblnSigue = True
    End Sub
    Public Sub escribeConsola(ByVal VPMensaje As String)
        Dim VLstrTipo As String = ""
        If VGtipGeneral.Tipo = TipoGuia.Linea Then
            VLstrTipo = "4014"
        ElseIf VGtipGeneral.Tipo = TipoGuia.Filtros Then
            VLstrTipo = "4012"
        End If

        Try
            If VMtcpAdmin.Estado = Cliente_TCP.EstadoConexion.Conectado Then
                VMtcpAdmin.EnviarDatos(STX & VLstrTipo & FS & VPMensaje & ETX)
            Else
                escribeVisorEventos("Error al conectarse con el Administrador. Se desactivara el modo de Depuracion consola", EventLogEntryType.Error)
                ConfigurarAdmin()
            End If
        Catch ex As Exception
            escribeVisorEventos("Error al conectarse con el Administrador. Se desactivara el modo de Depuracion consola" & vbNewLine & ex.Message, EventLogEntryType.Error)
            _Log = enuDepuracion.Ninguna
        End Try
    End Sub
#End Region
    Public Sub New()
        VPioLogTran = Nothing
        VPblnLogtran = False
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
    End Sub
    Private Sub VMtmrReconecta_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrReconecta.Elapsed
        Dim VLstrRespuesta As String

        VMtmrReconecta.Enabled = False

        VMbytIntentaReconexion += 1
        VMblnReconectando = True

        VLstrRespuesta = STX & 7014 & FS & ACK & FS & "01" & FS & "Reconectando con servidor transaccional [Intento " & VMbytIntentaReconexion & "]..." & ETX
        VMtcpAdmin.EnviarDatos(VLstrRespuesta)

        SyncLock VMobjBloques
            VMobjBloques.Clear()
        End SyncLock

        Try
            VMtcpCliente.Cerrar()
        Catch ex As Exception

        End Try

        VMtcpCliente = New Cliente_TCP
        VMtcpCliente.Puerto = VGtipGeneral.Puerto
        VMtcpCliente.DireccionIP = VGtipGeneral.IP
        VMtcpCliente.Conectar()
        VMblnSigue = False
        Do
            Thread.Sleep(1000)
        Loop While Not VMblnSigue

        If VMtcpCliente.Estado = Cliente_TCP.EstadoConexion.Conectado Then
            VMblnListo1 = False
            Inicia()

            Notifica("Se restableció correctamente la comunicación con el servidor trasacccional [Intento: " & VMbytIntentaReconexion & "] ", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True)
            VLstrRespuesta = STX & 7014 & FS & ACK & FS & "00" & FS & "Se restableció correctamente la comunicación con el servidor trasacccional [Intento: " & VMbytIntentaReconexion & "] " & ETX
            VMtcpAdmin.EnviarDatos(VLstrRespuesta)

            VLstrRespuesta = STX & "7004" & FS & ACK & FS & "00" & FS & VMintBloques & FS & VMintProcesados & FS & VMobjBloques.Count & FS & Guia.UsuariosConectados & FS & VMobjClientes.Count - Guia.UsuariosConectados & FS & VMstrUltimoBloque & FS & VMintIgnorados & ETX

            VMtcpAdmin.EnviarDatos(VLstrRespuesta)
            VMbytIntentaReconexion = 0
            VMintSinTransaccion = 0
            VMtmrTiempoSinTran = New Timers.Timer
            VMtmrTiempoSinTran.Interval = 60000
            VMtmrTiempoSinTran.Enabled = True
        Else
            Notifica("No se pudo restablecer la comunicación con el servidor transaccional [Intento: " & VMbytIntentaReconexion & "] " & vbNewLine & "Se reintentará en " & IIf(VMbytIntentaReconexion < 5, "1", "5") & " minuto(s)", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True)
            If VMbytIntentaReconexion < 5 Then
                VMtmrReconecta.Interval = 60000
            Else
                VMtmrReconecta.Interval = 300000
            End If
            VMtmrReconecta.Enabled = True
        End If
        VMblnReconectando = False
    End Sub
    Private Function ReinciaSvrTxn(ByVal Horario As String) As Boolean
        Notifica("No se han recibido transacciones en los últimos " & VMintSinTransaccion & " minuto(s) [Horario: " & Horario & "] " & vbNewLine & "Reiniciando conexión con servidor transaccional", Me, enuEscribeLog.Ambos, enuTipoAviso.Advertencia, , True )
        VMtmrReconecta.Interval = 100
        VMtmrReconecta.Enabled = True
    End Function
    Public Sub validaEnvio()
        'Verifica que en la cola de salida no se estén acumulando los mensajes
        If Not VMblnValidaEnvio Then
            VMblnValidaEnvio = True
            Notifica("Validando estado de envío [" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "]", Me, enuEscribeLog.Ambos, enuTipoAviso.Informativo)
            'Validar el estado de los clientes
            If Not _Procesando And VMobjBloques.Count > 0 Then
                VMtmrProcesa.Enabled = True
            End If
            VMblnValidaEnvio = False
        End If
    End Sub
    Private Sub VMtmrProcesa_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrProcesa.Elapsed
        Dim VLstrMensaje As String
        VMtmrProcesa.Enabled = False
        _Procesando = True
        Try
            While VMobjBloques.Count > 0
                SyncLock VMobjBloques
                    VLstrMensaje = VMobjBloques.Dequeue
                End SyncLock
                fncCargarBloque(VLstrMensaje)
            End While

        Catch ex As Exception
            Notifica("  Error al procesar los bloques " & vbNewLine & ex.Message, Me, enuEscribeLog.Ambos, enuTipoAviso.MensajeError)
            'Informar a la consola
        End Try
        _Procesando = False
    End Sub

End Class