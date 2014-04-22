''' <summary>
''' Permite agregar los mensajes dependiendo el tipo de depuracion
''' </summary>
''' <remarks></remarks>
Public Class Depuracion

#Region "Estructuras"
    Public Enum TipoRegistro
        Visor = 0
        Correo = 1
        Archivo = 2
        Administrador = 3
    End Enum

    Public Enum TipoDepuracion
        Normal = 0
        Simple = 1
        Completa = 2
        Consola = 3
    End Enum

    Private Structure Mensajes
        Dim Mensaje As String
        Dim Tipo As System.Diagnostics.EventLogEntryType
    End Structure

    Public Enum ProductoKS
        KeyMonitor = 1
        Perfiles = 2
    End Enum
#End Region

#Region "Variables"
    Private _Producto As ProductoKS
    Private _NombreArchivo As String
    Private _RutaDirectorio As String
    Private _NombreAplicacion As String
    Private _TipoDepuracion As TipoDepuracion
    Private _Registro As TipoRegistro
    Private _Teminar As Boolean
    Private _IdConsola As String

    Private VMcolaMensajes As Queue
    Private VMcolaAdmin As Queue
    Private VMhiloAdmin As Threading.Thread
    Private VMhiloProcesar As Threading.Thread
    Private VMtcpAdministrador As Cliente_TCP
    Private Shared VMobjDepuracion As Depuracion


    Private WithEvents VMtmrLog As Timers.Timer      ''Timer para depurar carpeta \log
    Private VMintTiempo As Integer                   ''Tiempo en milisegundos para depurar carpeta \log
#End Region

#Region "Propiedades Publicas"
   
    ''' <summary>
    ''' Nombre con el cual se va a registrar en el visor de eventos
    ''' </summary>
    Public Property NombreAplicacion() As String
        Get
            Return _NombreAplicacion
        End Get
        Set(ByVal value As String)
            _NombreAplicacion = value
        End Set
    End Property
    ''' <summary>
    ''' Indica el tipo de depuracion que se debe realizar
    ''' </summary>
    Public Property Depuracion() As TipoDepuracion
        Get
            Return _TipoDepuracion
        End Get
        Set(ByVal value As TipoDepuracion)
            _TipoDepuracion = value
            Select Case value
                Case TipoDepuracion.Completa
                    _Registro = TipoRegistro.Archivo
                Case TipoDepuracion.Consola
                    _Registro = TipoRegistro.Administrador
                Case TipoDepuracion.Normal
                    _Registro = TipoRegistro.Visor
                Case TipoDepuracion.Simple
                    _Registro = TipoRegistro.Archivo
            End Select
        End Set
    End Property
    ''' <summary>
    ''' Conexion TCP que se tiene con el adminsitrador
    ''' </summary>
    Public Property Adminstrador() As Cliente_TCP
        Get
            Return VMtcpAdministrador
        End Get
        Set(ByVal value As Cliente_TCP)
            VMtcpAdministrador = value
        End Set
    End Property
    ''' <summary>
    ''' El numero de mensajes que faltan de procesar
    ''' </summary>
    Public ReadOnly Property Alertas() As Integer
        Get
            Return VMcolaMensajes.Count
        End Get
    End Property
    ''' <summary>
    ''' Producto que esta generando el log
    ''' </summary>
    Public Property Producto() As ProductoKS
        Get
            Return _Producto
        End Get
        Set(ByVal value As ProductoKS)
            _Producto = value
        End Set
    End Property
    ''' <summary>
    ''' Identificador que se le tiene que enviar a la consola
    ''' </summary>
    Public Property IdConsola() As String
        Get
            Return _IdConsola
        End Get
        Set(ByVal value As String)
            _IdConsola = value
        End Set
    End Property
    ''' <summary>
    ''' Devuelve la instancia unica de trabajo
    ''' </summary>
    Public Shared ReadOnly Property Instancia() As Depuracion
        Get
            If VMobjDepuracion Is Nothing Then
                VMobjDepuracion = New Depuracion
            End If
            Return VMobjDepuracion
        End Get
    End Property
#End Region

#Region "Funciones Privadas"
    ''' <summary>
    ''' Permite guardar el mensaje en el visor de eventos
    ''' </summary>
    Private Sub VisorEventos(ByVal VPstrMensaje As String, Optional ByVal VPobjTypo As System.Diagnostics.EventLogEntryType = EventLogEntryType.Warning)
        Dim VLobjVisor As EventLog
        Try
            If _NombreAplicacion Is Nothing Then
                _NombreAplicacion = My.Application.Info.Title
            End If
            If Not EventLog.SourceExists(_NombreAplicacion) Then
                Dim VLobjDatos As EventSourceCreationData
                If _Producto = ProductoKS.KeyMonitor Then
                    VLobjDatos = New EventSourceCreationData(_NombreAplicacion, "Key Monitor")
                Else
                    VLobjDatos = New EventSourceCreationData(_NombreAplicacion, "Perfiles")
                End If
                EventLog.CreateEventSource(VLobjDatos)
            End If
            If _Producto = ProductoKS.KeyMonitor Then
                VLobjVisor = New EventLog("Key Monitor")
            Else
                VLobjVisor = New EventLog("Perfiles")
            End If
            VLobjVisor.Source = _NombreAplicacion
            VLobjVisor.WriteEntry(VPstrMensaje, VPobjTypo)
            VLobjVisor.Close()
            VLobjVisor.Dispose()
        Catch ex As Exception
            ArchivoPlano(VPstrMensaje)
        End Try
    End Sub

    ''' <summary>
    ''' Permite guardar el mensaje en un archivo plano
    ''' </summary>
    Private Sub ArchivoPlano(ByVal VPstrMensaje As String)
        Dim VLioArchivo As System.IO.StreamWriter
        Dim VLstrArchivo As String
        Try
            VLstrArchivo = "KSLOG" & Format(Date.Now, "yyyyMMdd")
            VLioArchivo = New System.IO.StreamWriter(_RutaDirectorio & "\" & VLstrArchivo & ".log", True, System.Text.Encoding.Default)
            VLioArchivo.WriteLine(Format(Date.Now, "dd/MM/yyyy HH:mm:ss") & vbTab & VPstrMensaje)
            VLioArchivo.Close()
            VLioArchivo.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Determina en donde se debe almacenar el mensaje
    ''' </summary>
    Private Sub ProcesarMensaje()
        Dim VLobjMensaje As Mensajes
        Try
            If VMcolaMensajes.Count > 0 Then
                Do
                    SyncLock VMcolaMensajes
                        VLobjMensaje = VMcolaMensajes.Dequeue
                    End SyncLock
                    If _Registro = TipoRegistro.Visor Then
                        VisorEventos(VLobjMensaje.Mensaje, VLobjMensaje.Tipo)
                    ElseIf _Registro = TipoRegistro.Archivo Then
                        ArchivoPlano(VLobjMensaje.Mensaje)
                    ElseIf _Registro = TipoRegistro.Administrador Then
                        '<STX>ID<FS> Mensaje<ETX>
                        EnviarMensaje(STX & _IdConsola & FS & VLobjMensaje.Mensaje & ETX)
                    End If
                Loop While VMcolaMensajes.Count > 0
            End If
        Catch ex As Exception
            ArchivoPlano("Problema al procesar la depuracion: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Envia los mensajes a la Consola
    ''' </summary>
    Private Sub EnviarAdministrador()
        Dim VLstrMensaje As String

        Try
            Do
                SyncLock VMcolaAdmin
                    VLstrMensaje = VMcolaAdmin.Dequeue
                End SyncLock
                Try
Intentar:           If VMtcpAdministrador.Conectado Then
                        VMtcpAdministrador.EnviarDatos(VLstrMensaje)
                    ElseIf VMcolaAdmin.Count > 100 Or _Teminar Then
                        ArchivoPlano("No se puede enviar el mensaje al administrador ya que no se encuentra conectado" & vbNewLine & "Mensaje: " & VLstrMensaje)
                    Else
                        Threading.Thread.Sleep(100)
                        GoTo Intentar
                    End If
                Catch ex As Exception
                    ArchivoPlano("Problema al enviar los mensajes al administrador: " & ex.Message)
                End Try
            Loop While VMcolaAdmin.Count > 0
        Catch ex As Exception
            ArchivoPlano("Problema al procesar los mensajes que se envian al administrador: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' DE LA CARPETA ..\LOG DEJA LOS ARCHIVOS CON FECHA DE CREACIÓN HASTA 14 DÍAS ANTERIORES
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DepurarLog()
        Try
            If My.Computer.FileSystem.DirectoryExists(_RutaDirectorio) Then
                Dim VLobjFso As Object = CreateObject("scripting.filesystemobject")
                Dim VLobjCarpeta As Object = VLobjFso.getfolder(_RutaDirectorio)
                Dim VLobjArchivos As Object = VLobjCarpeta.files

                For Each VLobjArch As Object In VLobjArchivos
                    If DateDiff(DateInterval.Day, VLobjArch.datecreated, Now()) >= 15 Then
                        My.Computer.FileSystem.DeleteFile(VLobjArch.path)
                    End If
                Next
            End If
        Catch ex As Exception
            ArchivoPlano("Problema al ejecutar la depuración: " & ex.Message)
        End Try
    End Sub

    Private Sub VMtmrLog_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles VMtmrLog.Elapsed
        DepurarLog()
        VMtmrLog.Interval = CalculoTiempo()
        VMtmrLog.Start()
    End Sub

    ''' <summary>
    ''' SE CALCULA EL TIEMPO PARA DEPURAR CARPETA \log
    ''' </summary>
    Private Function CalculoTiempo() As Double
        Try
            Dim VLintSegInicia, VLintSegApli, VLintMinInicia, VLintMinApli, VLintHoraInicia, VLintHoraApli As Integer
            Dim VLdteHora As DateTime

            VLdteHora = Format(Now, "HH:mm:ss")

            VLintSegInicia = Format(CDate("03:00:00 AM"), "ss")
            VLintSegApli = Format(VLdteHora, "ss")
            VLintMinInicia = Format(CDate("03:00:00 AM"), "mm")
            VLintMinApli = Format(VLdteHora, "mm")
            VLintHoraInicia = Format(CDate("03:00:00 AM"), "HH")
            VLintHoraApli = Format(VLdteHora, "HH")

            If VLintSegInicia < VLintSegApli Then
                VMintTiempo = (VLintSegInicia - VLintSegApli + 60) * 1000
                VLintMinApli += 1
                If VLintMinInicia < VLintMinApli Then
                    VMintTiempo = VMintTiempo + (VLintMinInicia - VLintMinApli + 60) * 60 * 1000
                    VLintHoraApli += 1
                    If VLintHoraInicia < VLintHoraApli Then
                        VMintTiempo = VMintTiempo + (VLintHoraInicia - VLintHoraApli + 24) * 60 * 60 * 1000
                    Else
                        VMintTiempo = VMintTiempo + (VLintHoraInicia - VLintHoraApli) * 60 * 60 * 1000
                    End If
                Else
                    VMintTiempo = VMintTiempo + (VLintMinInicia - VLintMinApli) * 60 * 1000
                End If
            Else
                VMintTiempo = (VLintSegInicia - VLintSegApli) * 1000
            End If
        Catch ex As Exception
            ArchivoPlano("KONV 01001 - Se produjo un error al calcular el tiempo por el siguiente motivo: " & vbNewLine & ex.Message)
        End Try
        Return VMintTiempo
    End Function
#End Region

#Region "Funciones Publicas"
    Public Sub New()
        VMcolaAdmin = New Queue
        VMcolaMensajes = New Queue
        VMhiloAdmin = New Threading.Thread(AddressOf EnviarAdministrador)
        VMhiloProcesar = New Threading.Thread(AddressOf ProcesarMensaje)
        _TipoDepuracion = TipoDepuracion.Normal
        _Registro = TipoRegistro.Visor
        _RutaDirectorio = My.Application.Info.DirectoryPath & "\log"
        If Not My.Computer.FileSystem.DirectoryExists(_RutaDirectorio) Then
            My.Computer.FileSystem.CreateDirectory(_RutaDirectorio)
        End If

        VMtmrLog = New Timers.Timer
        VMtmrLog.AutoReset = False
        VMtmrLog.Interval = 100
        VMtmrLog.Start()
    End Sub
    ''' <summary>
    ''' Agrega un mensaje para que sea almecenado dependiendo el tipo de depuración
    ''' </summary>
    Public Sub AgregarMensaje(ByVal VPstrMensaje As String, Optional ByVal VPobjTipoEvento As System.Diagnostics.EventLogEntryType = EventLogEntryType.Warning)
        Dim VLobjDepuracion As Mensajes
        Try
            VLobjDepuracion.Mensaje = VPstrMensaje
            VLobjDepuracion.Tipo = VPobjTipoEvento
            SyncLock VMcolaMensajes
                VMcolaMensajes.Enqueue(VLobjDepuracion)
            End SyncLock
            If Not VMhiloProcesar.IsAlive Then
                VMhiloProcesar = New Threading.Thread(AddressOf ProcesarMensaje)
                VMhiloProcesar.Name = "Procesar depuración"
                VMhiloProcesar.Start()
            End If
        Catch ex As Exception
            ArchivoPlano("Problema al agregar el mensaje: " & ex.Message & vbNewLine & "Mensaje: " & VPstrMensaje)
        End Try
    End Sub
    ''' <summary>
    ''' Permite enviarle el mensaje al adminstrador
    ''' </summary>
    Public Sub EnviarMensaje(ByVal Mensaje As String)
        Try
            SyncLock VMcolaAdmin
                VMcolaAdmin.Enqueue(Mensaje)
            End SyncLock
            If Not VMhiloAdmin.IsAlive Then
                VMhiloAdmin = New Threading.Thread(AddressOf EnviarAdministrador)
                VMhiloAdmin.Name = "Mensajes administrador"
                VMhiloAdmin.Start()
            End If
        Catch ex As Exception
            ArchivoPlano("Problema al acumular los mensajes del administrador: " & ex.Message & vbNewLine & "Mensaje: " & Mensaje)
        End Try
    End Sub
    ''' <summary>
    ''' Termina todos los procesos pendientes
    ''' </summary>
    Public Sub Cerrar()
        _Teminar = True

        If VMtmrLog.Enabled Then
            VMtmrLog.Stop()
        End If

    End Sub
#End Region


End Class
