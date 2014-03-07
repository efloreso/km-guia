Imports MySql.Data.MySqlClient
Imports System
Imports System.Net
Imports System.Net.Dns
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports System.Net.Mail

Module GuiaPublica
    Public Enum enuTipoAviso
        Ninguno = 0
        MensajeError
        Informativo
        Advertencia
        Pregunta
    End Enum
    Public Enum enuEscribeLog
        Completo
        Sencillo
        Ambos
    End Enum
    Public Enum TipoBD As Byte
        Central
        Transaccional
        Bitacora
    End Enum
    Public Enum enuTipoMensaje As Byte
        Corto = 1
        Normal = 2
    End Enum
    Public Enum enuServOrigen As Byte
        Estandar = 1
        Stratus = 2
    End Enum
    Public Enum enuTipoSolicitud As Byte
        Si = 1
        No = 2
    End Enum
    Public Enum enuAplicaciones
        KGintAdministrador = 1
        KGintAdminTrans = 2
        KGintAnalizador = 3
        KGintReceptor = 4
        KGintBurbuja = 5
        KGintEstadisticos = 6
        KGintBitacora = 7
        KGintGuiaHttp = 10
        KGintGuiaHttpOff = 12
        KGintBurbujaOff = 13
    End Enum
    Public Structure tipHdrComunicaciones
        Dim Pathway As String '* 16   '16
        Dim Servidor As String '* 15   '31
        Dim Longitud As String '* 5    '36
    End Structure
    Public Structure tipGeneral
        Dim IP As String           'IP del servidor del cual se recibe información del SERVIDOR TRANSACCIONAL
        Dim Puerto As String       'Puerto en el cual se recibe información del SERVIDOR TRANSACCIONAL
        Dim Licencias As String    'Licencias
        Dim PuertoLocal As String  'Puerto Inicial
        Dim Mensaje As enuTipoMensaje 'Tipo Mensaje (corto/normal)
        Dim MinDia As String       'Reinicio dia 
        Dim MinNoche As String      'Reinicio Noche
        Dim InicioDia As DateTime   'Rango Inicial dia 
        Dim InicioNoche As DateTime 'Rango Inicial Noche
        Dim Acknowled As Boolean    'Indica si se va enviar "#"
        Dim Tipo As TipoGuia       'Indica que tipo de guia es
        Dim LongitudE As Integer    'Longitud del encabezado de la transacción
        Dim MensajeOrigen As enuServOrigen 'Servidor Origen (Estandar/ Stratus)
        Dim Confirmaciones As Integer 'Numero de confirmaciones 
    End Structure
    Public Structure tipCorreo
        Dim Servidor As String
        Dim Usuario As String
        Dim Nombre As String
        Dim Cuenta As String
        Dim Contrasena As String
        Dim Autenticacion As Boolean
        Dim Para As String
        Dim CC As String
        Dim CO As String
        Dim Activo As Boolean
    End Structure
    Public Structure tipEspecial
        Dim Confirmar As String             'Confirmar cada X
        Dim Espera As enuTipoSolicitud      'Esperar solicitud del replicador
        Dim Memoria As String
    End Structure
    Public Structure tipLayout
        Dim LayOut As Integer
        Dim Formato As Integer
    End Structure
    Public Structure tipBD
        Dim Tipo As TipoBase
        Dim Servidor As String
        Dim Puerto As Integer
        Dim Base As String
        Dim SI As Boolean
        Dim Usuario As String
        Dim Contrasena As String
        Dim Instancia As String
    End Structure
    Public Structure tipPuertos
        Dim Layout As Integer
        Dim Formato As Integer
        Dim Puerto As Integer
    End Structure
    Public Enum TipoBase As Byte
        Access = 1
        Oracle = 2
        SQLServer = 3
        Informix = 4
        MySQL = 5
        Sybase = 6
    End Enum
    Public Enum TipoGuia
        Linea = 1
        Recuperacion = 2
        Filtros = 3
        Replicacion = 4
    End Enum
    Public Enum enuDepuracion
        Ninguna
        Sencilla
        Completa
        Consola
    End Enum
    Public _Configuracion As String                                 'Nombre de la configuración actual
    Public _TipoGuia As Byte                                       '1=Replicacion,2=Filtros,3=Recuperaciones,4=Linea
    Public VGstrLocalHostName As String
    Public VGstrLocalIPAddress As String
    Public VGstrUsuario As String
    Public VGintAplicacion As String
    Public _Log As enuDepuracion                    'Tipo de depuración del servicio
    Public VGioArchivoLog As System.IO.StreamWriter 'Archivo para la depuración
    Public VPioLogTran As System.IO.StreamWriter    'Archivo para guardar las transacciones en línea
    Public VPblnLogtran As Boolean                  'Si es true indica que se está escribiendo en el archivo
    Public VGbytErrorLog As Byte                    'Guarda los errores consecutivos al escribir en el log (si ocuren 10 errores consecutivos se desactiva la escritura en el log)
    Public VGtipGeneral As tipGeneral               'Contiene la información General (Puertos, Servidores, tipos de mensajes, reinicio dia y noche, )
    Public VGtipCorreo As tipCorreo                'Contiene la información del correo
    Public VGtipLayout As tipLayout                 'Almacena la configuración de la pestaña Layout del registro 
    Public VGtipBDCentral As tipBD                  'Almacena la información de la Base Central
    Public VGtipBDBitacora As tipBD                 'Almacena la información de la Base Bitacora
    Public Sub Notifica(ByVal Mensaje As String, ByVal VPclsReplicador As GuiaHTTP, Optional ByVal TipoLog As enuEscribeLog = enuEscribeLog.Completo, Optional ByVal Tipo As enuTipoAviso = enuTipoAviso.Ninguno, Optional ByVal Complemento As String = "", Optional ByVal Correo As Boolean = False)
        Dim VLobjCorreo As Threading.Thread
        Dim VLstrMensaje As String

        If (TipoLog = enuEscribeLog.Sencillo Or TipoLog = enuEscribeLog.Ambos) And _Log = enuDepuracion.Sencilla Then
            escribeLog(vbNewLine & Mensaje)
        ElseIf (TipoLog = enuEscribeLog.Completo Or TipoLog = enuEscribeLog.Ambos) And _Log = enuDepuracion.Completa Then
            escribeLog(vbNewLine & Mensaje & Complemento)
        ElseIf _Log = enuDepuracion.Consola Then
            VPclsReplicador.escribeConsola(Mensaje)
        End If

        If Tipo <> enuTipoAviso.Ninguno Then
            Dim VLobjTipo As EventLogEntryType
            Select Case Tipo
                Case enuTipoAviso.MensajeError
                    VLobjTipo = EventLogEntryType.Error
                Case enuTipoAviso.Informativo
                    VLobjTipo = EventLogEntryType.Information
                Case enuTipoAviso.Advertencia
                    VLobjTipo = EventLogEntryType.Warning
            End Select
            escribeVisorEventos(Mensaje, VLobjTipo)
        End If

        VLstrMensaje = Mensaje & "||" & Tipo
        If Correo Then
            VLobjCorreo = New Threading.Thread(AddressOf enviaCorreo)
            VLobjCorreo.Start(VLstrMensaje)
        End If
    End Sub
    Public Sub escribeLog(ByVal Mensaje As String)
        Try
            VGioArchivoLog.WriteLine(Mensaje)
            VGioArchivoLog.Flush()
            VGbytErrorLog = 0
        Catch ex As Exception
            VGbytErrorLog += 1
            If VGbytErrorLog >= 10 Then
                _Log = enuDepuracion.Ninguna
                escribeVisorEventos("Se inició el servicio en modo de depuración pero no se ha podido escribir en el archivo por el siguiente motivo:" & vbNewLine & ex.Message & vbNewLine & "Intento #" & VGbytErrorLog & " se deshabilitará la escritura del archivo", EventLogEntryType.Error)
            Else
                escribeVisorEventos("Se inició el servicio en modo de depuración pero no se ha podido escribir en el archivo por el siguiente motivo:" & vbNewLine & ex.Message, EventLogEntryType.Warning)
            End If
        End Try

    End Sub
    Function RetIPAddress(ByVal mStrHost As String) As String
        Dim VLipEquipo As IPHostEntry = GetHostEntry(My.Computer.Name)
        Dim VLipListaIPs As IPAddress() = VLipEquipo.AddressList()
        Dim VLobjIPLocal As Object
        Dim VLstrIP As String

        VLstrIP = "No disponible"
        For Each VLobjIPLocal In VLipListaIPs
            If VLobjIPLocal.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                VLstrIP = VLobjIPLocal.ToString
                Exit For
            End If
        Next

        Return VLstrIP
    End Function
    Public Sub abreArchivoLog(Optional ByVal Visor As Boolean = False)
        Try
            VGioArchivoLog = New StreamWriter(My.Application.Info.DirectoryPath & "\MS" & Format(Now, "ddMMyyyy") & ".dat", True)
            VGioArchivoLog.WriteLine("-----INICIO DEPURACIÓN " & Format(Now, "hh:mm:ss"))
            VGioArchivoLog.Flush()
            If Not Visor Then
                escribeVisorEventos("Se ha iniciado el servicio en modo de depuración", EventLogEntryType.Information)
            End If
        Catch ex As Exception
            If Not Visor Then
                escribeVisorEventos("Se ha iniciado el servicio en modo de depuracion pero no se pudo crear el archivo de salida" & vbNewLine & ex.Message, EventLogEntryType.Warning)
                _Log = enuDepuracion.Ninguna
            End If
        End Try
    End Sub
    Public Sub escribeVisorEventos(ByVal Mensaje As String, ByVal Tipo As System.Diagnostics.EventLogEntryType)
        Dim VLstrTipo As String

        Select Case _TipoGuia
            Case 2
                VLstrTipo = "de Filtros "
            Case 1
                VLstrTipo = "de Replicación"
            Case 3
                VLstrTipo = "de Recuperación "
            Case Else
                'Linea
                VLstrTipo = "HTTP "
        End Select
        Try
            If IsNothing(_Configuracion) Then
                _Configuracion = "S/D  "

            End If
            Diagnostics.EventLog.WriteEntry("Key Monitor Guia " & VLstrTipo & " " & _Configuracion.Substring(0, _Configuracion.Length - 2).Trim, Mensaje.Trim, Tipo)
        Catch ex As Exception
            abreArchivoLog(True)
            escribeLog("No se pudo escribir en el visor " & ex.Message & vbNewLine & "Key Monitor Guia " & VLstrTipo & _Configuracion.Substring(0, _Configuracion.Length - 2).Trim & Mensaje)
        End Try
    End Sub
    Public Function claveRutaUnica() As String
        '********************************************************************************
        'Obtiene un identificador único para el nombre y ruta del archivo ejecutable    '
        '---VALORES DEVUELTOS                                                           '
        '   Cadena con el identificador único                                           '
        '********************************************************************************

        Dim VLintX As Integer
        Dim VLlngLlave As Double
        Dim VLstrRuta As String
        Dim VLintUbica As Integer

        VLintUbica = Strings.InStr(Strings.StrReverse(My.Application.Info.DirectoryPath), "\")
        VLstrRuta = Strings.Left(My.Application.Info.DirectoryPath, Len(My.Application.Info.DirectoryPath) - VLintUbica)
        For VLintX = 1 To VLstrRuta.Length - 1
            VLlngLlave = VLlngLlave + (Asc(VLstrRuta.Substring(VLintX - 1, 1)) * VLintX)
        Next VLintX
        Return "?" & VLlngLlave.ToString.Trim & "?"
    End Function
    Private Function encriptaValor(ByVal VLstrValor As String) As String
        '******************************************************************************
        'Función encriptar Valores 
        'Parámetros:
        '              VLstrValor= Valor a encriptar
        'Valores devueltos:
        '              Cadena encriptada       
        '******************************************************************************
        Dim VLstrCadenaEncriptada As String
        Dim VLstrSemilla As String
        Dim VLstrLlave As String
        Dim VLbytSemilla() As Byte
        Dim VLbytLlave() As Byte
        Dim VLbytCadena() As Byte
        Dim memStream As MemoryStream
        Dim VLcryEncripta As TripleDESCryptoServiceProvider
        Dim VLcryTransforma As ICryptoTransform
        Dim VLcryCS As CryptoStream


        VLstrCadenaEncriptada = ""
        VLstrLlave = "%KeY*MoNiTor%V30"
        VLstrSemilla = "**TrA*&nSaCcI?oNal**"

        VLbytSemilla = Encoding.UTF8.GetBytes(VLstrSemilla)
        VLbytLlave = Encoding.UTF8.GetBytes(VLstrLlave)

        memStream = Nothing
        VLbytCadena = Encoding.UTF8.GetBytes(VLstrValor)

        memStream = New MemoryStream(VLstrValor.Length * 2)
        VLcryEncripta = New TripleDESCryptoServiceProvider
        VLcryTransforma = VLcryEncripta.CreateEncryptor(VLbytLlave, VLbytSemilla)
        VLcryCS = New CryptoStream(memStream, VLcryTransforma, CryptoStreamMode.Write)
        VLcryCS.Write(VLbytCadena, 0, VLbytCadena.Length)
        VLcryCS.Close()

        VLstrCadenaEncriptada = Convert.ToBase64String(memStream.ToArray)

        encriptaValor = VLstrCadenaEncriptada
    End Function
    Public Function obtieneErrorBD(ByVal VPobjError As MySqlException) As String
        Select Case VPobjError.Number
            Case 1042
                obtieneErrorBD = "El servidor remoto no se encuentra o no está disponible"
            Case 1045
                obtieneErrorBD = "Nombre de usuario y/o contraseña incorrecto(s)"
            Case 1049
                obtieneErrorBD = "Nombre de base de datos desconocido [" & VPobjError.Message.Substring(16).Trim & "]"
            Case Else
                obtieneErrorBD = "[#" & VPobjError.Number & "#]" & VPobjError.Message
        End Select
        obtieneErrorBD = VPobjError.Message
    End Function
    Public Function obtieneErrorBD(ByVal VPobjError As SqlClient.SqlException) As String
        obtieneErrorBD = VPobjError.Message
    End Function
    Private Function desencriptaValor(ByVal VLstrValor As String) As String
        '******************************************************************************
        'Función desencriptar Valores 
        'Parámetros:
        '              VLstrValor= Valor a desencriptar
        'Valores devueltos:
        '              Cadena desencriptada       
        '******************************************************************************
        Dim VLstrLlave As String
        Dim VLbytLlave() As Byte
        Dim VLstrSemilla As String
        Dim VLbytSemilla() As Byte
        Dim memStream As MemoryStream
        Dim VLbytCadena() As Byte
        Dim VLcryEncripta As TripleDESCryptoServiceProvider
        Dim VLcryTransforma As ICryptoTransform
        Dim VLcryCS As CryptoStream
        Dim VLstrCadenaEncriptada As String


        VLstrLlave = "%KeY*MoNiTor%V30"
        VLstrSemilla = "**TrA*&nSaCcI?oNal**"
        Try

            VLbytSemilla = Encoding.UTF8.GetBytes(VLstrSemilla)
            VLbytLlave = Encoding.UTF8.GetBytes(VLstrLlave)
            memStream = Nothing
            VLbytCadena = Convert.FromBase64String(VLstrValor)
            memStream = New MemoryStream(VLstrValor.Length)
            VLcryEncripta = New TripleDESCryptoServiceProvider
            VLcryTransforma = VLcryEncripta.CreateDecryptor(VLbytLlave, VLbytSemilla)
            VLcryCS = New CryptoStream(memStream, VLcryTransforma, CryptoStreamMode.Write)
            VLcryCS.Write(VLbytCadena, 0, VLbytCadena.Length)
            VLcryCS.Close()
            VLstrCadenaEncriptada = Encoding.UTF8.GetString(memStream.ToArray)
        Catch ex As Exception
            VLstrCadenaEncriptada = "ERROR"
        End Try
        desencriptaValor = VLstrCadenaEncriptada
    End Function
    Public Sub enviaCorreo(ByVal Mensaje As String)
        Dim VLobjCorreo As New MailMessage
        Dim VLobjCliente As New SmtpClient
        Dim VLstrTemp() As String
        Dim VLintI As Integer
        Dim VLobjCredencial As Net.NetworkCredential
        Dim VLstrTipo As String
        Dim VLstrMensaje As String
        Dim VLobjVistaAlterna As AlternateView
        Dim VLstrTipoMensaje As String
        Dim VLstrDatos() As String

        VLstrDatos = Split(Mensaje, "||")

        Select Case _TipoGuia
            Case 2
                VLstrTipo = "de Filtros "
            Case 1
                VLstrTipo = "de Replicación"
            Case 3
                VLstrTipo = "de Recuperación "
            Case Else
                'Linea
                VLstrTipo = "HTTP "
        End Select

        If VLstrDatos(1) = 1 Then
            VLstrTipoMensaje = " Error"
        ElseIf VLstrDatos(1) = 2 Or VLstrDatos(1) = 0 Then
            VLstrTipoMensaje = " Informativo"
        Else
            VLstrTipoMensaje = " Advertencia"
        End If

        If VGtipCorreo.Activo Then
            'Arma el mensaje
            VLstrMensaje = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
            VLstrMensaje &= "<html xmlns=""http://www.w3.org/1999/xhtml"" >"
            VLstrMensaje &= "   <head>"
            VLstrMensaje &= "       <title>Mensaje Consola Key Monitor</title>"
            VLstrMensaje &= "   </head>"
            VLstrMensaje &= "   <body> "
            VLstrMensaje &= "       <table style=""width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 14px;"">"
            VLstrMensaje &= "           <tr>"
            VLstrMensaje &= "               <td><b>Aplicación</b></td>"
            VLstrMensaje &= "               <td>Key Monitor Guia " & VLstrTipo & "</td>"
            VLstrMensaje &= "           </tr>"
            VLstrMensaje &= "           <tr>"
            VLstrMensaje &= "               <td><b>Configuración</b></td>"
            VLstrMensaje &= "               <td>" & _Configuracion.Substring(0, _Configuracion.Length - 2).Trim & "</td>"
            VLstrMensaje &= "           </tr>"
            VLstrMensaje &= "            <tr>"
            VLstrMensaje &= "               <td><b>Fecha</b></td>"
            VLstrMensaje &= "               <td>" & Format(Now, "dd - MMMM - yyyy") & "</td>"
            VLstrMensaje &= "           </tr>"
            VLstrMensaje &= "           <tr>"
            VLstrMensaje &= "               <td><b>Hora</b></td>"
            VLstrMensaje &= "               <td>" & Format(Now, "HH:mm:ss") & "</td>"
            VLstrMensaje &= "           </tr>"
            VLstrMensaje &= "           <tr>"
            VLstrMensaje &= "               <td valign=""top""><b>Mensaje</b></td>"
            VLstrMensaje &= "               <td>" & VLstrDatos(0) & "</td>"
            VLstrMensaje &= "           </tr>"
            VLstrMensaje &= "       </table>"
            VLstrMensaje &= "   </body>"
            VLstrMensaje &= "</html>"

            'Crea la vista alternativa
            VLobjVistaAlterna = AlternateView.CreateAlternateViewFromString(VLstrMensaje, Nothing, "text/html")
            VLobjCorreo.AlternateViews.Add(VLobjVistaAlterna)

            Try
                VLobjCorreo.From = New MailAddress(VGtipCorreo.Cuenta, "Key Monitor Administrador Central")
                VLobjCorreo.To.Add(VGtipCorreo.Cuenta)

                If VGtipCorreo.Para.Trim.Length > 0 Then
                    VLstrTemp = Split(VGtipCorreo.Para, ";")
                    If VLstrTemp.Length >= 0 Then
                        For VLintI = 0 To VLstrTemp.Length - 1
                            VLobjCorreo.To.Add(VLstrTemp(VLintI))
                        Next
                    End If
                End If

                If VGtipCorreo.CC.Trim.Length > 0 Then
                    VLstrTemp = Split(VGtipCorreo.CC, ";")
                    If VLstrTemp.Length >= 0 Then
                        For VLintI = 0 To VLstrTemp.Length - 1
                            VLobjCorreo.CC.Add(VLstrTemp(VLintI))
                        Next
                    End If
                End If

                If VGtipCorreo.CO.Trim.Length > 0 Then
                    VLstrTemp = Split(VGtipCorreo.CO, ";")
                    If VLstrTemp.Length >= 0 Then
                        For VLintI = 0 To VLstrTemp.Length - 1
                            VLobjCorreo.Bcc.Add(VLstrTemp(VLintI))
                        Next
                    End If
                End If

                VLobjCliente.Host = VGtipCorreo.Servidor
                If VGtipCorreo.Autenticacion Then
                    VLobjCredencial = New Net.NetworkCredential(VGtipCorreo.Usuario, VGtipCorreo.Contrasena)
                    VLobjCliente.Credentials = VLobjCredencial
                End If

                VLobjCorreo.Subject = "Reporte de eventos de Key Monitor [" & VLstrTipoMensaje & "]" '"Key Monitor Guia " & VLstrTipo & " " & _Configuracion.Substring(0, _Configuracion.Length - 2).Trim
                VLobjCorreo.Priority = MailPriority.High
                VLobjCorreo.Body = VLstrMensaje
                Try
                    VLobjCliente.Send(VLobjCorreo)
                Catch ex As Exception
                    escribeVisorEventos("No se pudo enviar el correo por el siguiente motivo:  " & ex.Message, EventLogEntryType.Warning)
                End Try
            Catch ex As Exception
                escribeVisorEventos("No se pudo enviar el correo por el siguiente motivo:  " & ex.Message, EventLogEntryType.Warning)
            End Try
        End If

    End Sub
#Region "Bitacora"
    Function guardaBitacora(ByVal VPstrMensaje As String, ByVal VPintRegistro As Integer, Optional ByVal VPintStatus _
                        As Integer = 0, Optional ByVal VPstrDescFiltro As String = "", Optional ByVal VPstrEscala As String = "" _
                        , Optional ByVal VpstrLimite As String = "", Optional ByVal VPstrTiempo As String = "", Optional _
                        ByVal VPstrReferencia As String = "", Optional ByVal VPintFiltro As Integer = 0, Optional _
                        ByVal VPstrCondicion As String = "") As Boolean

        Dim VLstrSQL As String
        Dim VLdbsBase As Object
        Dim VLcmdComando As Object
        Dim VLrcsDatos As Object
        Dim KGintaPP As Integer
        Dim VGstrLocalHostName As String
        Dim VGstrLocalIPAddress As String
        VLdbsBase = Nothing
        VLcmdComando = Nothing

        Select Case VGtipBDBitacora.Tipo
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
        VLdbsBase.connectionstring = cadenaConexion(VGtipBDBitacora)
        Try
            VLdbsBase.open()
            If VGtipGeneral.Tipo = TipoGuia.Linea Or VGtipGeneral.Tipo = TipoGuia.Filtros Then
                KGintaPP = 10
            ElseIf VGtipGeneral.Tipo = TipoGuia.Recuperacion Then
                KGintaPP = 12
            End If
            VGstrLocalHostName = My.Computer.Name
            VGstrLocalIPAddress = RetIPAddress("localhost")

            'Checar en la tabla incremntar el campo Descripcion_Filtro
            If VPstrDescFiltro.Length > 50 Then
                VPstrDescFiltro = Strings.Left(VPstrDescFiltro, 50)
            End If



            VLstrSQL = ""
            Select Case VGtipBDBitacora.Tipo
                Case TipoBase.MySQL
                    VLstrSQL = "INSERT INTO TBITALARMA VALUES (" & Val(VGstrUsuario) & "," & VGtipLayout.LayOut & "," & _
                                                  VGtipLayout.Formato & "," & KGintaPP & "," & VPintRegistro & ",'" & VGstrLocalIPAddress & _
                                                  "','" & VGstrLocalHostName & "',NOW()," & VPintFiltro & ",'" & VPstrDescFiltro & _
                                                  "','" & VPstrEscala & "','" & Replace(VPstrMensaje, "'", "") & "','" & VpstrLimite & "','" & _
                                                  VPstrTiempo & "','" & VPstrReferencia & "'," & VPintStatus & ",'" & VPstrCondicion & "')"
                    VLcmdComando = New MySqlCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.Oracle
                    VLstrSQL = "INSERT INTO TBITALARMA VALUES (" & Val(VGstrUsuario) & "," & VGtipLayout.LayOut & "," & _
                               VGtipLayout.Formato & "," & KGintaPP & "," & VPintRegistro & ",'" & VGstrLocalIPAddress & _
                               "','" & VGstrLocalHostName & "',GETDATE()," & VPintFiltro & ",'" & VPstrDescFiltro & _
                               "','" & VPstrEscala & "','" & Replace(VPstrMensaje, "'", "") & "','" & VpstrLimite & "','" & _
                               VPstrTiempo & "','" & VPstrReferencia & "'," & VPintStatus & ",'" & VPstrCondicion & "')"

                    VLcmdComando = New OracleCommand(VLstrSQL, VLdbsBase)
                Case TipoBase.SQLServer
                    VLstrSQL = "INSERT INTO TBITALARMA VALUES (" & Val(VGstrUsuario) & "," & VGtipLayout.LayOut & "," & _
                               VGtipLayout.Formato & "," & KGintaPP & "," & VPintRegistro & ",'" & VGstrLocalIPAddress & _
                               "','" & VGstrLocalHostName & "',GETDATE()," & VPintFiltro & ",'" & VPstrDescFiltro & _
                               "','" & VPstrEscala & "','" & Replace(VPstrMensaje, "'", "") & "','" & VpstrLimite & "','" & _
                               VPstrTiempo & "','" & VPstrReferencia & "'," & VPintStatus & ",'" & VPstrCondicion & "')"

                    VLcmdComando = New SqlCommand(VLstrSQL, VLdbsBase)
            End Select

            If VPstrDescFiltro <> "" Then
                VPstrDescFiltro = " Se ejecuto filtro: " & VPstrDescFiltro
            End If

            Try
                VLrcsDatos = VLcmdComando.ExecuteReader
                VLrcsDatos.Close()
                VLcmdComando.Dispose()
            Catch ex As MySqlException
                'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & " Error Bitacora:" & ex.Message)
            Catch ex As OracleException
                'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & " Error Bitacora:" & ex.Message)
            Catch ex As SqlException
                'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & " Error Bitacora:" & ex.Message)
            End Try
            VLdbsBase.close()
        Catch ex As MySqlException
            'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & "Error Bitacora:" & ex.Message)
        Catch ex As OracleException
            'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & "Error Bitacora:" & ex.Message)
        Catch ex As SqlException
            'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & "Error Bitacora:" & ex.Message)
        Catch ex As Exception
            'bitacoraTexto(VPstrMensaje & " " & VPstrDescFiltro & "Error Bitacora:" & ex.Message)
        End Try
        Try
            VLdbsBase.close()

        Catch ex As Exception

        End Try

    End Function
#End Region
End Module
