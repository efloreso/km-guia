Imports Microsoft.Win32

Module Base_de_Datos

    


#Region "Variables de la Base de datos"
    Public VGstrErrorBase As String = ""
    Private VMbaseBase As tipBD
    'Dim VMrc4Encripta As RC4
#End Region
    Public Function cadenaConexion(ByVal datosBase As tipBD) As String
        Dim VLstrCadena As String

        With datosBase
            Select Case .Tipo
                Case TipoBase.Oracle
                    'VLstrCadena = "Server = " & .Servidor & ":" & .Puerto & _
                    '              ";uid = " & .Uusario & ";pwd = " & _
                    '              .Contrasena
                    '";database = " & .Base
                    VLstrCadena = "SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & .Servidor & ")(PORT=" & .Puerto & "))(CONNECT_DATA=(SERVICE_NAME=" & .Instancia & ")));uid=" & .Usuario & ";pwd=" & .Contrasena & ";"
                Case TipoBase.MySQL
                    VLstrCadena = "Server = " & .Servidor & ";uid = " & .Usuario & _
                                  ";pwd = " & .Contrasena & ";port = " & .Puerto & _
                                  ";database = " & .Base
                Case TipoBase.SQLServer
                    If .SI Then

                        VLstrCadena = "Data Source = " & .Servidor & "," & .Puerto & _
                                                 ";Initial Catalog = " & .Base & _
                                                 ";Integrated Security=SSPI; Connection Timeout=45;"
                    Else
                        VLstrCadena = "Data Source = " & .Servidor & "," & .Puerto & _
                                                     ";Initial Catalog = " & .Base & _
                                                     ";uid = " & .Usuario & _
                                                     ";pwd = " & .Contrasena
                        '",Port = " & txtPuertoTransaccional.Text
                    End If
                Case Else
                    VLstrCadena = ""
            End Select
        End With

        Return VLstrCadena
    End Function
    
    Public Function CargarBase(ByVal VPstrRegistro As String) As tipBD
        '**********************************************************************************************
        '       Carga una base de datos dependiendo de la ruta que se le de
        '**********************************************************************************************
        Dim VLkeyRegistro As RegistryKey    'Guarda la ruta donde se encuentran los datos de la base transaccional
        VGstrErrorBase = ""
        VLkeyRegistro = Registry.LocalMachine.OpenSubKey(VPstrRegistro, False)
        If Not IsNothing(VLkeyRegistro) Then
            Try
                VMbaseBase.Base = IIf(IsNothing(VLkeyRegistro.GetValue("Base")), "", VLkeyRegistro.GetValue("Base"))
                If Not IsNothing(VLkeyRegistro.GetValue("Contrasena")) Then
                    VMbaseBase.Contrasena = IIf(VLkeyRegistro.GetValue("Contrasena").ToString.Trim.Length > 0, desencriptaValor(VLkeyRegistro.GetValue("Contrasena")), "")
                End If
                VMbaseBase.Instancia = VLkeyRegistro.GetValue("Esquema")
                VMbaseBase.SI = VLkeyRegistro.GetValue("SI")
                VMbaseBase.Servidor = IIf(IsNothing(VLkeyRegistro.GetValue("Servidor")), "", VLkeyRegistro.GetValue("Servidor"))
                VMbaseBase.Tipo = IIf(IsNothing(VLkeyRegistro.GetValue("Tipo")), Nothing, VLkeyRegistro.GetValue("Tipo"))
                VMbaseBase.Usuario = VLkeyRegistro.GetValue("Usuario")
                If VMbaseBase.Base.ToString.Trim.Length = 0 Or VMbaseBase.Servidor.ToString.Trim.Length = 0 Or IsNothing(VMbaseBase.Tipo) Then
                    VGstrErrorBase = "Algunos de los datos de la base transaccional no estan completos"
                    Return Nothing
                Else
                    Return VMbaseBase
                End If
            Catch ex As Exception
                VGstrErrorBase = ex.Message
                Return Nothing
            End Try
            VLkeyRegistro.Close()
        Else
            VGstrErrorBase = "No se encontraron datos sobre la base transaccional"
            Return Nothing
        End If
    End Function

End Module
