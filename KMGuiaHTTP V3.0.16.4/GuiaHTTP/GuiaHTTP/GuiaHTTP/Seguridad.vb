Imports System.Security.Cryptography
Imports System.IO
Imports System.Text

Module Seguridad
    Public Function encriptaValor(ByVal VLstrValor As String) As String
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
    Public Function desencriptaValor(ByVal VLstrValor As String) As String
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
End Module
