Module Comunicaciones

    Public Const STX As String = ChrW(2)
    Public Const ETX As String = ChrW(3)
    Public Const SI As String = ChrW(15)
    Public Const SO As String = ChrW(14)
    Public Const FS As String = ChrW(28)
    Public Const ACK As String = ChrW(6)
    Public Const NAK As String = ChrW(21)

    Public Function ObtieneLRC(ByVal VPstrCadena As String) As String
        Dim VLvarRes
        Dim VLintI As Integer

        VLvarRes = Asc(Left(VPstrCadena, 1))
        VLvarRes = VLvarRes Xor Asc(Mid(VPstrCadena, 2, 1))
        For VLintI = 3 To Len(VPstrCadena)
            VLvarRes = VLvarRes Xor Asc(Mid(VPstrCadena, VLintI, 1))
        Next
        obtieneLRC = Chr(VLvarRes)
    End Function

    Public Function ConfiabilidadDatos(ByVal VPstrCadena As String) As Boolean
        Dim VLstrLRC As String
        Dim VLintInicio As Integer
        Dim VLintFinal As Integer
        Dim VLintLen As Integer
        Dim VLstrMensaje As String
        Dim VLstrDatos() As String

        VLintInicio = InStr(VPstrCadena, STX)
        VLintFinal = InStr(VPstrCadena, ETX) - 1
        VLintLen = VLintFinal - VLintInicio
        VLstrMensaje = VPstrCadena.Substring(VLintInicio, VLintLen)
        VLstrDatos = Split(VLstrMensaje, FS)
        VLstrLRC = VLstrDatos(VLstrDatos.Length - 1)
        ReDim Preserve VLstrDatos(VLstrDatos.Length - 2)
        VLstrMensaje = Join(VLstrDatos, FS) & FS
        If VLstrLRC = ObtieneLRC(VLstrMensaje) Then
            Return True
        Else
            Return False
        End If
    End Function
End Module
