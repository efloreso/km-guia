Public Class RC4
    'CLASE PARA ENCRIPTAR CADENAS EN FORMATO RC4
    Public Function RC4Encrypt(ByVal text As String, ByVal encryptkey As String)
        Dim sbox(256)
        Dim key(256)
        Dim temp As Integer
        Dim a As Long
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim cipherby As Integer
        Dim cipher As String

        i = 0
        j = 0
        cipher = ""

        RC4Initialize(encryptkey, key, sbox)

        For a = 1 To Len(text)
            i = (i + 1) Mod 256
            j = (j + sbox(i)) Mod 256
            temp = sbox(i)
            sbox(i) = sbox(j)
            sbox(j) = temp

            k = sbox((sbox(i) + sbox(j)) Mod 256)

            cipherby = (Asc(Mid$(text, a, 1))) Xor k
            If Len(Hex(cipherby)) = 1 Then
                cipher = cipher & "0" & Hex(cipherby)
            Else
                cipher = cipher & Hex(cipherby)
            End If
        Next

        RC4Encrypt = cipher
    End Function
    Public Function RC4Decrypt(ByVal text, ByVal encryptkey)
        Dim sbox(256)
        Dim key(256)
        Dim Text2 As String
        Dim temp As Integer
        Dim a As Long
        Dim i As Integer
        Dim j As Integer
        Dim w As Integer
        Dim k As Integer
        Dim cipherby As Integer
        Dim cipher As String

        Text2 = ""
        cipher = ""
        For w = 1 To Len(text) Step 2
            Text2 = Text2 & Chr(Dec(Mid$(text, w, 2)))
        Next

        i = 0
        j = 0

        RC4Initialize(encryptkey, key, sbox)

        For a = 1 To Len(Text2)
            i = (i + 1) Mod 256
            j = (j + sbox(i)) Mod 256
            temp = sbox(i)
            sbox(i) = sbox(j)
            sbox(j) = temp

            k = sbox((sbox(i) + sbox(j)) Mod 256)

            cipherby = Asc(Mid$(Text2, a, 1)) Xor k
            cipher = cipher & Chr(cipherby)
        Next

        RC4Decrypt = cipher
    End Function
    Public Sub RC4Initialize(ByVal strPwd, ByRef key, ByRef sbox)
        Dim tempSwap
        Dim a
        Dim b
        Dim intlength

        intlength = Len(strPwd)
        For a = 0 To 255
            key(a) = Asc(Mid$(strPwd, a Mod intlength + 1, 1))
            sbox(a) = a
        Next

        b = 0
        For a = 0 To 255
            b = (b + sbox(a) + key(a)) Mod 256
            tempSwap = sbox(a)
            sbox(a) = sbox(b)
            sbox(b) = tempSwap
        Next
    End Sub
    Public Function Dec(ByVal number) As String
        Dim base As String
        Dim iLen As Integer
        Dim iReturn As Long
        Dim i
        Dim iTemp

        base = "0123456789ABCDEF"
        iLen = Len(number)

        For i = 0 To iLen - 1
            iTemp = Mid$(number, iLen - i, 1)
            iReturn = iReturn + (16 ^ i) * (InStr(1, base, iTemp) - 1)
        Next

        Dec = iReturn
    End Function
End Class
