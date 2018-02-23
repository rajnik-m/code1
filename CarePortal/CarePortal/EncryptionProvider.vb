Public Class EncryptionProvider

  'RC4 stream cipher symmetric key algorithm
  'See http://www.ncat.edu/~grogans/main.htm

  Dim mvS(255) As Integer 'S-Box
  Dim mvKey As String

  Public Sub New()
    Init("CarePortal")
  End Sub

  Public Sub Init(ByRef pClientCode As String)
    mvKey = pClientCode & "PrivateKey"
  End Sub

  Private Sub RC4ini()
    Dim vTemp As Integer
    Dim vKep(255) As Integer
    Dim vA As Integer
    Dim vB As Integer

    'Save Key in Byte-Array
    vB = 0
    For vA = 0 To 255
      vB = vB + 1
      If vB > Len(mvKey) Then vB = 1
      vKep(vA) = Asc(Mid(mvKey, vB, 1))
    Next
    For vA = 0 To 255
      mvS(vA) = vA
    Next
    vB = 0
    For vA = 0 To 255
      vB = (vB + mvS(vA) + vKep(vA)) Mod 256
      ' Swap( mvS(i),mvS(j) )
      vTemp = mvS(vA)
      mvS(vA) = mvS(vB)
      mvS(vB) = vTemp
    Next
  End Sub

  Private Function EnDeCrypt(ByRef pPlainText As String) As String
    Dim vCipherByte As Byte
    Dim vResult As String = ""
    Dim vTemp As Integer
    Dim vA As Integer
    Dim vI As Integer
    Dim vJ As Integer
    Dim vK As Integer

    For vA = 1 To Len(pPlainText)
      vI = (vI + 1) Mod 256
      vJ = (vJ + mvS(vI)) Mod 256
      ' Swap( mvS(vI),mvS(vJ) )
      vTemp = mvS(vI)
      mvS(vI) = mvS(vJ)
      mvS(vJ) = vTemp
      'Generate Keybyte vK
      vK = mvS((mvS(vI) + mvS(vJ)) Mod 256)
      'Plaintextbyte xor Keybyte
      vCipherByte = CByte(Asc(Mid(pPlainText, vA, 1)) Xor vK)
      vResult = vResult & Chr(vCipherByte)
    Next
    EnDeCrypt = vResult
  End Function

  Public Function Decrypt(ByRef pEncryptedString As String) As String
    Dim vIndex As Integer
    Dim vSource As String = ""

    RC4ini()
    For vIndex = 1 To Len(pEncryptedString) Step 2
      vSource = vSource & GetCharFromHex(Mid(pEncryptedString, vIndex, 2))
    Next
    Decrypt = EnDeCrypt(vSource)
  End Function

  Private Function GetCharFromHex(ByRef pHex As String) As String
    Dim vASCII As Integer
    Dim vChar As String

    vChar = UCase(Left(pHex, 1))
    Select Case vChar
      Case "0" To "9"
        vASCII = CInt(vChar) * 16
      Case Else
        vASCII = ((Asc(vChar) - Asc("A")) + 10) * 16
    End Select
    vChar = UCase(Mid(pHex, 2))
    Select Case vChar
      Case "0" To "9"
        vASCII = vASCII + CInt(vChar)
      Case Else
        vASCII = vASCII + ((Asc(vChar) - Asc("A")) + 10)
    End Select
    GetCharFromHex = Chr(vASCII)
  End Function

  Public Function Encrypt(ByRef pSource As String) As String
    Dim vIndex As Integer
    Dim vEncoded As String
    Dim vResult As String = ""

    RC4ini()
    vEncoded = EnDeCrypt(pSource)
    For vIndex = 1 To Len(vEncoded)
      vResult = vResult & GetHexPair(Mid(vEncoded, vIndex, 1))
    Next
    Encrypt = vResult
  End Function

  Private Function GetHexPair(ByRef pChar As String) As String
    Dim vResult As String

    vResult = Hex(Asc(pChar))
    If Len(vResult) = 1 Then vResult = "0" & vResult
    GetHexPair = vResult
  End Function

End Class
