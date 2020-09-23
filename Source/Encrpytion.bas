Attribute VB_Name = "Encrpytion"
Option Explicit

'This is the Encrypt function (pretty basic but does the job)
Public Function Encrypt(StringToEncrypt As String) As String
    On Error GoTo ErrorHandler
    Dim Char As String, i
    Encrypt = ""
    
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i

    Exit Function
ErrorHandler:
    Encrypt = "Error encrypting string"
End Function

'This is the Decrypt function (pretty basic but does the job)
Public Function Decrypt(StringToDecrypt As String) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    
    Decrypt = ""
    
    Do
        CharPos = VBA.Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = VBA.Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
        
    Loop Until StringToDecrypt = ""
    Exit Function
ErrorHandler:
    Decrypt = "Error decrypting string"
End Function
