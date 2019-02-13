Attribute VB_Name = "modHash"
Option Explicit

Private Declare Function CryptAcquireContext _
                Lib "advapi32.dll" _
                Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                              ByVal pszContainer As String, _
                                              ByVal pszProvider As String, _
                                              ByVal dwProvType As Long, _
                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal Algid As Long, _
                                    ByVal hKey As Long, _
                                    ByVal dwFlags As Long, _
                                    ByRef phHash As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

Private Declare Function CryptHashData _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    pbData As Any, _
                                    ByVal cbData As Long, _
                                    ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    ByVal dwParam As Long, _
                                    pbData As Any, _
                                    ByRef pcbData As Long, _
                                    ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL       As Long = 1

Private Const PROV_RSA_AES        As Long = 24

Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL          As Long = 2

Private Const HP_HASHSIZE         As Long = 4

Private Const ALG_TYPE_ANY        As Long = 0

Private Const ALG_CLASS_HASH      As Long = 32768

Private Const ALG_SID_MD2         As Long = 1

Private Const ALG_SID_MD4         As Long = 2

Private Const ALG_SID_MD5         As Long = 3

Private Const ALG_SID_SHA         As Long = 4

Private Const ALG_SID_SHA_256     As Long = 12

Private Const ALG_SID_SHA_384     As Long = 13

Private Const ALG_SID_SHA_512     As Long = 14

Private Const CALG_MD2            As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)

Private Const CALG_MD4            As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)

Private Const CALG_MD5            As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)

Private Const CALG_SHA            As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)

Private Const CALG_SHA_256        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)

Private Const CALG_SHA_384        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)

Private Const CALG_SHA_512        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

' Create Hash
Private Function CreateHash(abytData() As Byte, _
                            ByVal lngAlgID As Long, _
                            Optional Lower As Boolean = False) As String

    Dim hProv             As Long, hHash As Long

    Dim abytHash(0 To 63) As Byte

    Dim lngLength         As Long

    Dim lngResult         As Long

    Dim strHash           As String

    Dim i                 As Long

    strHash = ""

    If CryptAcquireContext(hProv, vbNullString, vbNullString, IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), CRYPT_VERIFYCONTEXT) <> 0& Then

        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1

            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)

            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1

                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, 0&) <> 0& Then

                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next

                End If
            End If

            CryptDestroyHash hHash
        End If

        CryptReleaseContext hProv, 0&
    End If

    If Lower Then
        CreateHash = LCase$(strHash)
    Else
        CreateHash = strHash
    End If

End Function

' Create Hash From String(Shift_JIS)
Private Function CreateHashString(ByVal strData As String, _
                                  ByVal lngAlgID As Long, _
                                  Optional Lower As Boolean = False) As String
    CreateHashString = CreateHash(StrConv(strData, vbFromUnicode), lngAlgID, Lower)
End Function

' Create Hash From File
Private Function CreateHashFile(ByVal strFileName As String, _
                                ByVal lngAlgID As Long, _
                                Optional Lower As Boolean = False) As String

    Dim abytData() As Byte

    Dim intFile    As Integer

    Dim lngError   As Long

    On Error Resume Next

    If Len(Dir(strFileName)) > 0 Then
        intFile = FreeFile
        Open strFileName For Binary Access Read Shared As #intFile
        abytData() = InputB(LOF(intFile), #intFile)
        Close #intFile
    End If

    lngError = Err.Number

    On Error GoTo 0

    If lngError = 0 Then CreateHashFile = CreateHash(abytData(), lngAlgID, Lower) Else CreateHashFile = ""
End Function

' MD5
'Public Function CreateMD5Hash(abytData() As Byte, Optional Lower As Boolean = False) As String
'    CreateMD5Hash = CreateHash(abytData(), CALG_MD5, Lower)
'End Function
'Public Function CreateMD5HashString(ByVal strData As String, Optional Lower As Boolean = False) As String
'    CreateMD5HashString = CreateHashString(strData, CALG_MD5, Lower)
'End Function
'Public Function CreateMD5HashFile(ByVal strFileName As String, Optional Lower As Boolean = False) As String
'    CreateMD5HashFile = CreateHashFile(strFileName, CALG_MD5, Lower)
'End Function

' SHA-1
Public Function CreateSHA1Hash(abytData() As Byte, _
                               Optional Lower As Boolean = False) As String
    CreateSHA1Hash = CreateHash(abytData(), CALG_SHA, Lower)
End Function

Public Function CreateSHA1HashString(ByVal strData As String, _
                                     Optional Lower As Boolean = False) As String
    CreateSHA1HashString = CreateHashString(strData, CALG_SHA, Lower)
End Function

Public Function CreateSHA1HashFile(ByVal strFileName As String, _
                                   Optional Lower As Boolean = False) As String
    CreateSHA1HashFile = CreateHashFile(strFileName, CALG_SHA, Lower)
End Function

' SHA-256
Public Function CreateSHA256Hash(abytData() As Byte, _
                                 Optional Lower As Boolean = False) As String
    CreateSHA256Hash = CreateHash(abytData(), CALG_SHA_256, Lower)
End Function

Public Function CreateSHA256HashString(ByVal strData As String, _
                                       Optional Lower As Boolean = False) As String
    CreateSHA256HashString = CreateHashString(strData, CALG_SHA_256, Lower)
End Function

Public Function CreateSHA256HashFile(ByVal strFileName As String, _
                                     Optional Lower As Boolean = False) As String
    CreateSHA256HashFile = CreateHashFile(strFileName, CALG_SHA_256, Lower)
End Function

' SHA-384
Public Function CreateSHA384Hash(abytData() As Byte, _
                                 Optional Lower As Boolean = False) As String
    CreateSHA384Hash = CreateHash(abytData(), CALG_SHA_384, Lower)
End Function

Public Function CreateSHA384HashString(ByVal strData As String, _
                                       Optional Lower As Boolean = False) As String
    CreateSHA384HashString = CreateHashString(strData, CALG_SHA_384, Lower)
End Function

Public Function CreateSHA384HashFile(ByVal strFileName As String, _
                                     Optional Lower As Boolean = False) As String
    CreateSHA384HashFile = CreateHashFile(strFileName, CALG_SHA_384, Lower)
End Function

' SHA-512
Public Function CreateSHA512Hash(abytData() As Byte, _
                                 Optional Lower As Boolean = False) As String
    CreateSHA512Hash = CreateHash(abytData(), CALG_SHA_512, Lower)
End Function

Public Function CreateSHA512HashString(ByVal strData As String, _
                                       Optional Lower As Boolean = False) As String
    CreateSHA512HashString = CreateHashString(strData, CALG_SHA_512, Lower)
End Function

Public Function CreateSHA512HashFile(ByVal strFileName As String, _
                                     Optional Lower As Boolean = False) As String
    CreateSHA512HashFile = CreateHashFile(strFileName, CALG_SHA_512, Lower)
End Function

