Attribute VB_Name = "mMD5"
'计算文件    MD5  str1 = MD5FormFile("c:\123.exe",[True=32位,False=16位])
'计算字符串  MD5  str1 = MD5FormString("65464145",[True=32位,False=16位])
'计算 Byte() MD5  str1 = MD5FormByte(p() as byte,[True=32位,False=16位])

Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA1 = 4

Public Enum HashAlgorithm
    MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    md5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum

Private Const HP_HASHVAL = 2
Private Const HP_HASHSIZE = 4
Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private m_lOnBits(0 To 30)    As Long
Private m_l2Power(0 To 30)    As Long

'计算 文件 MD5
Public Function MD5FormFile(ByVal FilePath As String, Optional ByVal bOut32Bit As Boolean = True, Optional ByVal Algorithm As HashAlgorithm = md5, Optional ByVal UpperCase As Boolean = True) As String
    Dim hCtx       As Long
    Dim hHash      As Long
    Dim lFile      As Long
    Dim lRes       As Long
    Dim lLen       As Long
    Dim lIdx       As Long
    Dim lCount     As Long
    Dim lBlocks    As Long
    Dim lLastBlock As Long
    Dim abHash()   As Byte
    Dim sResult    As String

    If Len(Dir(FilePath)) = 0 Then GoTo ToExit

    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)

    If lRes = 0 And Err.LastDllError = &H80090016 Then
        lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    End If

    If lRes <> 0 Then
        lRes = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)

        If lRes <> 0 Then
            lFile = FreeFile

            Open FilePath For Binary As lFile

            If Err.Number = 0 Then

                Const BLOCK_SIZE As Long = 32 * 1024& ' 32K

                ReDim abBlock(1 To BLOCK_SIZE) As Byte

                lBlocks = LOF(lFile) \ BLOCK_SIZE
                lLastBlock = LOF(lFile) - lBlocks * BLOCK_SIZE

                For lCount = 1 To lBlocks
                    Get lFile, , abBlock

                    lRes = CryptHashData(hHash, abBlock(1), BLOCK_SIZE, 0)

                    If lRes = 0 Then Exit For
                Next

                If lLastBlock > 0 And lRes <> 0 Then
                    ReDim abBlock(1 To lLastBlock) As Byte
                    Get lFile, , abBlock

                    lRes = CryptHashData(hHash, abBlock(1), lLastBlock, 0)
                End If

                Close lFile
            End If

            If lRes <> 0 Then
                lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)

                If lRes <> 0 Then
                    ReDim abHash(0 To lLen - 1)
                    lRes = CryptGetHashParam(hHash, HP_HASHVAL, abHash(0), lLen, 0)

                    If lRes <> 0 Then

                        For lIdx = 0 To UBound(abHash)
                            sResult = sResult & Right$("0" & Hex$(abHash(lIdx)), 2)
                        Next

                    End If
                End If
            End If

            CryptDestroyHash hHash
        End If
    End If

    CryptReleaseContext hCtx, 0
    
    If lRes = 0 Then
        MsgBox Err.Description, vbInformation, "校验文件MD5时出错："
    End If
    
    If bOut32Bit = False And Len(sResult) = 32 Then
        sResult = Mid$(sResult, 9, 16)
    End If
    
    If UpperCase Then
        MD5FormFile = sResult
    Else
        MD5FormFile = LCase$(sResult)
    End If

ToExit:
End Function

'计算 字符串 MD5
Public Function MD5FormString(ByVal sText As String, Optional ByVal bOut32Bit As Boolean = True) As String
    MD5FormString = MD5FormByte(StrConv(sText, vbFromUnicode), bOut32Bit)
End Function

'计算 byte() MD5
Public Function MD5FormByte(ByRef bytMessage() As Byte, Optional ByVal bOut32Bit As Boolean = True) As String
    Dim X()   As Long
    Dim k     As Long
    Dim aa    As Long
    Dim bb    As Long
    Dim cc    As Long
    Dim dd    As Long
    Dim a     As Long
    Dim b     As Long
    Dim c     As Long
    Dim d     As Long
    
    Const S11 As Long = 7
    Const S12 As Long = 12
    Const S13 As Long = 17
    Const S14 As Long = 22
    Const S21 As Long = 5
    Const S22 As Long = 9
    Const S23 As Long = 14
    Const S24 As Long = 20
    Const S31 As Long = 4
    Const S32 As Long = 11
    Const S33 As Long = 16
    Const S34 As Long = 23
    Const S41 As Long = 6
    Const S42 As Long = 10
    Const S43 As Long = 15
    Const S44 As Long = 21
   
    Call InitMD5 '初始化 MD5
    X = ConvertByteArrayToWordArray(bytMessage)
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(X) Step 16
        aa = a
        bb = b
        cc = c
        dd = d

        FF a, b, c, d, X(k + 0), S11, &HD76AA478
        FF d, a, b, c, X(k + 1), S12, &HE8C7B756
        FF c, d, a, b, X(k + 2), S13, &H242070DB
        FF b, c, d, a, X(k + 3), S14, &HC1BDCEEE
        FF a, b, c, d, X(k + 4), S11, &HF57C0FAF
        FF d, a, b, c, X(k + 5), S12, &H4787C62A
        FF c, d, a, b, X(k + 6), S13, &HA8304613
        FF b, c, d, a, X(k + 7), S14, &HFD469501
        FF a, b, c, d, X(k + 8), S11, &H698098D8
        FF d, a, b, c, X(k + 9), S12, &H8B44F7AF
        FF c, d, a, b, X(k + 10), S13, &HFFFF5BB1
        FF b, c, d, a, X(k + 11), S14, &H895CD7BE
        FF a, b, c, d, X(k + 12), S11, &H6B901122
        FF d, a, b, c, X(k + 13), S12, &HFD987193
        FF c, d, a, b, X(k + 14), S13, &HA679438E
        FF b, c, d, a, X(k + 15), S14, &H49B40821
        GG a, b, c, d, X(k + 1), S21, &HF61E2562
        GG d, a, b, c, X(k + 6), S22, &HC040B340
        GG c, d, a, b, X(k + 11), S23, &H265E5A51
        GG b, c, d, a, X(k + 0), S24, &HE9B6C7AA
        GG a, b, c, d, X(k + 5), S21, &HD62F105D
        GG d, a, b, c, X(k + 10), S22, &H2441453
        GG c, d, a, b, X(k + 15), S23, &HD8A1E681
        GG b, c, d, a, X(k + 4), S24, &HE7D3FBC8
        GG a, b, c, d, X(k + 9), S21, &H21E1CDE6
        GG d, a, b, c, X(k + 14), S22, &HC33707D6
        GG c, d, a, b, X(k + 3), S23, &HF4D50D87
        GG b, c, d, a, X(k + 8), S24, &H455A14ED
        GG a, b, c, d, X(k + 13), S21, &HA9E3E905
        GG d, a, b, c, X(k + 2), S22, &HFCEFA3F8
        GG c, d, a, b, X(k + 7), S23, &H676F02D9
        GG b, c, d, a, X(k + 12), S24, &H8D2A4C8A
        HH a, b, c, d, X(k + 5), S31, &HFFFA3942
        HH d, a, b, c, X(k + 8), S32, &H8771F681
        HH c, d, a, b, X(k + 11), S33, &H6D9D6122
        HH b, c, d, a, X(k + 14), S34, &HFDE5380C
        HH a, b, c, d, X(k + 1), S31, &HA4BEEA44
        HH d, a, b, c, X(k + 4), S32, &H4BDECFA9
        HH c, d, a, b, X(k + 7), S33, &HF6BB4B60
        HH b, c, d, a, X(k + 10), S34, &HBEBFBC70
        HH a, b, c, d, X(k + 13), S31, &H289B7EC6
        HH d, a, b, c, X(k + 0), S32, &HEAA127FA
        HH c, d, a, b, X(k + 3), S33, &HD4EF3085
        HH b, c, d, a, X(k + 6), S34, &H4881D05
        HH a, b, c, d, X(k + 9), S31, &HD9D4D039
        HH d, a, b, c, X(k + 12), S32, &HE6DB99E5
        HH c, d, a, b, X(k + 15), S33, &H1FA27CF8
        HH b, c, d, a, X(k + 2), S34, &HC4AC5665
        II a, b, c, d, X(k + 0), S41, &HF4292244
        II d, a, b, c, X(k + 7), S42, &H432AFF97
        II c, d, a, b, X(k + 14), S43, &HAB9423A7
        II b, c, d, a, X(k + 5), S44, &HFC93A039
        II a, b, c, d, X(k + 12), S41, &H655B59C3
        II d, a, b, c, X(k + 3), S42, &H8F0CCC92
        II c, d, a, b, X(k + 10), S43, &HFFEFF47D
        II b, c, d, a, X(k + 1), S44, &H85845DD1
        II a, b, c, d, X(k + 8), S41, &H6FA87E4F
        II d, a, b, c, X(k + 15), S42, &HFE2CE6E0
        II c, d, a, b, X(k + 6), S43, &HA3014314
        II b, c, d, a, X(k + 13), S44, &H4E0811A1
        II a, b, c, d, X(k + 4), S41, &HF7537E82
        II d, a, b, c, X(k + 11), S42, &HBD3AF235
        II c, d, a, b, X(k + 2), S43, &H2AD7D2BB
        II b, c, d, a, X(k + 9), S44, &HEB86D391

        a = AddUnsigned(a, aa)
        b = AddUnsigned(b, bb)
        c = AddUnsigned(c, cc)
        d = AddUnsigned(d, dd)
    Next

    MD5FormByte = LCase$(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))

    If bOut32Bit = False Then MD5FormByte = Mid$(MD5FormByte, 9, 16)
End Function

'初始化 MD5
Private Sub InitMD5()
    m_lOnBits(0) = 1           ' 00000000000000000000000000000001
    m_lOnBits(1) = 3           ' 00000000000000000000000000000011
    m_lOnBits(2) = 7           ' 00000000000000000000000000000111
    m_lOnBits(3) = 15          ' 00000000000000000000000000001111
    m_lOnBits(4) = 31          ' 00000000000000000000000000011111
    m_lOnBits(5) = 63          ' 00000000000000000000000000111111
    m_lOnBits(6) = 127         ' 00000000000000000000000001111111
    m_lOnBits(7) = 255         ' 00000000000000000000000011111111
    m_lOnBits(8) = 511         ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023        ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047       ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095       ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191       ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383      ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767      ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535      ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071     ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143     ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287     ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575    ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151    ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303    ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607    ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215   ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431   ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863   ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727  ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455  ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911  ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823 ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647 ' 01111111111111111111111111111111

    m_l2Power(0) = 1           ' 00000000000000000000000000000001
    m_l2Power(1) = 2           ' 00000000000000000000000000000010
    m_l2Power(2) = 4           ' 00000000000000000000000000000100
    m_l2Power(3) = 8           ' 00000000000000000000000000001000
    m_l2Power(4) = 16          ' 00000000000000000000000000010000
    m_l2Power(5) = 32          ' 00000000000000000000000000100000
    m_l2Power(6) = 64          ' 00000000000000000000000001000000
    m_l2Power(7) = 128         ' 00000000000000000000000010000000
    m_l2Power(8) = 256         ' 00000000000000000000000100000000
    m_l2Power(9) = 512         ' 00000000000000000000001000000000
    m_l2Power(10) = 1024       ' 00000000000000000000010000000000
    m_l2Power(11) = 2048       ' 00000000000000000000100000000000
    m_l2Power(12) = 4096       ' 00000000000000000001000000000000
    m_l2Power(13) = 8192       ' 00000000000000000010000000000000
    m_l2Power(14) = 16384      ' 00000000000000000100000000000000
    m_l2Power(15) = 32768      ' 00000000000000001000000000000000
    m_l2Power(16) = 65536      ' 00000000000000010000000000000000
    m_l2Power(17) = 131072     ' 00000000000000100000000000000000
    m_l2Power(18) = 262144     ' 00000000000001000000000000000000
    m_l2Power(19) = 524288     ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576    ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152    ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304    ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608    ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216   ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432   ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864   ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728  ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456  ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912  ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824 ' 01000000000000000000000000000000
End Sub

Private Function ConvertByteArrayToWordArray(bytMessage() As Byte) As Long()
    Dim lMessageLength   As Long
    Dim lNumberOfWords   As Long
    Dim lWordArray()     As Long
    Dim lBytePosition    As Long
    Dim lByteCount       As Long
    Dim lWordCount       As Long

    Const MODULUS_BITS   As Long = 512
    Const CONGRUENT_BITS As Long = 448

    lMessageLength = UBound(bytMessage) + 1
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)

    ReDim lWordArray(lNumberOfWords - 1)
    lBytePosition = 0
    lByteCount = 0

    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(bytMessage(lByteCount), lBytePosition)
        
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertByteArrayToWordArray = lWordArray
End Function

Private Function WordToHex(ByVal lValue As Long) As String
    Dim lByte  As Long
    Dim lCount As Long
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right$("0" & Hex$(lByte), 2)
    Next
End Function

Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If

        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue

        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If

        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End Function

Private Function RShiftSigned(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShiftSigned = lValue

        Exit Function
    ElseIf iShiftBits = 31 Then
        If (lValue And &H80000000) Then
            RShiftSigned = -1
        Else
            RShiftSigned = 0
        End If

        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    RShiftSigned = Int(lValue / m_l2Power(iShiftBits))
End Function

Private Function RotateLeft(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

Private Function f(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    f = (X And Y) Or ((Not X) And z)
End Function

Private Function G(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    G = (X And z) Or (Y And (Not z))
End Function

Private Function h(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    h = (X Xor Y Xor z)
End Function

Private Function i(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    i = (Y Xor (X Or (Not z)))
End Function

Private Sub FF(a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(f(b, c, d), X), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub GG(a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), X), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub HH(a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(h(b, c, d), X), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub II(a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal X As Long, ByVal s As Long, ByVal ac As Long)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(i(b, c, d), X), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub
