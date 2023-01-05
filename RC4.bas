Sub RC4_Crypt_test()

    Debug.Print RC4_DeCrypt(RC4_Crypt("This is a Test 4 Jesus Christ!!!", "WNakrMGtEduppvhffyxCZp", True, 99), "WNakrMGtEduppvhffyxCZp", True, 99)

End Sub

Public Function RC4_Crypt(sPlainText As String, sKey As String, Optional blSpritz As Boolean = False, Optional bytW As Byte = 1) As String
    
    Dim bytPlainArray() As Byte
    Dim bytKeyArray() As Byte
    Dim bytCipherArray() As Byte
    Dim bytKeyStreamArray() As Byte
    Dim strFName4Hex As String
    
    Text2BytArray sKey, bytArray:=bytKeyArray(), blFileMode:=False
    Text2BytArray sPlainText, bytArray:=bytPlainArray(), blFileMode:=False
    Crypt bytPlainArray:=bytPlainArray(), bytKeyArray:=bytKeyArray(), bytCipherArray:=bytCipherArray(), bytKeyStreamArray:=bytKeyStreamArray(), blSpritz:=blSpritz, bytW:=bytW
    BytArray2Hex strFNameOrText:=strFName4Hex, bytArray:=bytCipherArray(), blFileMode:=False
    RC4_Crypt = strFName4Hex

End Function

Public Function RC4_DeCrypt(sHex As String, sKey As String, Optional blSpritz As Boolean = False, Optional bytW As Byte = 1) As String
    
    Dim bytPlainArray() As Byte
    Dim bytKeyArray() As Byte
    Dim bytCipherArray() As Byte
    Dim bytKeyStreamArray() As Byte
    Dim strFName4Plain As String
    
    Text2BytArray sKey, bytArray:=bytKeyArray(), blFileMode:=False
    Hex2BytArray strFNameOrText:=sHex, bytArray:=bytCipherArray(), blFileMode:=False
    Crypt bytPlainArray:=bytCipherArray(), bytKeyArray:=bytKeyArray(), bytCipherArray:=bytPlainArray(), bytKeyStreamArray:=bytKeyStreamArray(), blSpritz:=blSpritz, bytW:=bytW
    BytArray2Text strFNameOrText:=strFName4Plain, bytArray:=bytPlainArray(), blFileMode:=False
    RC4_DeCrypt = strFName4Plain
    
End Function

Private Sub Text2BytArray(ByVal strFNameOrText As String, bytArray() As Byte, Optional blFileMode As Boolean = True)
    
    Dim i As Long
    
    If blFileMode Then
        Dim intFNo As Integer
        intFNo = FreeFile
        
        Open strFNameOrText For Binary As #intFNo
        ReDim bytArray(LOF(intFNo) - 1)
        Get #intFNo, , bytArray()
        Close #intFNo
    Else
        bytArray = StrConv(strFNameOrText, vbFromUnicode)
    End If
    
End Sub

Private Sub BytArray2Text(strFNameOrText As String, bytArray() As Byte, Optional blFileMode As Boolean = True)
    
    If blFileMode Then
        Dim intFNo As Integer
        intFNo = FreeFile
        
        Open strFNameOrText For Output As #intFNo
        Close #intFNo
        
        intFNo = FreeFile
        Open strFNameOrText For Binary As #intFNo
        Put #intFNo, , bytArray()
        Close #intFNo
    Else
        strFNameOrText = StrConv(bytArray(), vbUnicode)
    End If
    
End Sub

Private Sub Hex2BytArray(ByVal strFNameOrText As String, bytArray() As Byte, Optional blFileMode As Boolean = True)
    
    Dim i As Long
    Dim j As Integer
    Dim strHex As String
    Dim lonLen As Long
    Dim strHex4Byte As String
    
    If blFileMode Then
        Dim intFNo As Integer
        
        intFNo = FreeFile
        Open strFNameOrText For Input As #intFNo
        Line Input #intFNo, strHex
        Close #intFNo
    Else
        strHex = strFNameOrText
    End If
    
    lonLen = Len(strHex)
    j = 0
    For i = 1 To lonLen - 1 Step 2
        strHex4Byte = Mid(strHex, i, 2)
        ReDim Preserve bytArray(j)
        bytArray(j) = CByte("&H" & strHex4Byte)
        j = j + 1
    Next i
    
End Sub

Private Sub BytArray2Hex(strFNameOrText As String, bytArray() As Byte, Optional blFileMode As Boolean = True)
    
    Dim strHex As String
    Dim i As Integer
    
    For i = LBound(bytArray()) To UBound(bytArray) Step 1
        strHex = strHex & Right("0" & Hex(bytArray(i)), 2)
    Next i
    
    If blFileMode Then
        Dim intFNo As Integer
    
        intFNo = FreeFile
        Open strFNameOrText For Output As #intFNo
        Print #intFNo, strHex
        Close #intFNo
    Else
        strFNameOrText = strHex
    End If
    
End Sub

Private Sub Crypt(bytPlainArray() As Byte, bytKeyArray() As Byte, bytCipherArray() As Byte, bytKeyStreamArray() As Byte, Optional blSpritz As Boolean = False, Optional bytW As Byte = 1)
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim w As Byte
    Dim S(255) As Integer
    Dim intKeyLen As Integer
    Dim intTemp As Integer
    Dim m As Integer
    Dim intPlainLen As Integer
    Dim z As Integer
    
    intKeyLen = UBound(bytKeyArray) - LBound(bytKeyArray) + 1
    intPlainLen = UBound(bytPlainArray) - LBound(bytPlainArray) + 1
    ReDim bytCipherArray(intPlainLen - 1)
    ReDim bytKeyStreamArray(intPlainLen - 1)
    
    For i = 0 To 255 Step 1
        S(i) = i
    Next i
    
    j = 0
    For i = 0 To 255 Step 1
        j = (j + S(i) + bytKeyArray(i Mod intKeyLen)) Mod 256
        intTemp = S(i)
        S(i) = S(j)
        S(j) = intTemp
    Next i
    
    i = 0
    j = 0
    z = 0
    For m = 0 To intPlainLen - 1 Step 1
        If blSpritz Then
            
            ' see: https://www.schneier.com/blog/archives/2014/10/spritz_a_new_rc.html
            ' i = i + w
            ' j = k + S[j + S[i]]
            ' k = i + k + S[j]
            ' SWAP(S[i];S[j])
            ' z = S[j + S[i + S[z + k]]]
            ' Return z
            
            k = 0
            w = bytW
            i = (i + w) Mod 256
            j = (k + S((j + S(i)) Mod 256)) Mod 256
            k = (i + k + S(j)) Mod 256
            Call intArraySwitchElement(S, i, j) ' switch element
            z = S((j + S((i + S((z + k) Mod 256)) Mod 256)) Mod 256)
            
        Else
            
            ' see: https://www.schneier.com/blog/archives/2014/10/spritz_a_new_rc.html
            ' i = i + 1
            ' j = j + S[i]
            ' SWAP(S[i];S[j])
            ' z = S[S[i] + S[j]]
            ' Return z
        
            i = (i + 1) Mod 256
            j = (j + S(i)) Mod 256
            Call intArraySwitchElement(S, i, j) ' switch element
            z = S((S(i) + S(j)) Mod 256)
            
        End If
        
        bytKeyStreamArray(m) = z
        bytCipherArray(m) = CByte(bytPlainArray(m) Xor z)
    Next m
End Sub

Private Sub intArraySwitchElement(intArray() As Integer, i As Integer, j As Integer)
    Dim intTemp As Integer
    
    intTemp = intArray(i) ' save
    intArray(i) = intArray(j) ' switch
    intArray(j) = intTemp ' resave

End Sub

Private Sub bytArrayDebugPrint(bytArray() As Byte)
    Dim i As Long

    For i = LBound(bytArray) To UBound(bytArray)
        Debug.Print i, bytArray(i), Chr$(bytArray(i))
    Next i

End Sub
