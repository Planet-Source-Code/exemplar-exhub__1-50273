Attribute VB_Name = "modDCFunctions"
Option Explicit

Function CreatePK()
    Dim i As Integer
    Dim strPK As String
    
    strPK = "eXHub"
    
    For i = Len(strPK) To 15
        strPK = strPK + Chr(Int(Rnd * 26) + 65)
    Next
    
    CreatePK = strPK
End Function

Function CreateLock() As String
    Dim i As Integer
    Dim strLock As String
    
    strLock = "eXHub"
    
    For i = 1 To 80
        strLock = strLock + Chr(Int(Rnd * 26) + 65)
    Next
    
    CreateLock = strLock
End Function

Function Lock2Key(strLock As String) As String
    Dim TLock2Key As String, TChar As Integer, i As Integer
    If Len(strLock) < 3 Then Lock2Key = Left$("BROKENCLIENT", Len(strLock))
    TLock2Key = Chr$(Asc(Left$(strLock, 1)) Xor Asc(Right$(strLock, 1)) Xor Asc(Mid$(strLock, Len(strLock) - 1, 1)) Xor 5)
    For i = 2 To Len(strLock)
        TLock2Key = TLock2Key + Chr$(Asc(Mid$(strLock, i, 1)) Xor Asc(Mid$(strLock, i - 1, 1)))
    Next i
    For i = 1 To Len(TLock2Key)
        TChar = Asc(Mid$(TLock2Key, i, 1))
        TChar = (TChar + ((TChar Mod 17) * 15))
        While TChar > 255: TChar = TChar - 255: Wend
        Select Case TChar
        Case 0, 5, 35, 96, 214, 126
            Lock2Key = Lock2Key + "/%DCN" + Format(TChar, "000") + "%/"
        Case Else
            Lock2Key = Lock2Key + Chr$(TChar)
        End Select
    Next i
End Function

Public Function DC1_Lock2Key(Lck$) As String
    Dim pos%, K$, i%, T%
    'Remade optimized and works with clients
    pos = InStr(Lck, " ")
    If pos Then 'If not, assume it was pre-parsed
        Lck = Left(Lck, pos - 1)
    End If
    If Len(Lck) < 3 Then    'Too short!
        DC1_Lock2Key = "Invalid Lock < 3 chars"
        Exit Function
    End If
    K = ""
    For i = 1 To Len(Lck)
        T = Asc(Mid(Lck, i))
        If i = 1 Then
            T = T Xor 5
        Else
            T = T Xor Asc(Mid(Lck, i - 1))
        End If
        T = (T + ((T Mod 17) * 15))
        'Can't just Mod 255 cuz if # /IS/ 255, we don't change
        Do Until T <= 255
            T = T - 255
        Loop
        Select Case T
        Case 0, 5, 96, 124, 126, 36
            K = K + "/%DCN" + Right("00" + CStr(T), 3) + "%/"
        Case Else
            K = K + Chr(T)
        End Select
    Next
    Mid(K, 1, 1) = Chr(Asc(K) Xor Asc(Mid(K, Len(K))))
    DC1_Lock2Key = K
End Function
