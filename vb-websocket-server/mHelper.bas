Attribute VB_Name = "mHelper"
Option Explicit
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private cFreq As Currency

Function LikeB(bIn() As Byte, bPattern() As Byte, Optional iStart As Long, _
    Optional iLength As Long) As Boolean
'Characters in pattern   Matches in string
'?   Any single character
'*   Zero or more characters
'#   Any single digit (0–9)
'[charlist]  Any single character in charlist
'[!charlist] Any single character not in charlist
Dim i As Long, j As Long, ubIn As Long, ubPat As Long, chrIn As Long, chrP As Long
Dim b As Boolean, lastP As Long, lasti As Long, patFound As Boolean, charListOpen As Boolean
Dim bNB As Boolean, bNA As Boolean, bCL As Boolean, bcLSep As Boolean, charList(1) As Long
Dim PatternList() As Long, PatternCount As Long, patType As Long, patContinue As Boolean
    ubIn = UBound(bIn)
        If iStart > 0 And iStart < ubIn + 1 Then
            iStart = iStart - 1
        Else
            iStart = 0
        End If
        If iLength > 0 And iLength <= ubIn + 1 Then
            ubIn = iLength - 1
        End If
    ubPat = UBound(bPattern)
        LikeB = True
If ubIn = -1 Or ubPat = -1 Then Exit Function
    
        charList(0) = -1
        charList(1) = -1
        charListOpen = False
        patType = -1
        patContinue = False
            ReDim PatternList(0 To ubPat, 0 To 5)
    For i = 0 To ubPat
            chrP = bPattern(i)
        Select Case chrP
            Case 63 '?   Any single character
                patType = 63
            Case 42 '*   Zero or more characters
               patType = 42
            Case 35 '#   Any single digit (0–9)
                patType = 35
            Case 33 '!
                If bCL = True Then
                    bNA = True 'After[!]
                Else
                    bNB = True 'Before ![]
                End If
                charListOpen = True
                patType = -1
            Case 91 '[
                bCL = True
                charListOpen = True
                patType = -1
            Case 93 ']
                bCL = False
                charListOpen = False
                patType = -1
            Case 45 '-
                bcLSep = True
                patType = -1
            Case Else
                    patType = chrP
                If bCL = True And charList(0) = -1 Then
                    charList(0) = chrP
                        patType = -1
                End If
                If bcLSep = True And charList(1) = -1 Then
                    charList(1) = chrP
                        patType = -1
                End If
        End Select
            If charList(1) <> -1 Then
                If lastP = 91 Then
                    patContinue = True
                End If
                patType = 91
                lastP = patType
                bcLSep = False
            End If
            If bCL = False And charListOpen = False Then
                        patType = chrP
                    If bNB = True Then
                    End If
                lastP = -1
                patContinue = False
            End If
        If patType <> -1 Then
            j = PatternCount
            PatternList(j, 0) = patType
            PatternList(j, 1) = bNB
            PatternList(j, 2) = bNA
            PatternList(j, 3) = patContinue
            PatternList(j, 4) = charList(0)
            PatternList(j, 5) = charList(1)
                PatternCount = PatternCount + 1
                charList(0) = -1
                charList(1) = -1
                charListOpen = False
            If patType = 93 Then
                bNB = False
            End If
                bNA = False
                patContinue = False
            patType = -1
        End If
    Next
     
        b = False
        lastP = -1
        charList(0) = -1
        charList(1) = -1
        charListOpen = False
            j = 0
    For i = iStart To ubIn
            chrIn = bIn(i)
        If j > PatternCount - 1 Then
            Exit For
        End If
            patType = PatternList(j, 0)
        Select Case patType
            Case 63 '?   Any single character
                b = True
            Case 42 '*   Zero or more characters
                b = True
                lastP = patType
            Case 35 '#   Any single digit (0–9)
                If chrIn > 47 And chrIn < 58 Then
                    b = True
                Else
                    b = False
                    Exit For
                End If
            Case 91 '[
                If PatternList(j, 3) = -1 Then 'continue
                    b = b Or (chrIn >= PatternList(j, 4) And chrIn <= PatternList(j, 5))
                Else
                    b = chrIn >= PatternList(j, 4) And chrIn <= PatternList(j, 5)
                End If
                If PatternList(j, 2) = -1 Then
                    b = Not b
                End If
            Case 93 ']
                If PatternList(j, 1) = -1 Then
                    b = Not b
                End If
            Case Else
                b = chrIn = patType
        End Select
            j = j + 1
        If b = False And lastP = 42 And i <> ubIn Then
            b = True
            j = j - 1
        End If
        If patType = 91 Then 'check 93
            i = i - 1
        End If
        If b = False Then
            Exit For
        End If
    Next
        If j < PatternCount - 1 Then
            b = False
        End If
Erase PatternList
LikeB = b
End Function
Sub TestLikeB()
Dim testCheck As Boolean
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("A", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("a", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("AAA", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("aBBBa", vbFromUnicode), StrConv("a*a", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("aBBBa", vbFromUnicode), StrConv("a*b", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("[A-Z]", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("![A-Z]", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("[A-CX-Z]", vbFromUnicode)) 'true
testCheck = LikeB(StrConv("A", vbFromUnicode), StrConv("![A-CX-Z]", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("a2a", vbFromUnicode), StrConv("a#a", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("aM5b", vbFromUnicode), StrConv("a[L-P]#[!c-e]", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("BAT123khg", vbFromUnicode), StrConv("B?T*", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("CAT123khg", vbFromUnicode), StrConv("B?T*", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("aM5b", vbFromUnicode), StrConv("a![L-P]#[!c-e]", vbFromUnicode)) 'False
testCheck = LikeB(StrConv("GET /public/index.php/login HTTP/1.1", vbFromUnicode), StrConv("GET*HTTP/1.?*", vbFromUnicode)) 'True
testCheck = LikeB(StrConv("POST /public/index.php/login HTTP/1.1", vbFromUnicode), StrConv("GET*HTTP/1.?*", vbFromUnicode)) 'True

End Sub
Sub TestLikeBPerf()
    Const iterations    As Long = 10000 '0 ' 400000 '
    Dim i           As Long
    Dim cValue As Currency
    Dim dStartTime  As Double
    Dim dStopTime   As Double
    Dim str As String, strM As String, b As Boolean
    Dim bIn() As Byte, bM() As Byte
            QueryPerformanceFrequency cFreq

        str = "GET /public/index.php/login HTTP/1.1"
        strM = "GET*HTTP/1.?*"
        bIn = StrConv(str, vbUnicode)
        bM = StrConv(str, vbUnicode)
    
    QueryPerformanceCounter cValue
    dStartTime = cValue / cFreq  ' The division will return a double, and it also cancels out the Currency decimal points.
    For i = 1& To iterations
        b = str Like strM
    Next
            QueryPerformanceCounter cValue
    dStopTime = cValue / cFreq  ' The division will return a double, and it also cancels out the Currency decimal points.
    Debug.Print "Like Time:" & (dStopTime - dStartTime)
    
    QueryPerformanceCounter cValue
    dStartTime = cValue / cFreq  ' The division will return a double, and it also cancels out the Currency decimal points.
    For i = 1& To iterations
        b = LikeB(bIn, bM)
    Next
            QueryPerformanceCounter cValue
    dStopTime = cValue / cFreq  ' The division will return a double, and it also cancels out the Currency decimal points.
    Debug.Print "LikeB Time:" & (dStopTime - dStartTime)
'Compile to p-code seems likeb faster
End Sub
