Attribute VB_Name = "modSHA1"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' TITLE:
' Secure Hash Algorithm, SHA-1

' AUTHORS:
' Adapted by Iain Buchan from Visual Basic code posted at Planet-Source-Code by Peter Girard
' http://www.planetsourcecode.com/xq/ASP/txtCodeId.13565/lngWId.1/qx/vb/scripts/ShowCode.htm

' PURPOSE:
' Creating a secure identifier from person-identifiable data

' The function SecureHash generates a 160-bit (20-hex-digit) message digest for a given message (String).
' It is computationally infeasable to recover the message from the digest.
' The digest is unique to the message within the realms of practical probability.
' The only way to find the source message for a digest is by hashing all possible messages and comparison of their digests.

' REFERENCES:
' For a fuller description see FIPS Publication 180-1:
' http://www.itl.nist.gov/fipspubs/fip180-1.htm

' SAMPLE:
' Message: "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
' Returns Digest: "84983E441C3BD26EBAAE4AA1F95129E5E54670F1"
' Message: "abc"
' Returns Digest: "A9993E364706816ABA3E25717850C26C9CD0D89D"

Private Type Word
B0 As Byte
B1 As Byte
B2 As Byte
B3 As Byte
End Type

'Public Function idcode(cr As Range) As String
' Dim tx As String
' Dim ob As Object
' For Each ob In cr
' tx = tx & LCase(CStr(ob.Value2))
' Next
' idcode = sha1(tx)
'End Function

Private Function AndW(w1 As Word, w2 As Word) As Word
AndW.B0 = w1.B0 And w2.B0
AndW.B1 = w1.B1 And w2.B1
AndW.B2 = w1.B2 And w2.B2
AndW.B3 = w1.B3 And w2.B3
End Function

Private Function OrW(w1 As Word, w2 As Word) As Word
OrW.B0 = w1.B0 Or w2.B0
OrW.B1 = w1.B1 Or w2.B1
OrW.B2 = w1.B2 Or w2.B2
OrW.B3 = w1.B3 Or w2.B3
End Function

Private Function XorW(w1 As Word, w2 As Word) As Word
XorW.B0 = w1.B0 Xor w2.B0
XorW.B1 = w1.B1 Xor w2.B1
XorW.B2 = w1.B2 Xor w2.B2
XorW.B3 = w1.B3 Xor w2.B3
End Function

Private Function NotW(W As Word) As Word
NotW.B0 = Not W.B0
NotW.B1 = Not W.B1
NotW.B2 = Not W.B2
NotW.B3 = Not W.B3
End Function

Private Function AddW(w1 As Word, w2 As Word) As Word
Dim i As Long, W As Word

i = CLng(w1.B3) + w2.B3
W.B3 = i Mod 256
i = CLng(w1.B2) + w2.B2 + (i \ 256)
W.B2 = i Mod 256
i = CLng(w1.B1) + w2.B1 + (i \ 256)
W.B1 = i Mod 256
i = CLng(w1.B0) + w2.B0 + (i \ 256)
W.B0 = i Mod 256

AddW = W
End Function

Private Function CircShiftLeftW(W As Word, n As Long) As Word
Dim d1 As Double, d2 As Double

d1 = WordToDouble(W)
d2 = d1
d1 = d1 * (2 ^ n)
d2 = d2 / (2 ^ (32 - n))
CircShiftLeftW = OrW(DoubleToWord(d1), DoubleToWord(d2))
End Function

Private Function WordToHex(W As Word) As String
WordToHex = Right$("0" & Hex$(W.B0), 2) & Right$("0" & Hex$(W.B1), 2) _
& Right$("0" & Hex$(W.B2), 2) & Right$("0" & Hex$(W.B3), 2)
End Function

Private Function HexToWord(H As String) As Word
HexToWord = DoubleToWord(val("&H" & H & "#"))
End Function

Private Function DoubleToWord(n As Double) As Word
DoubleToWord.B0 = Int(DMod(n, 2 ^ 32) / (2 ^ 24))
DoubleToWord.B1 = Int(DMod(n, 2 ^ 24) / (2 ^ 16))
DoubleToWord.B2 = Int(DMod(n, 2 ^ 16) / (2 ^ 8))
DoubleToWord.B3 = Int(DMod(n, 2 ^ 8))
End Function

Private Function WordToDouble(W As Word) As Double
WordToDouble = (W.B0 * (2 ^ 24)) + (W.B1 * (2 ^ 16)) + (W.B2 * (2 ^ 8)) _
+ W.B3
End Function

Private Function DMod(Value As Double, divisor As Double) As Double
DMod = Value - (Int(Value / divisor) * divisor)
If DMod < 0 Then DMod = DMod + divisor
End Function

Private Function F(t As Long, b As Word, c As Word, D As Word) As Word
Select Case t
Case Is <= 19
F = OrW(AndW(b, c), AndW(NotW(b), D))
Case Is <= 39
F = XorW(XorW(b, c), D)
Case Is <= 59
F = OrW(OrW(AndW(b, c), AndW(b, D)), AndW(c, D))
Case Else
F = XorW(XorW(b, c), D)
End Select
End Function
Public Function StringSHA1(inMessage As String) As String
' ??????SHA1??
Dim inLen As Long
Dim inLenW As Word
Dim padMessage As String
Dim numBlocks As Long
Dim W(0 To 79) As Word
Dim blockText As String
Dim wordText As String
Dim i As Long, t As Long
Dim temp As Word
Dim k(0 To 3) As Word
Dim H0 As Word
Dim H1 As Word
Dim H2 As Word
Dim H3 As Word
Dim H4 As Word
Dim A As Word
Dim b As Word
Dim c As Word
Dim D As Word
Dim E As Word

inMessage = StrConv(inMessage, vbFromUnicode)

inLen = LenB(inMessage)
inLenW = DoubleToWord(CDbl(inLen) * 8)

padMessage = inMessage & ChrB(128) _
& StrConv(String((128 - (inLen Mod 64) - 9) Mod 64 + 4, Chr(0)), 128) _
& ChrB(inLenW.B0) & ChrB(inLenW.B1) & ChrB(inLenW.B2) & ChrB(inLenW.B3)

numBlocks = LenB(padMessage) / 64

' initialize constants
k(0) = HexToWord("5A827999")
k(1) = HexToWord("6ED9EBA1")
k(2) = HexToWord("8F1BBCDC")
k(3) = HexToWord("CA62C1D6")

' initialize 160-bit (5 words) buffer
H0 = HexToWord("67452301")
H1 = HexToWord("EFCDAB89")
H2 = HexToWord("98BADCFE")
H3 = HexToWord("10325476")
H4 = HexToWord("C3D2E1F0")

' each 512 byte message block consists of 16 words (W) but W is expanded
For i = 0 To numBlocks - 1
blockText = MidB$(padMessage, (i * 64) + 1, 64)
' initialize a message block
For t = 0 To 15
wordText = MidB$(blockText, (t * 4) + 1, 4)
W(t).B0 = AscB(MidB$(wordText, 1, 1))
W(t).B1 = AscB(MidB$(wordText, 2, 1))
W(t).B2 = AscB(MidB$(wordText, 3, 1))
W(t).B3 = AscB(MidB$(wordText, 4, 1))
Next

' create extra words from the message block
For t = 16 To 79
' W(t) = S^1 (W(t-3) XOR W(t-8) XOR W(t-14) XOR W(t-16))
W(t) = CircShiftLeftW(XorW(XorW(XorW(W(t - 3), W(t - 8)), _
W(t - 14)), W(t - 16)), 1)
Next

' make initial assignments to the buffer
A = H0
b = H1
c = H2
D = H3
E = H4

' process the block
For t = 0 To 79
temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), _
F(t, b, c, D)), E), W(t)), k(t \ 20))
E = D
D = c
c = CircShiftLeftW(b, 30)
b = A
A = temp
Next

H0 = AddW(H0, A)
H1 = AddW(H1, b)
H2 = AddW(H2, c)
H3 = AddW(H3, D)
H4 = AddW(H4, E)
Next

StringSHA1 = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) _
& WordToHex(H3) & WordToHex(H4)

End Function

Public Function SHA1(inMessage() As Byte) As Byte()
' ???????SHA1??
Dim inLen As Long
Dim inLenW As Word
Dim numBlocks As Long
Dim W(0 To 79) As Word
Dim blockText As String
Dim wordText As String
Dim t As Long
Dim temp As Word
Dim k(0 To 3) As Word
Dim H0 As Word
Dim H1 As Word
Dim H2 As Word
Dim H3 As Word
Dim H4 As Word
Dim A As Word
Dim b As Word
Dim c As Word
Dim D As Word
Dim E As Word
Dim i As Long
Dim lngPos As Long
Dim lngPadMessageLen As Long
Dim padMessage() As Byte

inLen = UBound(inMessage) + 1
inLenW = DoubleToWord(CDbl(inLen) * 8)

lngPadMessageLen = inLen + 1 + (128 - (inLen Mod 64) - 9) Mod 64 + 8
ReDim padMessage(lngPadMessageLen - 1) As Byte
For i = 0 To inLen - 1
padMessage(i) = inMessage(i)
Next i
padMessage(inLen) = 128
padMessage(lngPadMessageLen - 4) = inLenW.B0
padMessage(lngPadMessageLen - 3) = inLenW.B1
padMessage(lngPadMessageLen - 2) = inLenW.B2
padMessage(lngPadMessageLen - 1) = inLenW.B3

numBlocks = lngPadMessageLen / 64

' initialize constants
k(0) = HexToWord("5A827999")
k(1) = HexToWord("6ED9EBA1")
k(2) = HexToWord("8F1BBCDC")
k(3) = HexToWord("CA62C1D6")

' initialize 160-bit (5 words) buffer
H0 = HexToWord("67452301")
H1 = HexToWord("EFCDAB89")
H2 = HexToWord("98BADCFE")
H3 = HexToWord("10325476")
H4 = HexToWord("C3D2E1F0")

' each 512 byte message block consists of 16 words (W) but W is expanded
' to 80 words
For i = 0 To numBlocks - 1
' initialize a message block
For t = 0 To 15
W(t).B0 = padMessage(lngPos)
W(t).B1 = padMessage(lngPos + 1)
W(t).B2 = padMessage(lngPos + 2)
W(t).B3 = padMessage(lngPos + 3)
lngPos = lngPos + 4
Next

' create extra words from the message block
For t = 16 To 79
' W(t) = S^1 (W(t-3) XOR W(t-8) XOR W(t-14) XOR W(t-16))
W(t) = CircShiftLeftW(XorW(XorW(XorW(W(t - 3), W(t - 8)), _
W(t - 14)), W(t - 16)), 1)
Next

' make initial assignments to the buffer
A = H0
b = H1
c = H2
D = H3
E = H4

' process the block
For t = 0 To 79
temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), _
F(t, b, c, D)), E), W(t)), k(t \ 20))
E = D
D = c
c = CircShiftLeftW(b, 30)
b = A
A = temp
Next

H0 = AddW(H0, A)
H1 = AddW(H1, b)
H2 = AddW(H2, c)
H3 = AddW(H3, D)
H4 = AddW(H4, E)
Next
Dim byt(19) As Byte
CopyMemory byt(0), H0, 4
CopyMemory byt(4), H1, 4
CopyMemory byt(8), H2, 4
CopyMemory byt(12), H3, 4
CopyMemory byt(16), H4, 4
SHA1 = byt
End Function
Sub testCRC32()
'Dim c As New EasyCrc
Dim b() As Byte
    b = "Test"
'    Debug.Print c.CalcCrc32Array(StrConv("Test", vbFromUnicode))
'Set c = Nothing
End Sub

Sub Test()
Dim D As Date, dd As Double
    D = Now
Debug.Print D
    dd = D
Debug.Print dd
End Sub
