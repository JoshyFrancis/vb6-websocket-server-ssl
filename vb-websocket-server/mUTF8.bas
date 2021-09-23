Attribute VB_Name = "mUTF8"
Option Explicit


'UTF-8
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
'utf8?unicode
Public Function ToUnicodeData(ByRef Utf() As Byte) As Byte()
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim BT() As Byte
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2 - 1
    ReDim BT(lBufferSize)
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, VarPtr(BT(0)), lBufferSize + 1)
    If lret <> 0 Then
        ReDim Preserve BT(lret - 1)
        ToUnicodeData = BT
    End If
End Function
'utf8?unicode
Public Function ToUnicodeString(ByRef Utf() As Byte) As String
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    On Error GoTo errline:
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    ToUnicodeString = String$(lBufferSize, Chr(0))
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(ToUnicodeString), lBufferSize)
    If lret <> 0 Then
        ToUnicodeString = Left(ToUnicodeString, lret)
    End If
    Exit Function
errline:
    ToUnicodeString = ""
End Function
'unicode?utf8
Public Function Encoding(ByVal UCS As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    lLength = Len(UCS)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
    If lResult <> 0 Then
    lResult = lResult - 1
    ReDim Preserve abUTF8(lResult)
    Encoding = abUTF8
    End If
End Function
Public Function ConvertFromUTF8(ByRef Source() As Byte) As String
  Dim Size As Long
  Dim Pointer As Long
  Dim Length As Long
  Dim Buffer As String
  
  Size = UBound(Source) - LBound(Source) + 1
  Pointer = VarPtr(Source(LBound(Source)))
    
  If pUTF8header(Source) Then
    ' Ignore BOM header
    ' http://en.wikipedia.org/wiki/Byte_order_mark
    Size = Size - 3
    Pointer = Pointer + 3
  End If
    
  Length = MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, 0, 0)
  Buffer = Space$(Length)
  MultiByteToWideChar CP_UTF8, 0, Pointer, Size, StrPtr(Buffer), Length
  ConvertFromUTF8 = Buffer
End Function

'---------------------------------------------------------------------------------------
' Procedure : pUTF8header
' DateTime  : 14-6-2013
' Reference : http://en.wikipedia.org/wiki/Byte_order_mark
'---------------------------------------------------------------------------------------
Private Function pUTF8header(Source() As Byte) As Boolean

  If UBound(Source) >= 2 Then
    If Source(0) = &HEF Then
      If Source(1) = &HBB Then
        If Source(2) = &HBF Then
          pUTF8header = True
          Exit Function
        End If
      End If
    End If
  End If

End Function
