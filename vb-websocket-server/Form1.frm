VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Websocket Server"
   ClientHeight    =   2952
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10116
   LinkTopic       =   "Form1"
   ScaleHeight     =   2952
   ScaleWidth      =   10116
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSendTextMessage 
      Caption         =   "Send Text Message"
      Enabled         =   0   'False
      Height          =   492
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   1812
   End
   Begin VB.CommandButton cmdCloseWebSockets 
      Caption         =   "Close Web Sockets"
      Height          =   492
      Left            =   7440
      TabIndex        =   4
      Top             =   840
      Width           =   2292
   End
   Begin VB.CommandButton cmdSendLengthTextMessage 
      Caption         =   "Send Lengthy Text Message"
      Enabled         =   0   'False
      Height          =   492
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   1812
   End
   Begin VB.TextBox Text1 
      Height          =   1332
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1440
      Width           =   7692
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start WSS Server"
      Height          =   516
      Left            =   252
      TabIndex        =   1
      Top             =   1680
      Width           =   1524
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start WS Server"
      Height          =   516
      Left            =   252
      TabIndex        =   0
      Top             =   924
      Width           =   1524
   End
   Begin WinsockTest.ctxWinsock ctxServer 
      Index           =   0
      Left            =   2604
      Top             =   840
      _ExtentX        =   677
      _ExtentY        =   677
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type ws_Data
    SocketIndex As Long
    bData() As Byte
    ContentLength As Long
    sData As String
End Type
Private a_q_ws() As ws_Data
Private Count_a_q_ws As Long


Private m_oRootCa               As cTlsSocket

Private webSocket_clients() As Long, websocket_count As Long
Private websocket_data_available As Boolean, DataLength As Long, isMasked As Boolean, mask() As Byte, last_mask_i As Long

Function Base64EncodeEX(str() As Byte) As String
    On Error GoTo over
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(str) + 1) Mod 3
    Length = UBound(str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (str(i) And &H3) * &H10 + (str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (str(i + 1) And &HF) * &H4 + (str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10 + (str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (str(Length + 1) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64EncodeEX = Base64EncodeEX + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function


'hybi10Encode and hybi10Decode taken from   https://github.com/bloatless/php-websocket
'/**
' * Simple WebSocket server implementation in PHP.
' *
' * @author Simon Samtleben <foo@bloatless.org>
' * @author Nico Kaiser <nico@kaiser.me>
' * @version 2.0
' */
Function hybi10Encode(payload() As Byte, Optional ByVal stype As String = "text", _
    Optional ByVal masked As Boolean = True) As Byte()
Dim frame() As Byte, frameLength As Long, payloadLength As Long, i As Long
Dim l(7) As Byte
        payloadLength = UBound(payload) + 1
        frameLength = 0
    ReDim frame(0)
    Select Case stype
        Case "text"
'            // first byte indicates FIN, Text-Frame (10000001):
            frame(0) = 129
        Case "close"
'            // first byte indicates FIN, Close Frame(10001000):
            frame(0) = 136
        Case "ping"
'            // first byte indicates FIN, Ping frame (10001001):
            frame(0) = 137
        Case "pong"
'            // first byte indicates FIN, Pong frame (10001010):
            frame(0) = 138
    End Select

'        // set mask and payload length (using 1, 3 or 9 bytes)
    If payloadLength > 65535 Then
'        payloadLengthBin = str_split(sprintf('%064b', payloadLength), 8);
'        frame[1] = (masked === true) ? 255 : 127;
'        for (i = 0; i < 8; i++) {
'            frame[i + 2] = bindec(payloadLengthBin[i]);
'        }
'        // most significant bit MUST be 0 (close connection if frame too big)
'        if (frame[2] > 127) {
'            this->close(1004);
'            throw new \RuntimeException('Invalid payload. Could not encode frame.');
'        }
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = IIf(masked, 255, 127)
            frameLength = frameLength + 8
            ReDim Preserve frame(frameLength)
'        Dim sBin As String, sPart As String
'            sBin = DecimalToBinary(payloadLength)
'        Dim remainder As Long
'            remainder = payloadLength Mod 256
'                sPart = sBin
'        For i = 7 To 0 Step -1
'            frame(i + 2) = BinaryToDecimal(Right(sPart, 8))
'            sPart = Left(sPart, Len(sPart) - 8)
'            If sPart = "" Then
'                Exit For
'            End If
'        Next
        CopyMemory l(0), payloadLength, 4
        For i = 7 To 0 Step -1
            frame(i + 2) = l(7 - i)
        Next
    ElseIf payloadLength > 125 Then
'        payloadLengthBin = str_split(sprintf('%016b', payloadLength), 8);
'        frame[1] = (masked === true) ? 254 : 126;
'        frame[2] = bindec(payloadLengthBin[0]);
'        frame[3] = bindec(payloadLengthBin[1]);
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = IIf(masked, 254, 126)
'            sBin = DecimalToBinary(payloadLength, 2, IIf(payloadLength < 256, 16, 8))
'
'        frameLength = frameLength + 1
'        ReDim Preserve frame(frameLength)
'        frame(frameLength) = BinaryToDecimal(Mid$(sBin, 1, 8)) ' payloadLength And Not 255
'        frameLength = frameLength + 1
'        ReDim Preserve frame(frameLength)
'        frame(frameLength) = BinaryToDecimal(Mid$(sBin, 9, 8))  'payloadLength And 255
        frameLength = frameLength + 2
            ReDim Preserve frame(frameLength)
        CopyMemory l(0), payloadLength, 2
        For i = 1 To 0 Step -1
            frame(i + 2) = l(1 - i)
        Next
    Else
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = IIf(masked, payloadLength + 128, payloadLength)
    End If

'        // convert frame-head to string:
Dim eMask(3) As Byte
        If masked = True Then
'            // generate a random mask:
                Randomize
            For i = 0 To 3
'                mask[i] = chr(rand(0, 255));
                eMask(i) = CByte(Rnd * 255)
            Next
            
'            frame = array_merge(frame, mask);
            frameLength = frameLength + 4
            ReDim Preserve frame(frameLength)
            CopyMemory frame(frameLength - 3), eMask(0), 4
            
        End If
        
'        // append payload to frame:
        frameLength = frameLength + payloadLength
        ReDim Preserve frame(frameLength)
        For i = 0 To payloadLength - 1
'            frame .= (masked === true) ? payload[i] ^ mask[i % 4] : payload[i];
            If masked = True Then
                frame((frameLength - (payloadLength - 1)) + i) = payload(i) Xor eMask(i Mod 4) And 255
            Else
                frame((frameLength - (payloadLength - 1)) + i) = payload(i)
            End If
        Next

    hybi10Encode = frame
End Function

Public Function BinaryToDecimal(ByVal sBin As String, Optional ByVal BaseB As Currency = 2) As Currency
Dim A As Currency, c As Currency, D As Currency, E As Currency, F As Currency
    If sBin = "" Then
        Exit Function
    End If
Do
    A = Val(Right$(sBin, 1)) 'this is where the conversion happens
    sBin = Left$(sBin, Len(sBin) - 1) 'this is where the conversion happens
    If F > 49 Then 'overflow
        Exit Do
    End If
    c = BaseB ^ F 'conversion
    D = A * c 'conversion
    E = E + D 'conversion
    F = F + 1 'counter
Loop Until sBin = ""
BinaryToDecimal = E
End Function

Public Function DecimalToBinary(ByVal iDec As Currency, Optional ByVal BaseD As Currency = 2, _
    Optional ByVal bits As Long = 8) As String
Dim bit As Long, sBin As String, D As Currency
DecimalToBinary = ""
sBin = ""
    D = iDec
    bit = 1
Do
    sBin = (D Mod BaseD) & sBin   'conversion
    D = D \ BaseD 'conversion
            If bit = 8 Or D = 0 Then
                bit = 0
                DecimalToBinary = Right$(String$(bits, "0") & sBin, bits) & DecimalToBinary
                sBin = ""
            End If
        bit = bit + 1
Loop Until D = 0
    If sBin <> "" Then
        DecimalToBinary = Right$(String$(bits, "0") & sBin, bits) & DecimalToBinary
    End If
End Function


Function hybi10Decode(Data() As Byte, cw As ctxWinsock) As String
Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long, payloadLength As Long
Dim payloadOffset As Long, i As Long, b() As Byte, sData As String, j As Long
Dim ubData As Long, dataLengthInt As Integer, dataLengthCur As Currency, sBin As String
'    unmaskedPayload = '';
'        decodedData = [];
'        // estimate frame type:
            ubData = UBound(Data)
        firstByteBinary = Data(0)
        secondByteBinary = Data(1)
        
        opcode = firstByteBinary And Not 128
'        isMasked = secondByteBinary[0] === '1';
        isMasked = secondByteBinary And 128 > 0
        payloadLength = Data(1) And 127

'        // close connection if unmasked frame is received:
'        if (isMasked === false) {
'            this->close(1002);
'        }
        Select Case opcode
            Case 1 ' text frame:
            Case 2 'binary
            Case 8 'connection close frame
            Case 9 'ping frame
            Case 10 'pong frame
            Case Else
'                // Close connection on unknown opcode:
'                this->close(1003);
                Exit Function
        End Select

            ReDim mask(3)
        If payloadLength = 126 Then
'            mask = substr(data, 4, 4);
'            payloadOffset = 8;
'            dataLength = bindec(sprintf('%08b', ord(data[2])) . sprintf('%08b', ord(data[3]))) + payloadOffset;
            CopyMemory mask(0), Data(4), Len(Data(2)) * 4
            payloadOffset = 8
            DataLength = (Data(2) * 255) + (Data(3))
            DataLength = DataLength + payloadOffset
         ElseIf payloadLength = 127 Then
'            mask = substr(data, 10, 4);
            CopyMemory mask(0), Data(10), Len(Data(2)) * 4
            payloadOffset = 14
'            CopyMemory ByVal VarPtr(dataLengthCur), Data(2), 8
                dataLengthCur = 0
                sBin = ""
            For i = 0 To 7
''                tmp .= sprintf('%08b', ord(data[i + 2]));
'                If i < 7 Then
'                    dataLengthCur = dataLengthCur + (Data(i + 2) * 255)
'                Else
'                    dataLengthCur = dataLengthCur + (Data(i + 2))
'                End If
'                    If (Data(i + 2) And 1) Then
'                        dataLengthCur = (dataLengthCur + 1) * 255
'                    End If
'                    dataLengthCur = dataLengthCur + Data(i + 2)
                sBin = sBin & DecimalToBinary(Data(i + 2))
            Next
'            dataLength = bindec(tmp) + payloadOffset;
                 
            dataLengthCur = BinaryToDecimal(sBin)
                
            DataLength = dataLengthCur + payloadOffset
'            unset(tmp);
        Else
'            mask = substr(data, 2, 4);
            CopyMemory mask(0), Data(2), Len(Data(2)) * 4
            payloadOffset = 6
            DataLength = payloadLength + payloadOffset
            
        End If

'        /**
'         * We have to check for large frames here. socket_recv cuts at 1024 bytes
'         * so if websocket-frame is > 1024 bytes we have to wait until whole
'         * data is transferd.
'         */
'        if (strlen(data) < dataLength) {
'            return [];
'        }
'       If UBound(Data) < dataLength Then
'            Exit Function
'        End If
        If isMasked = True Then
'            sData = ""
''            For i = payloadOffset To dataLength - 1
'            For i = payloadOffset To ubData
'                j = i - payloadOffset
'                j = j Mod 4
''                sData = sData & Chr(Data(i) ^ mask(j Mod 4))
'                sData = sData & Chr(Data(i) Xor mask(j) And 255)
'            Next
'                dataLength = dataLength - payloadOffset
'                websocket_data_available = dataLength > ubData
            sData = decodeMasked(Data, payloadOffset)
             
        Else
            payloadOffset = payloadOffset - 4
            sData = Space(DataLength)
            CopyMemory ByVal VarPtr(sData), Data(payloadOffset), DataLength
        End If
'
'        return decodedData;
    hybi10Decode = sData
End Function
Function decodeMasked(Data() As Byte, ByVal offset As Long) As String
Dim sData As String, i As Long, j As Long, ubData As Long
        ubData = UBound(Data)
    If isMasked = True Then
        sData = ""
        For i = offset To ubData
            If websocket_data_available Then
                j = i + ((last_mask_i Mod 4) - 1) - offset
            Else
                j = i - offset
                    last_mask_i = i
            End If
            j = j Mod 4
            sData = sData & Chr(Data(i) Xor mask(j) And 255)
        Next
            DataLength = DataLength - offset
            websocket_data_available = DataLength > ubData
    Else
'        payloadOffset = payloadOffset - 4
'        sData = Space(dataLength)
'        CopyMemory ByVal VarPtr(sData), Data(payloadOffset), dataLength
    End If
        decodeMasked = sData
End Function

Public Sub SendTextMessageTo(ByVal Msg As String, Optional ByVal stype As String = "text")
  Dim b() As Byte, i As Long, webSocketIndex As Long
  b = StrConv(Msg, vbFromUnicode)
For i = 0 To websocket_count - 1
    webSocketIndex = webSocket_clients(i)
     ctxServer(webSocketIndex).SendData hybi10Encode(b, stype, False)   'A server must not mask any frames that it sends to the client.
Next
End Sub
Public Sub CloseWebSockets(Optional ByVal statusCode As Long = 1000)
  Dim b() As Byte, i As Long, webSocketIndex As Long, sBin As String, sPart As String, sPayLoad As String
        sPayLoad = ""
        sBin = DecimalToBinary(statusCode)
            sPart = sBin
        For i = 0 To 1
            sPayLoad = Chr(BinaryToDecimal(Right$(sPart, 8))) & sPayLoad
            sPart = Left$(sPart, Len(sPart) - 8)
            If sPart = "" Then
                Exit For
            End If
        Next
    Select Case statusCode
        Case 1000
            sPayLoad = sPayLoad & "normal closure"
        Case 1001
           sPayLoad = sPayLoad & "going away"
        Case 1002
            sPayLoad = sPayLoad & "protocol error"
        Case 1003
            sPayLoad = sPayLoad & "unknown data (opcode)"
        Case 1004
            sPayLoad = sPayLoad & "frame too large"
        Case 1007
            sPayLoad = sPayLoad & "utf8 expected"
        Case 1008
            sPayLoad = sPayLoad & "message violates server policy"
    End Select
  b = StrConv(sPayLoad, vbFromUnicode)
For i = 0 To websocket_count - 1
    webSocketIndex = webSocket_clients(i)
     ctxServer(webSocketIndex).SendData hybi10Encode(b, "close", False)    'A server must not mask any frames that it sends to the client.
Next
End Sub
Private Sub cmdCloseWebSockets_Click()
    CloseWebSockets
End Sub

Private Sub cmdSendLengthTextMessage_Click()
'    SendTextMessageTo "Test"
    SendTextMessageTo "<BEGIN>" & String(65535, "Test") & "<END>"
End Sub

  

Private Sub cmdSendTextMessage_Click()
    SendTextMessageTo "<BEGIN>" & String(125, "Test") & "<END>"
End Sub

Private Sub Command2_Click()
    Text1 = ""
    ctxServer(0).Close_
    ctxServer(0).Protocol = sckTCPProtocol
    ctxServer(0).Bind 8088, "127.0.0.1"
    ctxServer(0).Listen
    Shell "cmd /c start http://localhost:8088/"
End Sub

Private Sub Command4_Click()
    Text1 = ""
Dim str As String, F As Integer
    F = FreeFile
            str = ""
        Open App.Path & "\rootCA.crt" For Input As F
            str = Input(LOF(F), F)
        Close F
        Open App.Path & "\rootCA.key" For Input As F
            str = str & vbLf & Input(LOF(F), F)
        Close F
        Open App.Path & "\rootCA.pem" For Output As F
            Print #F, str
        Close F
    
    ctxServer(0).Close_
    
   ctxServer(0).pvSocket.PkiPemImportRootCaCertStore App.Path & "\rootCA.pem"
 
            
            str = ""
        Open App.Path & "\server.crt" For Input As F
            str = Input(LOF(F), F)
        Close F
        Open App.Path & "\server.key" For Input As F
            str = str & vbLf & Input(LOF(F), F)
        Close F
        Open App.Path & "\server.pem" For Output As F
            Print #F, str
        Close F
    
    ctxServer(0).Protocol = sckTLSProtocol
'    ctxServer(0).Bind 8088, "127.0.0.1"
    ctxServer(0).Bind 8088, "0.0.0.0"
'    ctxServer(0).Listen App.Path & "\localhost.pem", "", "1234"
    ctxServer(0).Listen App.Path & "\server.pem", "", "1234"
    Shell "cmd /c start https://localhost:8088/"
        Text1 = Text1 & vbCrLf & "Host : " & ctxServer(0).LocalHostName
        Text1 = Text1 & vbCrLf & "IP : " & ctxServer(0).LocalIP
        Text1 = Text1 & vbCrLf & "Port : " & ctxServer(0).LocalPort
        
End Sub
  

Private Sub ctxServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'    Debug.Print "ctxServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & ctxServer(Index).RemoteHostIP & ", RemotePort=" & ctxServer(Index).RemotePort, Timer
    Load ctxServer(ctxServer.UBound + 1)
    ctxServer(ctxServer.UBound).Protocol = ctxServer(Index).Protocol
    ctxServer(ctxServer.UBound).Accept requestID
End Sub

Private Sub ctxServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sRequest            As String
    Dim vSplit              As Variant
    Dim sBody               As String
    Dim bRequest() As Byte
'    Debug.Print "ctxServer_DataArrival, bytesTotal=" & bytesTotal, Timer
        bRequest = vbNullString
If bytesTotal > -1 Then
'    ctxServer(Index).GetData sRequest
    ctxServer(Index).GetData bRequest, vbByte + vbArray
End If
    Dim secKey As String, i As Long, b() As Byte, AcceptKey As String, Origin As String, Host As String
'    If sRequest Like "GET*HTTP/1.?*" Then
    If UBound(bRequest) > 0 And LikeB(bRequest, StrConv("GET*HTTP/1.?*", vbFromUnicode), 0, 100) Then
        sRequest = ctxServer(Index).pvSocket.FromTextArray(bRequest, ucsScpAcp)
        vSplit = Split(sRequest, vbCrLf)
        If InStr(vSplit(0), "/SubscribeToChatMsg") > 0 Then
            For i = 0 To UBound(vSplit)
                If InStr(LCase(vSplit(i)), "sec-websocket-key: ") > 0 Then
                    secKey = Trim(Mid(vSplit(i), InStr(vSplit(i), ":") + 1))
                End If
                If InStr(LCase(vSplit(i)), "origin: ") > 0 Then
                    Origin = Trim(Mid(vSplit(i), InStr(vSplit(i), ":") + 1))
                End If
                If InStr(LCase(vSplit(i)), "host: ") > 0 Then
                    Host = Trim(Mid(vSplit(i), InStr(vSplit(i), ":") + 1))
                End If
            Next
             
'            b = SHA1(StrConv(secKey & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11", vbFromUnicode))
            b = SHA1(StrConv(secKey & "258EAFA5-E914-47DA-95CA-C5AB0DC85B11", vbFromUnicode))
            AcceptKey = Base64EncodeEX(b)
            
                 
            ctxServer(Index).SendData "HTTP/1.1 101 Web Socket Protocol Handshake" & vbCrLf & _
                "Upgrade: websocket" & vbCrLf & _
                "Connection: Upgrade" & vbCrLf & _
                "Sec-WebSocket-Accept: " & AcceptKey & vbCrLf & _
                "Sec-WebSocket-Origin: " & Origin & vbCrLf & _
                "WebSocket-Location: " & Host & vbCrLf & _
                 vbCrLf
                    ReDim Preserve webSocket_clients(websocket_count)
                    webSocket_clients(websocket_count) = Index
                    websocket_count = websocket_count + 1
            cmdSendLengthTextMessage.Enabled = True
            cmdSendTextMessage.Enabled = True
        Else
'            Debug.Print vSplit(0)
            sBody = "<html><body><p>" & Join(vSplit, "</p>" & vbCrLf & "<p>" & Index & ": ") & "</p>" & vbCrLf & _
                "<p>" & Index & ": Current time is " & Now & "</p>" & _
                "<p>" & Index & ": RemoteHostIP is " & ctxServer(Index).RemoteHostIP & "</p>" & vbCrLf & _
                "<p>" & Index & ": RemotePort is " & ctxServer(Index).RemotePort & "</p>" & vbCrLf & _
                "</body></html>" & vbCrLf
                Dim F As Integer
                    F = FreeFile
                Open App.Path & "\wsTest.html" For Input As F
                    sBody = Input(LOF(F), F)
                Close F
                
            ctxServer(Index).SendData "HTTP/1.1 200 OK" & vbCrLf & _
                "Content-Type: text/html" & vbCrLf & _
                "Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
                sBody
        End If
    Else 'Handle websocket packets
        Dim c As Long, wsd As ws_Data, wi As Long
            wi = -1
            wi = Find_ws_Data_By_Index(Index)
            wsd.bData = vbNullString
        If wi <> -1 Then
            wsd = a_q_ws(wi)
            Remove_ws_Data wi
        End If
            wsd.SocketIndex = Index
            
        Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long
        Dim payloadOffset As Long, payloadLength As Long, l(3) As Byte
        Dim ubData As Long, dataLengthInt As Integer
        Dim DataLength As Long, isMasked As Boolean, mask() As Byte, hasContinuation As Boolean
        Dim j As Long, k As Long
        Dim bSuccess As Boolean, sData As String
'            If wsd.RawData <> "" Then
            k = UBound(wsd.bData)
            If k > -1 Then
'                sRequest = wsd.RawData & sRequest
'                wsd.RawData = ""
                j = UBound(bRequest)
                If j > -1 Then
                    mask = bRequest
                    bRequest = wsd.bData
                    ReDim Preserve bRequest((k + j + 1))
                    CopyMemory bRequest(k + 1), mask(0), j + 1
                Else
'                    ReDim bRequest(k)
'                    CopyMemory bRequest(0), wsd.bData(0), k + 1
                    bRequest = wsd.bData
                End If
                wsd.bData = vbNullString
            End If
'            Data = StrConv(sRequest, vbFromUnicode)
                ubData = UBound(bRequest)
                    hasContinuation = False
            If ubData > -1 Then
                firstByteBinary = bRequest(0)
                    If firstByteBinary = 1 Then
                        firstByteBinary = 129
                        'Has Continuation frames
                        hasContinuation = True
                    ElseIf firstByteBinary = 128 Then
                        'Is Continuation
                        hasContinuation = True
                    End If
                secondByteBinary = bRequest(1)
                opcode = firstByteBinary And Not 128
                isMasked = secondByteBinary And 128 > 0
                payloadLength = bRequest(1) And 127
                 
            End If
            Select Case opcode
                Case 0 'Continuation
                Case 1 ' text frame:
                Case 2 'binary
                Case 8 'connection close frame
                    'When a browser closes the Tab or Reload the Tab
                    Exit Sub
                Case 9 'ping frame
                Case 10 'pong frame
                Case Else
                        If bytesTotal < -4 Then 'Escape possible 'Out of Stack space' Error
                            'Send close frame
                            ctxServer(Index).SendData hybi10Encode(StrConv("", vbFromUnicode), "close", False)   'A server must not mask any frames that it sends to the client.
                            Exit Sub
                        End If
                    If ubData > -1 Then
'                        wsd.RawData = wsd.RawData & sRequest
                            k = UBound(wsd.bData)
                            j = UBound(bRequest)
                            If k > -1 Then
                                ReDim Preserve wsd.bData((j + k + 1))
                                CopyMemory wsd.bData(k + 1), bRequest(0), j + 1
                            Else
'                                ReDim Preserve wsd.bData(k)
'                                CopyMemory wsd.bData(0), bRequest(0), j + 1
                                wsd.bData = bRequest
                            End If
                        Add_ws_Data wsd
                        ctxServer_DataArrival Index, IIf(bytesTotal > 0, -1, bytesTotal - 1)
                        Exit Sub
                    End If
            End Select
                DataLength = 0
            If payloadLength = 126 Then
                payloadOffset = 8
                    If isMasked = False Then
                        payloadOffset = 4
                    End If
'                dataLengthCur = 0
'                    sBin = ""
'                For i = 0 To 1
'                    sBin = sBin & DecimalToBinary(Data(i + 2))
'                Next
'                dataLengthCur = BinaryToDecimal(sBin)
'                DataLength = dataLengthCur + payloadOffset
                    l(0) = bRequest(3)
                    l(1) = bRequest(2)
                     CopyMemory DataLength, l(0), 4
                DataLength = DataLength + payloadOffset
             ElseIf payloadLength = 127 Then
                payloadOffset = 14
                    If isMasked = False Then
                        payloadOffset = 10
                    End If
'                    dataLengthCur = 0
'                    sBin = ""
'                For i = 0 To 7
'                    sBin = sBin & DecimalToBinary(Data(i + 2))
'                Next
'                dataLengthCur = BinaryToDecimal(sBin)
'                DataLength = dataLengthCur + payloadOffset
                l(0) = bRequest(9)
                l(1) = bRequest(8)
                l(2) = bRequest(7)
                l(3) = bRequest(6)
                CopyMemory DataLength, l(0), 4
                DataLength = DataLength + payloadOffset
            ElseIf payloadLength > 0 Then
                payloadOffset = 6
                If isMasked = False Then
                    payloadOffset = 2
                End If
                DataLength = payloadLength + payloadOffset
            End If
'            If Len(sRequest) < DataLength Then
            If (ubData + 1) < DataLength Then
'                wsd.RawData = sRequest
                wsd.bData = bRequest
                Add_ws_Data wsd
                Exit Sub
            Else
'                sBin = deCodeFrame(Data, bSuccess, NextFrame)
                sData = ToUnicodeString(deCodeFrame(bRequest, bSuccess))
                If bSuccess = False Then
'                    wsd.RawData = sRequest
                    wsd.bData = bRequest
                Else
                    wsd.sData = wsd.sData & sData
                    wsd.ContentLength = wsd.ContentLength + (DataLength - payloadOffset)
                        If opcode = 0 And hasContinuation = True Then
                            hasContinuation = False
                        Else
                            If opcode = 0 Then
                                hasContinuation = True
                            End If
                        End If
                End If
            End If
        If hasContinuation = False And wsd.ContentLength = Len(wsd.sData) Then 'And Right$(wsd.sData, 1) = "}" Then
            web_socket_DataArrival wsd.sData, ctxServer(Index)
        Else
            Add_ws_Data wsd
            Exit Sub
        End If
    End If
'    Debug.Print "ctxServer_DataArrival, done", Timer
End Sub
'Function deCodeFrame(ByVal RawData As String, bSuccess As Boolean) As String
Function deCodeFrame(Data() As Byte, bSuccess As Boolean) As Byte()
Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long, payloadLength As Long
Dim payloadOffset As Long, sData As String, j As Long, l(3) As Byte
Dim ubData As Long, DataLength As Long, isMasked As Boolean, mask() As Byte
Dim offset As Long, i As Long, DataAvailable As Long
Dim mbuff() As Byte
    bSuccess = False
'                sData = ""
                mbuff = vbNullString
                ubData = UBound(Data)
                deCodeFrame = mbuff
            If ubData = -1 Then
                Exit Function
            End If
            firstByteBinary = Data(0)
            secondByteBinary = Data(1)
            
            opcode = firstByteBinary And Not 128
            isMasked = secondByteBinary And 128 > 0
            payloadLength = Data(1) And 127
    
            Select Case opcode
                Case 0 ' Continuation
                Case 1 ' text frame:
                Case 2 'binary
                Case 8 'connection close frame
                Case 9 'ping frame
                Case 10 'pong frame
                Case Else
                    Exit Function
            End Select
                ReDim mask(3)
            If payloadLength = 126 Then
                CopyMemory mask(0), Data(4), 4
                payloadOffset = 8
                    If isMasked = False Then
                        payloadOffset = 4
                    End If
'                dataLengthCur = 0
'                    sBin = ""
'                For i = 0 To 1
'                    sBin = sBin & DecimalToBinary(Data(i + 2))
'                Next
'                dataLengthCur = BinaryToDecimal(sBin)
'                DataLength = dataLengthCur + payloadOffset
                l(0) = Data(3)
                l(1) = Data(2)
                CopyMemory DataLength, l(0), 4
                DataLength = DataLength + payloadOffset
             ElseIf payloadLength = 127 Then
                CopyMemory mask(0), Data(10), 4
                payloadOffset = 14
                    If isMasked = False Then
                        payloadOffset = 10
                    End If
'                    dataLengthCur = 0
'                    sBin = ""
'                For i = 0 To 7
'                    sBin = sBin & DecimalToBinary(Data(i + 2))
'                Next
'                dataLengthCur = BinaryToDecimal(sBin)
'                DataLength = dataLengthCur + payloadOffset
                l(0) = Data(9)
                l(1) = Data(8)
                l(2) = Data(7)
                l(3) = Data(6)
                CopyMemory DataLength, l(0), 4
                DataLength = DataLength + payloadOffset
            Else
                CopyMemory mask(0), Data(2), 4
                payloadOffset = 6
                    If isMasked = False Then
                        payloadOffset = 2
                    End If
                DataLength = payloadLength + payloadOffset
            End If

        If DataLength > 0 And opcode < 3 Then
            If isMasked = True Then
                offset = payloadOffset
                DataAvailable = IIf((DataLength - 1) > ubData, ubData, DataLength - 1)
                If ((DataAvailable + 1)) <> DataLength Then
                    bSuccess = False
                    Exit Function
                End If
                ReDim mbuff(DataAvailable - offset)
                For i = offset To DataAvailable 'ubData
                        j = i - offset
                    j = j Mod 4
'                    sData = sData & Chr(Data(i) Xor mask(j) And 255)
                    mbuff(i - offset) = Data(i) Xor mask(j) And 255
                Next
'                DataAvailable = DataAvailable + offset
'                sData = ToUnicodeString(dataA)  '------------work ok'
            Else
                payloadOffset = payloadOffset - 4
                DataAvailable = IIf((DataLength - 1) > ubData, ubData, DataLength - 1)
                If ((DataAvailable + 1) - payloadOffset) <> DataLength Then
                    bSuccess = False
                    Exit Function
                End If
'                sData = Space(DataAvailable + 1)
'                CopyMemory ByVal VarPtr(sData), Data(payloadOffset), DataAvailable + 1 'DataLength
'                DataAvailable = DataAvailable + 4
                ReDim mbuff(DataAvailable)
                CopyMemory mbuff(0), Data(payloadOffset), DataAvailable + 1 'DataLength
'                sData = ToUnicodeString(mbuff)
            End If
        Else
            Exit Function
        End If
'    deCodeFrame = sData
    deCodeFrame = mbuff
        bSuccess = True
End Function
Private Function Find_ws_Data_By_Index(ByVal Index As Long) As Long
Dim i As Long, c As Long
    i = -1
For c = Count_a_q_ws - 1 To 0 Step -1
    If a_q_ws(c).SocketIndex = Index Then
        i = c
        Exit For
    End If
Next
    Find_ws_Data_By_Index = i
End Function
Private Function Remove_ws_Data(ByVal aIndex As Long) As Long
Dim c As Long, tmp() As ws_Data
    If Count_a_q_ws = 0 Then Exit Function
ReDim tmp(Count_a_q_ws - 1)
For c = 0 To aIndex - 1
    tmp(c) = a_q_ws(c)
Next
For c = aIndex + 1 To Count_a_q_ws - 1
    tmp(c - 1) = a_q_ws(c)
Next
    a_q_ws = tmp
        Erase tmp
Count_a_q_ws = Count_a_q_ws - 1
    Remove_ws_Data = Count_a_q_ws
End Function
Private Function Add_ws_Data(wsd As ws_Data) As Long
ReDim Preserve a_q_ws(Count_a_q_ws)
    a_q_ws(Count_a_q_ws) = wsd
Add_ws_Data = Count_a_q_ws
    Count_a_q_ws = Count_a_q_ws + 1
End Function
Sub web_socket_DataArrival(ByVal sData As String, ByVal cx As ctxWinsock)
  Dim b() As Byte, i As Long, webSocketIndex As Long
    If Len(sData) > 60 Then
        Text1 = Text1 & vbCrLf & Mid$(sData, 1, 30) & "..." & Mid$(sData, Len(sData) - 30)
    Else
        Text1 = Text1 & vbCrLf & sData
    End If
        
  b = StrConv("server time : " & Now, vbFromUnicode)
      cx.SendData hybi10Encode(b, "text", False)    'A server must not mask any frames that it sends to the client.
 
End Sub

Private Sub ctxServer_CloseEvent(Index As Integer)
Dim i As Long, c As Long, temp() As Long
For i = 0 To websocket_count - 1
    If webSocket_clients(i) <> Index Then
        ReDim Preserve temp(c)
        temp(c) = webSocket_clients(i)
        c = c + 1
    End If
Next
webSocket_clients = temp
websocket_count = c

    Unload ctxServer(Index)
End Sub

Private Sub ctxServer_Close(Index As Integer)
    ctxServer_CloseEvent Index
End Sub

Private Sub ctxWinsock_Error(ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Error: " & Description
End Sub
 
