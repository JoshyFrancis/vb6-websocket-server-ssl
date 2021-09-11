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
   Begin VB.CommandButton cmdCloseWebSockets 
      Caption         =   "Close Web Sockets"
      Height          =   492
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   2292
   End
   Begin VB.CommandButton cmdSendTextMessage 
      Caption         =   "Send Text Message"
      Enabled         =   0   'False
      Height          =   492
      Left            =   3480
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

Private Declare Function CreateNamedPipeW Lib "kernel32" (ByVal lpName As Long, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hPipe As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private hPipe As Long
Private iPipeFile As Long


Private Type ws_Data
    SocketIndex As Long
    RawData As String
    ContentLength As Long
    sData As String
End Type
Private q_ws As New Collection

Private m_oRootCa               As cTlsSocket

Private webSocket_clients() As Long, websocket_count As Long
Private websocket_data_available As Boolean, dataLength As Long, isMasked As Boolean, mask() As Byte, last_mask_i As Long

Function Base64EncodeEX(Str() As Byte) As String
    On Error GoTo over
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(Str) + 1) Mod 3
    Length = UBound(Str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10 + (Str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (Str(Length + 1) And &HF) * &H4
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
        Dim sBin As String, sPart As String
            sBin = DecimalToBinary(payloadLength)
        Dim remainder As Long
            remainder = payloadLength Mod 256
            frameLength = frameLength + 8
            ReDim Preserve frame(frameLength)
                sPart = sBin
        For i = 7 To 0 Step -1
            frame(i + 2) = BinaryToDecimal(Right(sPart, 8))
            sPart = Left(sPart, Len(sPart) - 8)
            If sPart = "" Then
                Exit For
            End If
        Next
    ElseIf payloadLength > 125 Then
'        payloadLengthBin = str_split(sprintf('%016b', payloadLength), 8);
'        frame[1] = (masked === true) ? 254 : 126;
'        frame[2] = bindec(payloadLengthBin[0]);
'        frame[3] = bindec(payloadLengthBin[1]);
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = IIf(masked, 254, 126)
            sBin = DecimalToBinary(payloadLength, 2, IIf(payloadLength < 256, 16, 8))
            
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = BinaryToDecimal(Mid$(sBin, 1, 8)) ' payloadLength And Not 255
        frameLength = frameLength + 1
        ReDim Preserve frame(frameLength)
        frame(frameLength) = BinaryToDecimal(Mid$(sBin, 9, 8))  'payloadLength And 255
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


Function hybi10Decode(data() As Byte, cw As ctxWinsock) As String
Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long, payloadLength As Long
Dim payloadOffset As Long, i As Long, b() As Byte, sData As String, j As Long
Dim ubData As Long, dataLengthInt As Integer, dataLengthCur As Currency, sBin As String
'    unmaskedPayload = '';
'        decodedData = [];
'        // estimate frame type:
            ubData = UBound(data)
        firstByteBinary = data(0)
        secondByteBinary = data(1)
        
        opcode = firstByteBinary And Not 128
'        isMasked = secondByteBinary[0] === '1';
        isMasked = secondByteBinary And 128 > 0
        payloadLength = data(1) And 127

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
            CopyMemory mask(0), data(4), Len(data(2)) * 4
            payloadOffset = 8
            dataLength = (data(2) * 255) + (data(3))
            dataLength = dataLength + payloadOffset
         ElseIf payloadLength = 127 Then
'            mask = substr(data, 10, 4);
            CopyMemory mask(0), data(10), Len(data(2)) * 4
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
                sBin = sBin & DecimalToBinary(data(i + 2))
            Next
'            dataLength = bindec(tmp) + payloadOffset;
                 
            dataLengthCur = BinaryToDecimal(sBin)
                
            dataLength = dataLengthCur + payloadOffset
'            unset(tmp);
        Else
'            mask = substr(data, 2, 4);
            CopyMemory mask(0), data(2), Len(data(2)) * 4
            payloadOffset = 6
            dataLength = payloadLength + payloadOffset
            
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
            sData = decodeMasked(data, payloadOffset)
             
        Else
            payloadOffset = payloadOffset - 4
            sData = Space(dataLength)
            CopyMemory ByVal VarPtr(sData), data(payloadOffset), dataLength
        End If
'
'        return decodedData;
    hybi10Decode = sData
End Function
Function decodeMasked(data() As Byte, ByVal offset As Long) As String
Dim sData As String, i As Long, j As Long, ubData As Long
        ubData = UBound(data)
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
            sData = sData & Chr(data(i) Xor mask(j) And 255)
        Next
            dataLength = dataLength - offset
            websocket_data_available = dataLength > ubData
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

Private Sub cmdSendTextMessage_Click()
'    SendTextMessageTo "Test"
    SendTextMessageTo "<BEGIN>" & String(65535, "Test") & "<END>"
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
Dim Str As String
            Str = ""
        Open App.Path & "\rootCA.crt" For Input As 1
            Str = Input(LOF(1), 1)
        Close 1
        Open App.Path & "\rootCA.key" For Input As 1
            Str = Str & vbLf & Input(LOF(1), 1)
        Close 1
        Open App.Path & "\rootCA.pem" For Output As 1
            Print #1, Str
        Close 1
    
    ctxServer(0).Close_
    
   ctxServer(0).pvSocket.PkiPemImportRootCaCertStore App.Path & "\rootCA.pem"

'
'            str = ""
'        Open App.Path & "\localhost.crt" For Input As 1
'            str = Input(LOF(1), 1)
'        Close 1
'        Open App.Path & "\localhost.key" For Input As 1
'            str = str & vbLf & Input(LOF(1), 1)
'        Close 1
'        Open App.Path & "\localhost.pem" For Output As 1
'            Print #1, str
'        Close 1
            
            Str = ""
        Open App.Path & "\server.crt" For Input As 1
            Str = Input(LOF(1), 1)
        Close 1
        Open App.Path & "\server.key" For Input As 1
            Str = Str & vbLf & Input(LOF(1), 1)
        Close 1
        Open App.Path & "\server.pem" For Output As 1
            Print #1, Str
        Close 1
    
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
    
'    Debug.Print "ctxServer_DataArrival, bytesTotal=" & bytesTotal, Timer
    ctxServer(Index).GetData sRequest
    Dim secKey As String, i As Long, b() As Byte, AcceptKey As String, Origin As String, Host As String
'    If Asc(sRequest) = &H81 Or websocket_data_available Then  '129 websocket send data
'                    b = StrConv(sRequest, vbFromUnicode)
'            If websocket_data_available And Asc(sRequest) <> &H81 Then
'                sBody = decodeMasked(b, 0)
'                Text1 = sBody
'            Else
'                websocket_data_available = False
'                sBody = hybi10Decode(b, ctxServer(Index))
'                Text1 = Text1 & vbCrLf & sBody
'            End If
'        Exit Sub
'    End If
'    If UBound(vSplit) >= 0 Then
    If sRequest Like "GET*HTTP/1.?*" Then
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
            cmdSendTextMessage.Enabled = True
        Else
'            Debug.Print vSplit(0)
            sBody = "<html><body><p>" & Join(vSplit, "</p>" & vbCrLf & "<p>" & Index & ": ") & "</p>" & vbCrLf & _
                "<p>" & Index & ": Current time is " & Now & "</p>" & _
                "<p>" & Index & ": RemoteHostIP is " & ctxServer(Index).RemoteHostIP & "</p>" & vbCrLf & _
                "<p>" & Index & ": RemotePort is " & ctxServer(Index).RemotePort & "</p>" & vbCrLf & _
                "</body></html>" & vbCrLf
                Open App.Path & "\wsTest.html" For Input As 1
                    sBody = Input(LOF(1), 1)
                Close 1
                
            ctxServer(Index).SendData "HTTP/1.1 200 OK" & vbCrLf & _
                "Content-Type: text/html" & vbCrLf & _
                "Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
                sBody
        End If
    Else 'Handle websocket packets
        Dim c As Long, wsd As ws_Data, wi As Long
            wi = -1
        For c = 1 To q_ws.Count
            DeserializeFromBytes q_ws.Item(c), wsd
            If wsd.SocketIndex = Index Then
                wi = c
                Exit For
            End If
        Next
        If wi <> -1 Then
             q_ws.Remove wi
        End If
          
        Dim data() As Byte
        Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long
        Dim payloadOffset As Long, payloadLength As Long
        Dim ubData As Long, dataLengthInt As Integer, dataLengthCur As Currency, sBin As String
                If sRequest = "" Then
                    sRequest = wsd.RawData
                    wsd.RawData = ""
                End If
            data = StrConv(sRequest, vbFromUnicode)
                ubData = UBound(data)
            If ubData > -1 Then
                firstByteBinary = data(0)
                secondByteBinary = data(1)
                opcode = firstByteBinary And Not 128
                isMasked = secondByteBinary And 128 > 0
                payloadLength = data(1) And 127
            End If
            Select Case opcode
                Case 1 ' text frame:
                Case 2 'binary
                Case 8 'connection close frame
                Case 9 'ping frame
                Case 10 'pong frame
                Case Else
                    If ubData > -1 Then
                        wsd.RawData = wsd.RawData & sRequest
                        wsd.SocketIndex = Index
                        q_ws.Add SerializeToBytes(wsd), "W" & Index
                        ctxServer_DataArrival Index, 0
                        Exit Sub
                    End If
            End Select
                dataLength = 0
            If payloadLength = 126 Then
                payloadOffset = 8
                dataLengthCur = 0
                    sBin = ""
                For i = 0 To 1
                    sBin = sBin & DecimalToBinary(data(i + 2))
                Next
                dataLengthCur = BinaryToDecimal(sBin)
                dataLength = dataLengthCur
             ElseIf payloadLength = 127 Then
                payloadOffset = 14
                    dataLengthCur = 0
                    sBin = ""
                For i = 0 To 7
                    sBin = sBin & DecimalToBinary(data(i + 2))
                Next
                dataLengthCur = BinaryToDecimal(sBin)
                dataLength = dataLengthCur
            ElseIf payloadLength > 0 Then
                payloadOffset = 6
                dataLength = payloadLength
            End If
    
            
            wsd.RawData = wsd.RawData & sRequest
        If UBound(data) + 1 < (dataLength + payloadOffset) Then
            wsd.SocketIndex = Index
            q_ws.Add SerializeToBytes(wsd), "W" & Index
            Exit Sub
        Else
            If wsd.sData = "" Then
                wsd.ContentLength = dataLength
            End If
            wsd.sData = wsd.sData & deCodeFrame(wsd.RawData)
        End If
         
        If Len(wsd.sData) = wsd.ContentLength Then 'Final Data
            web_socket_DataArrival wsd.sData, ctxServer(Index)
            Exit Sub
        End If
        
    End If
'    Debug.Print "ctxServer_DataArrival, done", Timer
End Sub
Sub web_socket_DataArrival(ByVal sData As String, ByVal cx As ctxWinsock)
  Dim b() As Byte, i As Long, webSocketIndex As Long

  b = StrConv("server time : " & Now, vbFromUnicode)
      cx.SendData hybi10Encode(b, "text", False)    'A server must not mask any frames that it sends to the client.
 
End Sub
Function deCodeFrame(ByVal RawData As String) As String
Dim data() As Byte
Dim firstByteBinary As Byte, secondByteBinary As Byte, opcode As Long, payloadLength As Long
Dim payloadOffset As Long, sData As String, j As Long
Dim ubData As Long, dataLengthInt As Integer, dataLengthCur As Currency, sBin As String
Dim offset As Long, i As Long
                sData = ""
            data = StrConv(RawData, vbFromUnicode)
                ubData = UBound(data)
            firstByteBinary = data(0)
            secondByteBinary = data(1)
            
            opcode = firstByteBinary And Not 128
            isMasked = secondByteBinary And 128 > 0
            payloadLength = data(1) And 127
    
            Select Case opcode
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
                CopyMemory mask(0), data(4), Len(data(2)) * 4
                payloadOffset = 8
'                dataLength = (data(2) * 255) + (data(3))
'                dataLength = dataLength + payloadOffset
                dataLengthCur = 0
                    sBin = ""
                For i = 0 To 1
                    sBin = sBin & DecimalToBinary(data(i + 2))
                Next
                dataLengthCur = BinaryToDecimal(sBin)
                dataLength = dataLengthCur + payloadOffset
             ElseIf payloadLength = 127 Then
                CopyMemory mask(0), data(10), Len(data(2)) * 4
                payloadOffset = 14
                    dataLengthCur = 0
                    sBin = ""
                For i = 0 To 7
                    sBin = sBin & DecimalToBinary(data(i + 2))
                Next
                dataLengthCur = BinaryToDecimal(sBin)
                dataLength = dataLengthCur + payloadOffset
            Else
                CopyMemory mask(0), data(2), Len(data(2)) * 4
                payloadOffset = 6
                dataLength = payloadLength + payloadOffset
            End If

        If dataLength > 0 Then
            If isMasked = True Then
                offset = payloadOffset
                    
                For i = offset To ubData
                    
                        j = i - offset
                            last_mask_i = i
                    
                    j = j Mod 4
                    sData = sData & Chr(data(i) Xor mask(j) And 255)
                Next
                    dataLength = dataLength - offset
                    websocket_data_available = dataLength > ubData
            Else
                payloadOffset = payloadOffset - 4
                 sData = Space(dataLength)
                CopyMemory ByVal VarPtr(sData), data(payloadOffset), dataLength
            End If
                 
        End If
    deCodeFrame = sData
End Function
Private Function SerializeToBytes(udt As ws_Data) As Byte()
    Dim Bytes As Long
    Dim b() As Byte
    '
    If iPipeFile = 0 Then Exit Function
    '
    Put iPipeFile, 1, udt
    PeekNamedPipe hPipe, ByVal 0&, 0, ByVal 0&, Bytes, 0
    If Bytes <= 0 Then Exit Function
    '
    ReDim Preserve b(0 To Bytes - 1)
    ReadFile hPipe, b(0), Bytes, Bytes, ByVal 0&
    SerializeToBytes = b
End Function

Private Sub DeserializeFromBytes(b() As Byte, udt As ws_Data)
    Dim Bytes As Long
    '
    If iPipeFile = 0 Then Exit Sub
    '
    WriteFile hPipe, b(0), UBound(b) + 1, Bytes, 0
    If Bytes <> UBound(b) + 1 Then Exit Sub
    '
    Get iPipeFile, 1, udt
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

