VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mySocket As cAsyncSocket
Attribute mySocket.VB_VarHelpID = -1
Event OnResolve(IpAddress As String)
Event OnAccept()
Event OnClose()
Event OnConnect()
Event OnReceive()
Event OnSend()
Event OnError(ByVal ErrorCode As Long, ByVal EventMask As Long)
Event OnMessagePending(Handled As Boolean)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private m_hSocket               As Long
Private Const STR_CHR1              As String = "" '--- CHAR(1)
Const ucsScpUtf8 = 65001
Private Sub mySocket_OnAccept()
    RaiseEvent OnAccept
End Sub

Private Sub mySocket_OnClose()
    RaiseEvent OnClose
End Sub

Private Sub mySocket_OnConnect()
    RaiseEvent OnConnect
End Sub

Private Sub mySocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As Long)
    RaiseEvent OnError(ErrorCode, EventMask)
End Sub

Private Sub mySocket_OnMessagePending(Handled As Boolean)
    RaiseEvent OnMessagePending(Handled)
End Sub

Private Sub mySocket_OnReceive()
    RaiseEvent OnReceive
End Sub

Private Sub mySocket_OnResolve(IpAddress As String)
    RaiseEvent OnResolve(IpAddress)
End Sub

Private Sub mySocket_OnSend()
    RaiseEvent OnSend
End Sub
Property Get SocketHandle() As Variant
    SocketHandle = CVar(mySocket.SocketHandle)
End Property

Property Get LastError() As Variant
    LastError = CVar(mySocket.LastError)
End Property

Property Get SockOpt(ByVal OptionName As Variant, Optional ByVal Level As Integer) As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        If IsMissing(Level) Then Level = &HFFFF&
        SockOpt = CVar(mySocket.SockOpt(CLng(OptionName), CLng(Level)))
    End If
End Property

Property Let SockOpt(ByVal OptionName As Variant, Optional ByVal Level As Integer, ByVal Value As Variant)
    If m_hSocket = 0 Then
        socket_not_init
    Else
        If IsMissing(Level) Then Level = &HFFFF&
        mySocket.SockOpt(CLng(OptionName), CLng(Level)) = CLng(Value)
    End If
End Property

Property Get AvailableBytes() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        AvailableBytes = CVar(mySocket.AvailableBytes)
    End If
End Property

Property Get HasPendingEvent() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        HasPendingEvent = CVar(mySocket.HasPendingEvent)
    End If
End Property

Property Get HasPendingResolve() As Variant
    If m_hSocket = 0 Then
        socket_not_init
    Else
        HasPendingResolve = mySocket.HasPendingResolve
    End If
End Property
Friend Property Get OwnAsyncSocket() As cAsyncSocket
    Set OwnAsyncSocket = mySocket
End Property
Friend Property Set OwnAsyncSocket(RHS As cAsyncSocket)
    m_hSocket = mySocket.SocketHandle
End Property
' Methods
Private Sub socket_not_init()
MyEr "Socket not created", "Ο σταθμός δεν έχει δημιουργηθεί"
End Sub
Public Function Create( _
            Optional ByVal SocketPort As Variant, _
            Optional ByVal SocketType As Variant, _
            Optional ByVal EventMask As Variant, _
            Optional SocketAddress As Variant) As Variant
        If IsMissing(SocketPort) Then SocketPort = 0
        If IsMissing(SocketType) Then SocketType = 0
        If IsMissing(EventMask) Then EventMask = 0
        Dim iSocketAddress As String
        If IsMissing(SocketAddress) Then iSocketAddress = vbNullString Else iSocketAddress = CVar(SocketAddress)
        If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
        Create = mySocket.Create(CLng(SocketPort), CLng(SocketType), CLng(EventMask), iSocketAddress)
        SocketAddress = iSocketAddress
End Function
Public Function GetPeerName(PeerAddress As Variant, PeerPort As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iPeerAddress As String, iPeerPort As Long
        GetPeerName = mySocket.GetPeerName(iPeerAddress, iPeerPort)
        PeerAddress = iPeerAddress
        PeerPort = iPeerPort
End Function
Public Function GetSockName(SocketAddress As Variant, SocketPort As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iSocketAddress As String, iSocketPort As Long
        GetSockName = mySocket.GetSockName(iSocketAddress, iSocketPort)
        SocketAddress = iSocketAddress
        SocketPort = iSocketPort
End Function
Public Function GetLocalHost(HostName As Variant, HostAddress As Variant) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iHostName As String, iHostAddress As String
        GetLocalHost = mySocket.GetLocalHost(iHostName, iHostAddress)
        HostName = iHostName
        HostAddress = iHostAddress
End Function
Public Function Accept(ConnectedSocket As Object, Optional SocketAddress As Variant, Optional SocketPort As Variant) As Variant
'' ConnectedSocket
    Dim iSocketAddress As String, iSocketPort As Long
    Dim iConnectedSocket As Socket, cSocket As cAsyncSocket
    If Typename(ConnectedSocket) <> "Socket" Then
    MyEr "Not a Socket Object", "Δεν βρήκα αντικείμενο ΣΤΑΘΜΟΣ"
    Exit Function
    End If
        Set iConnectedSocket = ConnectedSocket
        If iConnectedSocket.OwnAsyncSocket Is Nothing Then
            Set cSocket = New cAsyncSocket
            Set iConnectedSocket.OwnAsyncSocket = cSocket
            Accept = mySocket.Accept(cSocket, iSocketAddress, iSocketPort)
        End If
        SocketAddress = iSocketAddress
        SocketPort = iSocketPort
    
End Function
Public Function Bind(Optional SocketAddress As Variant, Optional ByVal SocketPort As Variant) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iSocketAddress As String, iSocketPort As Long
    If Not IsMissing(SocketPort) Then iSocketPort = SocketPort
        Bind = mySocket.Bind(iSocketAddress, iSocketPort)
        SocketAddress = iSocketAddress
    
End Function
Public Function Connect(HostAddress As Variant, ByVal HostPort As Variant) As Variant
    Dim iHostAddress As String, iHostPort As Long
    iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
    Connect = mySocket.Connect(iHostAddress, iHostPort)
    HostAddress = iHostAddress
End Function
Public Function IOCtl(ByVal Command As Variant, Argument As Variant) As Boolean
    Dim iCommmand As Long, iArgument As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    iCommmand = Command
    iArgument = iArgument
    If m_hSocket = 0 Then Set mySocket = New cAsyncSocket: m_hSocket = mySocket.SocketHandle
    IOCtl = mySocket.IOCtl(iCommmand, iArgument)
    Argument = iArgument
End Function
Public Function Listen(Optional ByVal ConnectionBacklog As Variant) As Variant
    Dim iConnectionBacklog As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If IsMissing(ConnectionBacklog) Then iConnectionBacklog = 5 Else iConnectionBacklog = ConnectionBacklog
    Listen = mySocket.Listen(iConnectionBacklog)
End Function
Public Function ReceiveText( _
            Optional HostAddress, _
            Optional HostPort, _
            Optional ByVal CodePage) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iHostAddress As String, iHostPort As Long, iCodePage As Long
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CodePage
    ReceiveText = mySocket.ReceiveText(iHostAddress, iHostPort, iCodePage)
    HostAddress = iHostAddress
    HostPort = iHostPort
End Function
' for buffers change to pass buffer object
' BufPtr from 0 as a multiple of basic item length
Public Function Receive( _
            Buf, _
            ByVal BufPtr, _
            ByVal BufLen, _
            Optional HostAddress, _
            Optional HostPort, _
            Optional ByVal Flags) As Variant
    Receive = 0&
    If Not MyIsObject(Buf) Then
100
        MyEr "need a Buffer", "χρειάζομαι μια διάρθρωση": Exit Function
    ElseIf Not TypeOf Buf Is MemBlock Then
        GoTo 100
    End If
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim b As MemBlock
    Set b = Buf
    Dim iBufPtr As Long, iBufLen As Long, iHostAddress As String, iHostPort As Long, iFlags As Long
    iBufPtr = b.GetPtr(BufPtr)
    If iBufPtr = 0 Then
        MyEr "offest is not in Buffer", "η αρχική διεύθυνση είναι εκτός διάρθρωσης": Exit Function
    End If
    If Not b.ValidArea2(CDbl(BufPtr), CDbl(BufLen)) Then
        MyEr "need a bigger Buffer", "χρειάζομαι μια μεγαλύτερη διάρθρωση": Exit Function
    End If
    iBufLen = BufLen
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If Not IsMissing(Flags) Then iFlags = Flags
    Receive = mySocket.Receive(iBufPtr, iBufLen, iHostAddress, iHostPort, iFlags)
    HostAddress = iHostAddress
    HostPort = iHostPort

End Function
Public Function SendText( _
            Text, _
            Optional HostAddress, _
            Optional ByVal HostPort, _
            Optional ByVal CodePage As Long) As Variant
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim iText As String, iHostAddress As String, iHostPort As Long, iCodePage As Long
    iText = Text
    iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If IsMissing(CodePage) Then iCodePage = ucsScpUtf8 Else iCodePage = CodePage
    SendText = mySocket.SendText(iText, iHostAddress, iHostPort, iCodePage)
    Text = iText
    HostAddress = iHostAddress
End Function
'change to send from buffer
Public Function Send( _
            Buf, _
            ByVal BufPtr, _
            ByVal BufLen, _
            Optional HostAddress, _
            Optional ByVal HostPort, _
            Optional ByVal Flags) As Variant
    Send = 0&
    If Not MyIsObject(Buf) Then
100
        MyEr "need a Buffer", "χρειάζομαι μια διάρθρωση": Exit Function
    ElseIf Not TypeOf Buf Is MemBlock Then
        GoTo 100
    End If
    If m_hSocket = 0 Then socket_not_init: Exit Function
    Dim b As MemBlock
    Set b = Buf
     Dim iBufPtr As Long, iBufLen As Long, iHostAddress As String, iHostPort As Long, iFlags As Long
    iBufPtr = b.GetPtr(BufPtr)
    If iBufPtr = 0 Then
        MyEr "offest is not in Buffer", "η αρχική διεύθυνση είναι εκτός διάρθρωσης": Exit Function
    End If
    If Not b.ValidArea2(CDbl(BufPtr), CDbl(BufLen)) Then
        MyEr "need a bigger Buffer", "χρειάζομαι μια μεγαλύτερη διάρθρωση": Exit Function
    End If
    iBufLen = BufLen
    If IsMissing(HostAddress) Then iHostAddress = STR_CHR1 Else iHostAddress = HostAddress
    If Not IsMissing(HostPort) Then iHostPort = HostPort
    If Not IsMissing(Flags) Then iFlags = Flags
    Send = mySocket.Send(iBufPtr, iBufLen, iHostAddress, iHostPort, iFlags)
    HostAddress = iHostAddress
    
End Function
Public Function ShutDown(Optional ByVal How) As Variant
    Dim iHow As Long
    If m_hSocket = 0 Then socket_not_init: Exit Function
    If IsMissing(How) = 1 Then iHow = 1 Else iHow = How
    ShutDown = mySocket.ShutDown(iHow)
End Function
Public Function GetErrorDescription(ByVal ErrorCode) As Variant
        If m_hSocket = 0 Then socket_not_init: Exit Function
        Dim iErrorCode As Long
        iErrorCode = ErrorCode
        GetErrorDescription = mySocket.GetErrorDescription(iErrorCode)
End Function
'' Public Sub PostEvent(ByVal EventMask As UcsAsyncSocketEventMaskEnum)
'' sync part not yet

'Public Function SyncConnect(HostAddress As String, ByVal HostPort As Long, Optional ByVal Timeout As Long) As Boolean

'End Function



