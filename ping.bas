Attribute VB_Name = "ping"
Private Const INADDR_NONE As Long = -1
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const NULL_VALUE As Long = 0 'Null pointers, handles.  Just a 0.

Public Enum PING_FAIL_REASON_ENUM
    PFR_NONE = 0
    PFR_BAD_IP
    PFR_ICMPCREATEFILE
    PFR_ICMPSENDECHO
    PFR_ICMPCLOSEHANDLE
End Enum
'Preserve case of these identifiers:
#If False Then
Dim PFR_NONE, PFR_ICMPCREATEFILE, PFR_ICMPSENDECHO, PFR_ICMPCLOSEHANDLE
#End If

Public Enum IP_STATUS_ENUM
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = 11001
    IP_DEST_NET_UNREACHABLE = 11002
    IP_DEST_HOST_UNREACHABLE = 11003
    IP_DEST_PROT_UNREACHABLE = 11004
    IP_DEST_PORT_UNREACHABLE = 11005
    IP_NO_RESOURCES = 11006
    IP_BAD_OPTION = 11007
    IP_HW_ERROR = 11008
    IP_PACKET_TOO_BIG = 11009
    IP_REQ_TIMED_OUT = 11010
    IP_BAD_REQ = 11011
    IP_BAD_ROUTE = 11012
    IP_TTL_EXPIRED_TRANSIT = 11013
    IP_TTL_EXPIRED_REASSEM = 11014
    IP_PARAM_PROBLEM = 11015
    IP_SOURCE_QUENCH = 11016
    IP_OPTION_TOO_BIG = 11017
    IP_BAD_DESTINATION = 11018
    IP_GENERAL_FAILURE = 11050
End Enum
'Preserve case of these identifiers:
#If False Then
Dim IP_SUCCESS, IP_BUF_TOO_SMALL, IP_DEST_NET_UNREACHABLE, IP_DEST_HOST_UNREACHABLE
Dim IP_DEST_PROT_UNREACHABLE, IP_DEST_PORT_UNREACHABLE, IP_NO_RESOURCES, IP_BAD_OPTION
Dim IP_HW_ERROR, IP_PACKET_TOO_BIG, IP_REQ_TIMED_OUT, IP_BAD_REQ, IP_BAD_ROUTE
Dim IP_TTL_EXPIRED_TRANSIT, IP_TTL_EXPIRED_REASSEM, IP_PARAM_PROBLEM, IP_SOURCE_QUENCH
Dim IP_OPTION_TOO_BIG, IP_BAD_DESTINATION, IP_GENERAL_FAILURE
#End If

Public Enum RESOLVE_ERROR_ENUM
    RES_SUCCESS = 0
    RES_FORMATTING_ERR = 1
    WSAEINTR = 10004
    WSAEFAULT = 10014
    WSAEINPROGRESS = 10036
    WSAENETDOWN = 10050
    WSAEPROCLIM = 10067
    WSASYSNOTREADY = 10091
    WSAVERNOTSUPPORTED = 10092
    WSANOTINITIALISED = 10093
    WSAHOST_NOT_FOUND = 11001
    WSATRY_AGAIN = 11002
    WSANO_RECOVERY = 11003
    WSANO_DATA = 11004
End Enum
'Protect case of these identifiers:
#If False Then
Dim RES_SUCCESS, RES_FORMATTING_ERR, WSAEINTR, WSAEFAULT, WSAEINPROGRESS, WSAENETDOWN
Dim WSAEPROCLIM, WSASYSNOTREADY, WSAVERNOTSUPPORTED, WSANOTINITIALISED, WSAHOST_NOT_FOUND
Dim WSATRY_AGAIN, WSANO_RECOVERY, WSANO_DATA
#End If

Private Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 127) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type IP_OPTION_INFORMATION
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long 'Pointer.
End Type

Public Type ICMP_ECHO_REPLY
    Address As Long
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    Data As Long
    Options As IP_OPTION_INFORMATION
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

'NULL_VALUE on failure, call WSAGetLastError, else pointer to hostent.
Private Declare Function gethostbyname Lib "wsock32" (ByVal name As String) As Long

'0 on failure, check LastDLLError.
Private Declare Function IcmpCloseHandle Lib "Icmp" (ByVal IcmpHandle As Long) As Long

'INVALID_HANDLE_VALUE on failure, check LastDLLError.
Private Declare Function IcmpCreateFile Lib "Icmp" () As Long

'0 on failure, check LastDLLError.
Private Declare Function IcmpSendEcho Lib "Icmp" ( _
    ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ByRef ReplyBuffer As Byte, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long

'INADDR_NONE on failure.
Private Declare Function inet_addr Lib "wsock32" (ByVal cp As String) As Long

'NULL_VALUE on failure, else pointer to IP string.
Private Declare Function inet_ntoa Lib "wsock32" (ByVal inAddr As Long) As Long

Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As Long, _
    ByVal iMaxLength As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As Long) As Long

'Non-0 on failure, result is error number.
Private Declare Function WSACleanup Lib "wsock32" () As Long

'Non-0 on failure, result is error number.
Private Declare Function WSAStartup Lib "wsock32" ( _
    ByVal wVersionRequested As Integer, _
    ByRef lpWSAData As WSAData) As Long

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByRef dest As Any, _
    ByVal numBytes As Long)

Public Reason As PING_FAIL_REASON_ENUM 'When = PFR_ICMPSENDECHO, Status may have
                                       'IP_STATUS_ENUM values.
Public Reply As ICMP_ECHO_REPLY
Public Status As IP_STATUS_ENUM        'May contain system error numbers or
                                       'IP_STATUS_ENUM values.


Public Function PingData( _
    ByVal IP As String, _
    Optional ByVal TimeoutMS As Long = 1000, _
    Optional ByVal Data As String = "") As Boolean
    
    Dim IPAddr As Long
    Dim hIcmp As Long
    Dim BufSize As Long
    Dim Buffer() As Byte
    Dim Replies As Long
    
    'Clear any prior call's reply and status.
    ZeroMemory Reply, LenB(Reply)
    Status = 0
    
    IPAddr = inet_addr(IP)
    If IPAddr = INADDR_NONE Then
        Reason = PFR_BAD_IP
    Else
        hIcmp = IcmpCreateFile()
        If hIcmp = INVALID_HANDLE_VALUE Then
            Reason = PFR_ICMPCREATEFILE
            Status = Err.LastDllError
        Else
            BufSize = LenB(Reply) + Len(Data) + 8
            ReDim Buffer(BufSize - 1)
            Replies = IcmpSendEcho(hIcmp, IPAddr, Data, Len(Data), _
                                   NULL_VALUE, Buffer(0), BufSize, TimeoutMS)
            If Replies = 0 Then
                Reason = PFR_ICMPSENDECHO
                Status = Err.LastDllError
            Else
                CopyMemory Reply, Buffer(0), LenB(Reply)
                PingData = True
            End If
            If IcmpCloseHandle(hIcmp) = 0 Then
                PingData = False
                Reason = PFR_ICMPCLOSEHANDLE
                Status = Err.LastDllError
            End If
        End If
    End If
End Function

Public Function ResolveData(ByVal NameOrIP As String, ByRef IP As String) As RESOLVE_ERROR_ENUM
    'Returns RES_SUCCESS (0) on good result, else error number.
    Dim IPAddr As Long
    Dim wsadStartup As WSAData
    Dim pHeResolve As Long
    Dim heResolve As hostent
    Dim pAddrList As Long
    Dim pIPString As Long
    Dim IPStringLength As Long
    Dim NewIP As String
    
    NameOrIP = Trim$(NameOrIP)
    IPAddr = inet_addr(NameOrIP)
    If IPAddr = INADDR_NONE Then
        ResolveData = WSAStartup(&H202, wsadStartup) 'Possibly a WSA error.
        If ResolveData = 0 Then
            pHeResolve = gethostbyname(NameOrIP)
            If pHeResolve = NULL_VALUE Then
                ResolveData = Err.LastDllError 'A WSA error.
            Else
                CopyMemory heResolve, ByVal pHeResolve, LenB(heResolve)
                CopyMemory pAddrList, ByVal heResolve.h_addr_list, LenB(pAddrList)
                CopyMemory IPAddr, ByVal pAddrList, LenB(IPAddr)
                pIPString = inet_ntoa(IPAddr)
                If pIPString = NULL_VALUE Then
                    ResolveData = RES_FORMATTING_ERR
                Else
                    IPStringLength = lstrlen(pIPString)
                    NewIP = Space$(IPStringLength)
                    pIPString = lstrcpyn(NewIP, pIPString, IPStringLength + 1)
                    If pIPString = NULL_VALUE Then
                        ResolveData = RES_FORMATTING_ERR
                    Else
                        IP = NewIP
                    End If
                End If
            End If
            If WSACleanup() <> 0 Then ResolveData = Err.LastDllError 'A WSA error.
        End If
    Else
        IP = NameOrIP
    End If
End Function




