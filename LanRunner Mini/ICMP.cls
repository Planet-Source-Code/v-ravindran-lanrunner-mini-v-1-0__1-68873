VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Due to not having access to winsock 2's raw sockets in vb6 this class uses
' the ICMP.dll to allow users to perform ping and trace route. ICMP is a
' network layer protocol that delivers flow control, error messages, routing,
' and other data between Internet hosts. ICMP is primarily used by application
' developers for a network ping, which is also known as Packet Internet Groper.
' A ping is the process of sending an echo message to an IP address and reading
' the reply to verify a connection between TCP/IP hosts.
'
' Some of you will no doubt recognise the ideas behind this code from way back
' but other examples had code all over the shop. Making it hard to just shove
' this functionality into any app you wish. Also the other code had variables missing,
' code that served no purpose, stuff that made no sence at all etc etc etc. So
' hopefully this class is a one stop shop for the icmp.dll at least until someone
' sorts out a version using winsock 2 :o
'
' Notice!!!:
' That the functions in icmp.dll are not considered part of the Win32 API and will
' not be supported in future releases. Once we have a more complete solution in the
' operating system, this DLL, and the functions it exports, will be dropped.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





' To use this class simply create the object call tracert or ping with you required
' settings. Their are 6 events in this class that you should be aware of these are

' Inform the user that the status of the tracert has changed
Public Event TracertStatus(strStatus As String)

' When a new hop has been found on the trace route tell the user about it here
' strHopNumber its a string number to include the preceading zero
Public Event TracertResponce(strHopNumber As String, strRespondingHost As String, lngTimeToLive As Long)

' Ping status has changed, if the rcode is zero then their was a time out else it
' returns the error from icmp.
Public Event PingStatus(strRespondingHost As String, RCode As String)

' When a ping reaches its destination and gets a responce its info is displayed here.
Public Event PingResponce(strRespondingHost As String, strBytes As String, lngTimeToLiveMs As Long, lngRoundTripTimeMs As Long)

' any errors the class has are raised here, winsock startup failures, icmp failures
' but not the tracert ping errors. thier raised in status. So i guess really you
' could call this TransportError
Public Event Error(msg As String)

' Fired when the class starts with various stats on the winsock setup if possible
' note this doesnt seem to work and i have no idea why.... see the winsock startup
' method. Perhaps one of you will tell me why the raiseevent doesnt seem to work..
Public Event SockStartup(MaxSockets As Long, MaxUDP As Long, Description As String, Status As String)











''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' lets any one who wants to use the libary if their was an error with the winsock
' startup. Basicly lets the caller prevent using the class if it cant possibly run.
' Error messages are returned via the error event for processing by the caller.
' note to self: this is now private as you cant use the public methods now if this
' is false also the user already recieved an error if winsock failed to load so
' this way should work fine...
Private boolError As Boolean

' Used by ICMP send echo and its start and stop methods
Private hIP                                       As Long

' lets the whole class know if were doing a traceroute or ping
Private TraceRT                                   As Boolean

' Time To Live (used by tracert)
Private hopnumber As Long

' used by gethostbyname and icmpsendecho
Private Addr                                      As Long

' Used by getrcode, icmpsendecho
Private IPLong                                    As Inet_address
Private Type Inet_address
    Byte4                                             As String * 1
    Byte3                                             As String * 1
    Byte2                                             As String * 1
    Byte1                                             As String * 1
End Type










'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Used by gethostbyname
Private Type HostEnt
    h_name                      As Long
    h_aliases                   As Long
    h_addrtype                  As Integer
    h_length                    As Integer
    h_addr_list                 As Long
End Type
Private Const SOCKET_ERROR                        As Long = -1
Private Declare Function gethostname Lib "wsock32.dll" (ByVal hostname As String, _
                                                        HostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal hostname As String) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''









'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Winsock Related
Private Type WSAdata
    wVersion                                          As Integer
    wHighVersion                                      As Integer
    szDescription(0 To 255)                           As Byte
    szSystemStatus(0 To 128)                          As Byte
    iMaxSockets                                       As Integer
    iMaxUdpDg                                         As Integer
    lpVendorInfo                                      As Long
End Type
' WSA Startup constants
Private Const WS_VERSION_MAJOR                    As Long = &H101 \ &H100 And &HFF
Private Const WS_VERSION_MINOR                    As Long = &H101 And &HFF
Private Const MIN_SOCKETS_REQD                    As Long = 0
Private Const WSA_DESCRIPTION_LEN                 As Long = 256
Private Const WSA_SYS_STATUS_LEN                  As Long = 128

Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, _
                                                       lpWSAData As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                                                                                                                         
                                                                     
                                                                     
                                                                     
                                                                     
                                                                     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This function creates a handle on which Internet Control Message Protocol (ICMP)
' requests can be issued.

' An ICMP handle indicates success. INVALID_HANDLE_VALUE indicates failure.
Private Const INVALID_HANDLE_VALUE = -1
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Parameters
'   IcmpHandle
'       [in] ICMP handle opened by IcmpCreateFile.

'Return Values
'   TRUE indicates success. FALSE indicates failure.

Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This function sends an Internet Control Message Protocol (ICMP) echo request,
' and returns one or more replies.

'Parameters
'   IcmpHandle
'       [in] ICMP handle opened by IcmpCreateFile.
'   DestinationAddress
'       [in] Specifies the destination of the echo request.
'   RequestData
'       [in] Buffer that contains the data to be sent in the request.
'   RequestSize
'       [in] Number of bytes in the RequestData buffer.
'   RequestOptions
'       [in] Pointer to the IP header options for the request; may be NULL.
'   ReplyBuffer
'       [out] Buffer to hold any replies to the request. When the function returns,
'             the buffer will contain one or more ICMP_ECHO_REPLY structures, followed
'             by options and data.
'   ReplySize
'       [out] Size, in bytes, of the reply buffer. The buffer must be large enough to
'             accommodate at least one ICMP_ECHO_REPLY structure plus eight additional
'             bytes (the size of an ICMP error message).
'   TimeOut
'       [in] Time, in milliseconds, to wait for replies.

'Return Values
'   The number of replies received and stored in the reply buffer indicates success. Zero indicates failure. Extended error information is available through GetLastError.

Private Type IP_OPTION_INFORMATION
    TTL                                               As Byte
    Tos                                               As Byte
    flags                                             As Byte
    OptionsSize                                       As Long
    OptionsData                                       As String * 128
End Type

Private Type IP_ECHO_REPLY
    Address(0 To 3)                                   As Byte
    Status                                            As Long
    RoundTripTime                                     As Long
    DataSize                                          As Integer
    Reserved                                          As Integer
    data                                              As Long
    Options                                           As IP_OPTION_INFORMATION
End Type

Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, _
                                                  ByVal DestAddress As Long, _
                                                  ByVal RequestData As String, _
                                                  ByVal RequestSize As Integer, _
                                                  RequestOptns As IP_OPTION_INFORMATION, _
                                                  ReplyBuffer As IP_ECHO_REPLY, _
                                                  ByVal ReplySize As Long, _
                                                  ByVal timeout As Long) As Boolean
                                                  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The CopyMemory function copies a block of memory from one location to another.

'Parameters
'Destination
'   [in] Pointer to the starting address of the copied block's destination.
'Source
'   [in] Pointer to the starting address of the block of memory to copy.
'Length
'   [in] Specifies the size, in bytes, of the block of memory to copy.

'Return Values
'   This function has no return value.

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                                     hpvSource As Any, _
                                                                     ByVal cbCopy As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



























'********************************************************************************
'*
'* Tracert determines the route taken to a destination by sending ICMP
'* echo packets with varying time-to-live (TTL) values to the destination.
'* Before forwarding a packet, each router along the path is required to
'* decrement the TTL value on a packet by at least 1, so the TTL value is
'* effectively a hop count. When the TTL value on a packet reaches 0, the
'* router sends back an ICMP "Time Exceeded" message to the source computer.
'* Tracert determines the route by sending the first echo packet with a TTL
'* value of 1 and incrementing the TTL value by 1 on each subsequent
'* transmission until the target responds, or the maximum TTL value is reached.
'* The route is determined by examining the ICMP "Time Exceeded" messages sent
'* back by intermediate routers. Some routers silently drop packets with expired
'* TTL values and are invisible to Tracert.
'*
'********************************************************************************

Public Sub getTraceRT(Optional strHost As String = vbNullString, _
                      Optional maxTTL As Long = 255)
    Dim ip                                        As String
    Dim RespondingHost As String
    
    ' if the winsock was unavailable then exit the sub
    If boolError Then Exit Sub
    
    ' we are doing a traceroute
    TraceRT = True
    
    ' if no ip/host provided use ours
    If Len(strHost) = 0 Then
        ' gets our hostname
        strHost = GetLocalIPAddress
    End If
    
    ' if no ip/host provided use ours
    ip = vbGetHostByName(strHost)  ' Get the IPAddress For the Target
    
    ' Get ICMP Handle
    If Not vbIcmpCreateFile Then
        Exit Sub
    End If
    
    ' Who we tracing the route to?
    RaiseEvent TracertStatus("Tracing Route To " & ip)
    
    Dim TTL As Long
    hopnumber = 1
    For TTL = 1 To maxTTL
        DoEvents
        ' Send the echo
        RespondingHost = vbIcmpSendEcho(TTL, 1)
        ' We reached the host ip were done
        If RespondingHost = ip Then
            RaiseEvent TracertStatus("Route Trace has Completed")
            Exit For
        End If
    Next TTL
    
    ' if we didnt reach the host then the max ttl was too low
    If Not RespondingHost = ip Then
        RaiseEvent TracertStatus("Reached maximum Time To Live before reaching the destination.")
    End If
    
    ' Close the ICMP Handle
    vbIcmpCloseHandle

End Sub

'***************************************************************************
'*
'* Ping verifies connections to remote computers. It sends Internet Control
'* Message Protocol (ICMP) echo packets to a computer and listens for echo
'* reply packets. Ping waits for up to 1 second for each packet sent, and
'* prints the number of packets transmitted and received to the console.
'* This tool is available only if you install TCP/IP.
'*
'***************************************************************************

Public Sub getPing(Optional strHost As String = vbNullString, _
                   Optional lngTimeToLive As Long = 255, _
                   Optional packets As Long = 1)
    ' if the winsock was unavailable then exit the sub
    If boolError Then Exit Sub
    
    ' were not doing a traceroute
    TraceRT = False
    
    ' if no ip/host provided use ours
    If Len(strHost) = 0 Then
        ' gets our hostname
        strHost = GetLocalIPAddress
    End If
    
    ' Get the IPAddress For the Target
    Call vbGetHostByName(strHost)
    
    ' Get ICMP Handle
    If Not vbIcmpCreateFile Then
        Exit Sub
    End If
    
    ' Send the ICMP Echo Request
    Call vbIcmpSendEcho(lngTimeToLive, packets)
    
    ' Close the ICMP Handle
    vbIcmpCloseHandle
End Sub






'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'*******************************************************************
'*
'* All internal functions/subs below this point
'*
'*******************************************************************

Private Function HiByte(ByVal wParam As Integer) As String

    HiByte = wParam \ &H100 And &HFF&

End Function

Private Function LoByte(ByVal wParam As Integer) As String

    LoByte = wParam And &HFF&

End Function

Private Sub vbIcmpCloseHandle()

    Call IcmpCloseHandle(hIP)

End Sub

Private Function vbIcmpCreateFile() As Boolean
    vbIcmpCreateFile = True
    hIP = IcmpCreateFile()
    If hIP = INVALID_HANDLE_VALUE Then
        vbIcmpCreateFile = False
        RaiseEvent Error("ICMP Error: " & Err.LastDllError)
    End If
End Function

'*******************************************************************
'* The ICMPSendEcho() function sends an ICMP echo request to the
'* specified destination IP address and returns any replies received
'* within the timeout specified. The API is synchronous, requiring
'* the process to spawn a thread before calling the API to avoid
'* blocking. An open IcmpHandle is required for the request to
'* complete. IcmpCreateFile() and IcmpCloseHandle() functions are
'* used to create and destroy the context handle.
'*******************************************************************

Private Function vbIcmpSendEcho(TTL As Long, packets As Long) As String
Dim szBuffer As String
Dim NbrOfPkts As Integer
Dim bReturn As Boolean
Dim pIPe As IP_ECHO_REPLY
Dim pIPo As IP_OPTION_INFORMATION
Dim RespondingHost As String
Static TraceFail As Integer

    pIPo.TTL = TTL

    ' i fully intend to implement different packet sizes i just havent got
    ' round to it.... yet.
    szBuffer = "abcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklm"
    szBuffer = Left$(szBuffer, Val(32))
    ' on one set of source code i tried they had this set to For NbrOfPkts = 1 To trim$(1)
    ' really would i lie to you :)
    For NbrOfPkts = 1 To packets
        DoEvents
        bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 0)
        If bReturn Then
            RespondingHost = CStr(pIPe.Address(0)) & "." & CStr(pIPe.Address(1)) & "." & CStr(pIPe.Address(2)) & "." & CStr(pIPe.Address(3))
            Call GetRCode(RespondingHost, pIPe)
            vbIcmpSendEcho = RespondingHost
        Else 'BRETURN = FALSE/0
            If TraceRT Then
                ' the reason i've changed this is simply to take into account
                ' ttls that wont EVER respond. You will end up in a infinite
                ' loop. Try out some of the other icmp code on psc to see what
                ' i mean.
                ' ------------------------------------------------------------
                ' Anyway basicly it allows 5 retries then it skips this ttl
                TraceFail = TraceFail + 1
                If Not TraceFail = 6 Then
                    ' prevent the ttl from increasing on this occasion lets
                    ' try again.
                    TTL = TTL - 1
                Else
                    TraceFail = 0
                End If
            Else 'TRACERT = FALSE/0
                ' ping status
                RaiseEvent PingStatus("Request Timeout", 0)
            End If
        End If
    Next NbrOfPkts

End Function

'*******************************************************************
'* This simply (or not so simply!) starts the winsock system.
'* Note that the Winsock 1.1 WSAStartup function must be called prior
'* to using the functions exposed by ICMP.DLL. If you do not do this,
'* the first call to IcmpSendEcho will fail with error 10091 (WSASYSNOTREADY).
'*******************************************************************

Private Function vbWSAStartup() As Boolean

Dim sMsg    As String
Dim i       As Long
Dim iReturn As Long
Dim WSAdata As WSAdata

    vbWSAStartup = False
    iReturn = WSAStartup(&H101, WSAdata)
    
    ' If WSock32 error, then tell me about it
    If iReturn <> 0 Then
        RaiseEvent Error("WSock32.dll is not responding")
    End If
    
    ' If the winsock version is incorrect tell the caller.
    With WSAdata
        If LoByte(.wVersion) < WS_VERSION_MAJOR Or (LoByte(.wVersion) = WS_VERSION_MAJOR And HiByte(.wVersion) < WS_VERSION_MINOR) Then
            sMsg = "WinSock Version " & Trim$(Str$(LoByte(.wVersion))) & "." & Trim$(Str$(HiByte(.wVersion)))
            sMsg = sMsg & " is Not supported "
            RaiseEvent Error(sMsg)
            Exit Function
        End If
    End With
    
    ' if the maximum sockets allowed is too low then let the caller know.
    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        RaiseEvent Error(sMsg)
        Exit Function
    End If
    
    Dim MaxSockets As Long
    MaxSockets = WSAdata.iMaxSockets
    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    
    Dim MaxUDP As Long
    MaxUDP = WSAdata.iMaxUdpDg
    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Dim Description As String
    Description = ""
    For i = 0 To WSA_DESCRIPTION_LEN
        If WSAdata.szDescription(i) = 0 Then
            Exit For
        End If
        Description = Description + Chr$(WSAdata.szDescription(i))
    Next i
    
    Dim Status As String
    Status = ""
    For i = 0 To WSA_SYS_STATUS_LEN
        If WSAdata.szSystemStatus(i) = 0 Then
            Exit For
        End If
        Status = Status + Chr$(WSAdata.szSystemStatus(i))
    Next i
        
    RaiseEvent SockStartup(MaxSockets, MaxUDP, Description, Status)
    
    vbWSAStartup = True

End Function

' Get this machines local ip address
Private Function GetLocalIPAddress() As String

Dim sMsg    As String
Dim strHost    As String

Dim HostLen As Long
    strHost = String$(64, &H0)
    If gethostname(strHost, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        RaiseEvent Error(sMsg)
    Else 'NOT GETHOSTNAME(HOST,...
        GetLocalIPAddress = Left$(Trim$(strHost), Len(Trim$(strHost)) - 1)
    End If

End Function

' get the targets hostname note this also assigns the Addr variable for use by icmp
Private Function vbGetHostByName(ByVal host As String) As String
    Dim szString As String
    Dim HostEnt                   As HostEnt
    Dim PointerToPointer          As Long
    Dim ListAddress               As Long
    Dim ListAddr                                  As Long
    
    host = Trim$(host)
    szString = String(64, &H0)
    host = host + Right$(szString, 64 - Len(host))

    If gethostbyname(host) = SOCKET_ERROR Then
        Dim sMsg As String
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        RaiseEvent Error(sMsg)
    Else
        PointerToPointer = gethostbyname(host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HostEnt.h_name, ByVal _
        PointerToPointer, Len(HostEnt) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HostEnt.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        vbGetHostByName = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
        + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
    End If
End Function

' Calculate the programs responce to the icmp reply
Private Sub GetRCode(RespondingHost As String, pIPe As IP_ECHO_REPLY)
    Dim RCode As String
    
    ' free up other events
    DoEvents
    
    ' The string result code
    RCode = strRcode(pIPe)
    
    ' Are we doing a traceroute?
    If Not TraceRT Then
        ' if the status is zero then send back all the data for this ping
        If pIPe.Status = 0 Then
            RaiseEvent PingResponce(RespondingHost, Trim$(CStr(pIPe.DataSize)), Trim$(CStr(pIPe.Options.TTL)), Trim$(CStr(pIPe.RoundTripTime)))
        Else 'NOT PIPE.STATUS...
            ' if the status not zero then resturn the responce code error
            RaiseEvent PingStatus(RespondingHost, RCode)
        End If
    Else ' is a tracert
        ' note you could set this to check to make sure the reply is in fact a time
        ' out. Then handle any other error codes returned normally.... Something
        ' like the commented out code perhaps. Just a thought :)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If pIPe.Status = 11013 Then
        ' if were doing a trace route then
        If hopnumber < 10 Then
            ' add preceading zero if its a single digit number
            RaiseEvent TracertResponce(" # 0" & CStr(hopnumber), RespondingHost, CLng(pIPe.Options.TTL))
        Else 'NOT hopnumber < 10...
            ' just give the hop number
            RaiseEvent TracertResponce(" # " & CStr(hopnumber), RespondingHost, CLng(pIPe.Options.TTL))
        End If
        hopnumber = hopnumber + 1
        'Else
        '    MsgBox RCode
        'End If
    End If

End Sub

Private Function strRcode(pIPe As IP_ECHO_REPLY) As String
' checkout "Internet Control Message Protocol (ICMP) Basics" - "KB: Windows" for
' more details on some responces you can get from icmp.
    Dim RCode As String
    Select Case pIPe.Status
    Case 0
        RCode = "Success"
    Case 11001
        RCode = "Buffer too Small"
    Case 11002
        RCode = "Dest Network Not Reachable"
    Case 11003
        RCode = "Dest Host Not Reachable"
    Case 11004
        RCode = "Dest Protocol Not Reachable"
    Case 11005
        RCode = "Dest Port Not Reachable"
    Case 11006
        RCode = "No Resources Available"
    Case 11007
        RCode = "Bad Option"
    Case 11008
        RCode = "Hardware Error"
    Case 11009
        RCode = "Packet too Big"
    Case 11010
        RCode = "Rqst Timed Out"
    Case 11011
        RCode = "Bad Request"
    Case 11012
        RCode = "Bad Route"
    Case 11013
        RCode = "TTL Exprd in Transit"
    Case 11014
        RCode = "TTL Exprd Reassemb"
    Case 11015
        RCode = "Parameter Problem"
    Case 11016
        RCode = "Source Quench"
    Case 11017
        RCode = "Option too Big"
    Case 11018
        RCode = "Bad Destination"
    Case 11019
        RCode = "Address Deleted"
    Case 11020
        RCode = "Spec MTU Change"
    Case 11021
        RCode = "MTU Change"
    Case 11022
        RCode = "Unload"
    Case 11050
        RCode = "General Failure"
    Case Else
        RCode = "Unknown Failure - " & pIPe.Status
    End Select
    RCode = RCode & " (" & CStr(pIPe.Status) & ")"
    strRcode = RCode
    
End Function

Private Sub Class_Initialize()
    ' initalise the winsock dll or die
    boolError = False
    If Not vbWSAStartup Then
        boolError = True
        Exit Sub
    End If
End Sub

Private Sub Class_Terminate()
    ' Close Winsock dll
    If Not boolError Then
        WSACleanup
    End If
End Sub

