VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LanRunner Mini V1.0-The UltraFast Network Scanner"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scan Network"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.ListBox lstping 
      Height          =   4740
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ListBox lstadr 
      Height          =   4740
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load IP"
      Height          =   400
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtLast 
      Height          =   405
      Left            =   2040
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   480
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cmbInterface 
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Text            =   "Interface List"
      Top             =   840
      Width           =   3840
   End
   Begin VB.Label Label4 
      Caption         =   "Online IP'S"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "IP'S to scan "
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "To IP"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From IP"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ProtocolBuilder         As clsProtocolInterface

Private WithEvents TCPDriver    As clsTCPProtocol
Attribute TCPDriver.VB_VarHelpID = -1
Private IPHEADER1 As clsIPHeader
'The complete number of bytes sent including those that make up the headers
Private BytesRecievedPackets    As Long
Dim i2 As Integer
'The number of bytes of data sent (i.e. exlcuding the packet headers)
Private BytesRecieved           As Long

'The number of packets recieved for each protocol
Private TCPPackets              As Long
Private TCPLog                  As Integer
Const IP_0_0_0_1 = 16777216
Const IP_0_0_1_0 = 65536
Const IP_0_1_0_0 = 256
Const IP_1_0_0_0 = 1
'API Declarations

Dim m(255) As String
Dim N(255) As String
Dim o(255) As String

Dim ret As Long
Private Declare Function inet_addr Lib "wsock32.dll" _
  (ByVal s As String) As Long

Private Declare Function SendARP Lib "iphlpapi.dll" _
  (ByVal DestIP As Long, _
   ByVal SrcIP As Long, _
   pMacAddr As Long, _
   PhyAddrLen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, _
   Src As Any, _
   ByVal bcount As Long)
   Private Const NO_ERROR = 0
   
   
   Const MAXLEN_PHYSADDR = 8
Private Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr As Long
    dwType As Long
End Type
Private Declare Function GetIpNetTable Lib "IPHlpApi" (pIpNetTable As Byte, pdwSize As Long, ByVal bOrder As Long) As Long
Private Const WSADESCRIPTION_LEN As Long = 256
Private Const WSASYS_STATUS_LEN As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2

Private Type WSAdata
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADESCRIPTION_LEN) As Byte
  szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
  iMaxSockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSAdata) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long



Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim WithEvents objICMP As clsICMP
Attribute objICMP.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim i As Integer
Dim iping As Integer

ProtocolBuilder.CloseRawSocket
ProtocolBuilder.CreateRawSocket Left$(cmbInterface.Text, InStr(1, cmbInterface, " ")), 7000, Me.hWnd


lstadr.Clear
lstping.Clear
    If IsAddress(txtIP) Then
     
        
        
      
            
            If Trim(txtLast) = "" Then 'Trim is used to "clean" adjacent spaces
            'User forgot to insert an address let's just ignore the interval:
                If Not EntryExists(txtIP, lstadr) Then
                    'Single address entry:
                    'There is no repeated entry, ok to proceed:
                    lstadr.AddItem txtIP
                    'Check if Rem button is enabled
                   ' If cmdDel.Enabled = False Then cmdDel.Enabled = True
                End If
            
            ElseIf IsAddress(txtLast) Then 'User has inserted some value
                'Interval of Adresses search routine:
                Dim FirstAddr As Double, LastAddr As Double, IpAddr As String, CurAddr As Double
                'First and last address converted to long so they can be calculated later:
                FirstAddr = AddrToLong(txtIP)
                LastAddr = AddrToLong(txtLast)
                'Obviously first addr must be smaller
                If FirstAddr < LastAddr Then  'OK to proceed
                    For CurAddr = FirstAddr To LastAddr
                        'From first address to the last:
                        'Convert it to IP - string - so we can show it in list
                        IpAddr = LongToAddr(CurAddr)
                        If Not EntryExists(IpAddr, lstadr) Then
                            'Assure there are no duplicates
                            lstadr.AddItem IpAddr
                        End If
                        'we don't wan't to make our app to "freeze"
                        DoEvents
                    Next CurAddr
                End If
                'at least one entry was made, so let's check for Remove button:
               ' If cmdDel.Enabled = False Then cmdDel.Enabled = True
            End If
     End If

    
DoEvents
 

For i = 0 To lstadr.ListCount - 1

With IPHEADER1
.Checksum = 32602
.DestAddress = inet_addr(lstadr.List(i))
.DestIP = CStr(lstadr.List(i))
.HeaderLength = 20
.ID = 15656
.Offset = 0
.PacketLength = 60
.Protocol = IPPROTO_ICMP
.SourceAddress = inet_addr("10.22.165.9")
.SourceIP = "10.22.165.9"
.TimeToLive = 128
.TypeOfService = 0
.Version = 4


End With
TCPDriver.SendPacket CStr(lstadr.List(i)), 80, IPHEADER1, "Hello World"
DoEvents
   
Next i


End Sub

Private Sub Command2_Click()
Call arpread
End Sub

Private Sub Command3_Click()
About.Show
End Sub

Private Sub Form_Load()
Dim str() As String, i As Integer
  Dim mcnt1, mcnt2, mcnt3 As Integer
    
    
    
   
    Set objICMP = New clsICMP
    Set IPHEADER1 = New clsIPHeader
     Set ProtocolBuilder = New clsProtocolInterface
    Set TCPDriver = New clsTCPProtocol
    
    ProtocolBuilder.AddinProtocol TCPDriver, "TCP", IPPROTO_TCP
   

    str = Split(EnumNetworkInterfaces(), ";")
        
    For i = 0 To UBound(str)
        If str(i) <> "127.0.0.1" Then
            cmbInterface.AddItem str(i) & " [" & GetHostNameByAddr(inet_addr(str(i))) & "]"
        End If
    Next
    
    cmbInterface.Text = cmbInterface.List(0)

  mcnt1 = InStr(1, cmbInterface.Text, ".")
  mcnt2 = InStr(mcnt1 + 1, cmbInterface.Text, ".")
 
  mcnt3 = InStr(mcnt2 + 1, cmbInterface.Text, ".")
  
  txtIP.Text = Mid(cmbInterface.Text, 1, mcnt3) & "1"
  txtLast.Text = Mid(cmbInterface.Text, 1, mcnt3) & "255"
End Sub
Function EntryExists(Entry As String, List As ListBox) As Boolean
    With List
        'Simple loop in matching each value of each list value with the disired
        Dim i As Integer, Last As Integer
        EntryExists = False
        Last = .ListCount - 1
        If Last < 0 Then Exit Function
        For i = 0 To Last
            If .List(i) = Entry Then
                EntryExists = True
                Exit Function
            End If
        Next i
    End With
End Function
Private Sub arpread()
    'KPD-Team 2001
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
   

     Dim mRow As ListItem
    Dim Listing() As MIB_IPNETROW, cnt As Long
    Dim bBytes() As Byte, bTemp(0 To 3) As Byte
    Dim fip As String
    lstping.Clear
    'lv.ListItems.Clear
    'set the graphics mode of this form to persistent
    Me.AutoRedraw = True
    'call the function to retrieve how many bytes are needed
    GetIpNetTable ByVal 0&, ret, False
    DoEvents
    'if it failed, exit the sub
    If ret <= 0 Then Exit Sub
    'redimension our buffer
    ReDim bBytes(0 To ret - 1) As Byte
    'retireve the data
    GetIpNetTable bBytes(0), ret, False
    'copy the number of entries to the 'Ret' variable
    CopyMemory ret, bBytes(0), 4
    'redimension the Listing
   
        If ret > 0 Then ReDim Listing(0 To ret - 1) As MIB_IPNETROW
        For cnt = 0 To ret - 1
        CopyMemory Listing(cnt), bBytes(4 + 24 * cnt), 24
        CopyMemory bTemp(0), Listing(cnt).dwAddr, 4
         fip = ConvertAddressToString(bTemp(), 4)
         
         If Asc(Mid(CStr(Listing(cnt).bPhysAddr), 1, 1)) = 63 Then
         m(cnt) = fip
         lstping.AddItem fip
         N(cnt) = saddr
         End If
         Next cnt
         ret = 0
End Sub
Public Function ConvertAddressToString(bArray() As Byte, lLength As Long) As String
    Dim cnt As Long
    For cnt = 0 To lLength - 1
        ConvertAddressToString = ConvertAddressToString + CStr(bArray(cnt)) + "."
    Next cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function
