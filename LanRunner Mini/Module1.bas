Attribute VB_Name = "Module1"
Option Explicit
'API Declares:
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
            "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
            ByVal lpKeyName As Any, ByVal lpDefault As String, _
            ByVal lpReturnedString As String, ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
            "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
            ByVal lpKeyName As Any, ByVal lpString As Any, _
            ByVal lpFileName As String) As Long


'Rects:
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Mouse:
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPIa) As Long
'Time windows is running:
Public Declare Function GetTickCount Lib "kernel32" () As Long
'Winsock:
'Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'These are used with Winsock and ICMP: -personaly i don't know how to use them, or what are they for ;)
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Systray:
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'INI:
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Windows general:
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'ICMP:
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal timeout As Long) As Boolean

'Types tacken from Winsock.BAS:
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

'Types for systray:
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Types for a rectangle (used in graphs)
'Public Type RECT
'    left As Long
'    tOp As Long
'    Right As Long
'    Bottom As Long
'End Type

'Type for mouse pointer:
Public Type POINTAPIa
    x As Long
    y As Long
End Type

'Wsock consts:
Const SOCKET_ERROR = 0

'Tray consts:
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

'Mouse consts:
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_RBUTTONUP = &H205

'Listbox consts:
Const LB_SETHORIZONTALEXTENT = &H194 'Barra Horizontal
Const LB_ITEMFROMPOINT = &H1A9

'Window consts:
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const HWND_TOP = 0

Public Function ObterINI(Secção As String, Chave As String, Ficheiro As String) As String
    'Get values from INI
    On Error Resume Next
    Dim sRet As String 'Texto a obter
    
    sRet = String(255, Chr(0)) 'Tamanho do string
    ObterINI = Left(sRet, GetPrivateProfileString(Secção, ByVal Chave, "", sRet, Len(sRet), Ficheiro))
End Function

Public Function DefinirINI(Secção_Definir As String, Chave_Definir As String, Valor_Definir As String, Ficheiro_Definir) As Integer
    'Write INI:
    On Error Resume Next
    'Escrever no INI
    Dim r
    r = WritePrivateProfileString(Secção_Definir, Chave_Definir, Valor_Definir, Ficheiro_Definir)
End Function

Public Sub MostraTray(Formulario As Form, ToolTipText As String, TrayIcon As NOTIFYICONDATA)
    'Show icon (form's Icon)
    TrayIcon.cbSize = Len(TrayIcon) ' tamanho..
    'definir a handle da janela (geralmente o formulário)
    TrayIcon.hwnd = Formulario.hwnd
    'Identificação para o icone na barra de tarefas
    TrayIcon.uId = 1&
    'Definir as flags
    TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Definir a callback
    TrayIcon.ucallbackMessage = WM_LBUTTONDOWN
    'Definir o icone (.ico)
    TrayIcon.hIcon = Formulario.Icon
    'Definir a tooltip (dica) - formatada...
    TrayIcon.szTip = ToolTipText & Chr$(0)
    'criar o icone (Add...)
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

Public Sub ApagaTray(Formulario As Form, TrayIcon As NOTIFYICONDATA)
    'Clean tray icon
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = Formulario.hwnd
    TrayIcon.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

Public Function LB_CordParaItem(LB As ListBox, XTwips As Single, YTwips As Single) As Long
    'Get the item number of where mouse is pointing(in a listbox)
    Dim x As Long, y As Long
    x = CLng(XTwips / Screen.TwipsPerPixelX)
    y = CLng(YTwips / Screen.TwipsPerPixelY)
    LB_CordParaItem = SendMessage(LB.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((y * 65536) + x))
End Function

Public Sub DefenirJanelaTopo(Formulario As Form)
    'Set window on top
    SetWindowPos Formulario.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub RetirarJanelaTopo(Formulario As Form)
    'UnSet window on top
    SetWindowPos Formulario.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function PingHostByAdress(HostDotNumber As String) As Long
    'Credits:
    'This function was taken (but altered) from AllAPI
    Dim hFile As Long ', lpWSAdata As WSADataType
    Dim hHostent As HostEnt, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Address = inet_addr(HostDotNumber)
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        PingHostByAdress = -1 '"Unable to Create File Handle"
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        PingHostByAdress = -2 '"Timeout"
    End If
    If EchoReply.Status = 0 Then
        PingHostByAdress = EchoReply.RoundTripTime '(Trim$(CStr(EchoReply.RoundTripTime)))
    Else
        PingHostByAdress = -3 '"Failure ..."
    End If
    Call IcmpCloseHandle(hFile)
End Function
Function getascip(ByVal inn As Long) As String
    'Credits:
    'This function was taken from WinsockAPI.BAS
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function

Public Function AddrToLong(Address As String) As Double
    Dim Addr() As String, i As Double
    If IsAddress(Address) Then
        Addr = Split(Address, ".")
        i = CInt(Addr(0)) * 2 ^ 24
        i = i + CInt(Addr(1)) * 2 ^ 16
        i = i + CInt(Addr(2)) * 2 ^ 8
        i = i + CInt(Addr(3))
        AddrToLong = i
    End If
End Function

Public Function LongToAddr(Address As Double) As String
    Dim Quad(3) As Double, i As Integer, temp As String
    Quad(0) = FstQuad(Address)
    Quad(1) = SndQuad(Address)
    Quad(2) = TrdQuad(Address)
    Quad(3) = FthQuad(Address)
    For i = 0 To 3
        If i > 0 Then
            temp = temp & "."
        End If
        temp = temp & Quad(i)
    Next i
    LongToAddr = temp
End Function

Function FstQuad(Addr As Double) As Double
    'Secondary function:
    'returns (numeric value) of first quadret numbers
    'in an IP address
    Dim temp
    temp = Int(Addr / (2 ^ 24))
    FstQuad = temp
End Function

Function SndQuad(Addr As Double) As Double
    'Secondary function:
    'Second quadret from an IP address
    SndQuad = Int((Addr - FstQuad(Addr) * 2 ^ 24) / (2 ^ 16))
End Function

Function TrdQuad(Addr As Double) As Double
    'Secondary function:
    'Third quadret from an IP address
    Dim temp As Double
    'no firs quad fe: 0.102.124.02
    temp = Addr - FstQuad(Addr) * (2 ^ 24)
    'nor 2nd fe: 0.0.124.02
    temp = temp - SndQuad(Addr) * (2 ^ 16)
    'get 3rd fe: 124
    TrdQuad = Int(temp / (2 ^ 8))
End Function

Function FthQuad(Addr As Double) As Double
    'Secondary function:
    'Fourth quadret from an IP address:
    Dim temp As Long
    'no firs quad fe: 0.102.124.02
    temp = Addr - FstQuad(Addr) * (2 ^ 24)
    'nor 2nd fe: 0.0.124.02
    temp = temp - SndQuad(Addr) * (2 ^ 16)
    'nor 3rd fe: 0.0.0.02
    temp = temp - TrdQuad(Addr) * (2 ^ 8)
    'get last fe:  02
    FthQuad = temp
End Function

Public Function IsAddress(Address As String) As Boolean
    Dim SplitedAddress() As String
    IsAddress = False
    SplitedAddress = Split(Address, ".")
    If UBound(SplitedAddress) = 3 Then
        Dim i As Integer, temp() As String
        For i = 0 To 3
            'Check if it's numeric:
            If IsNumeric(SplitedAddress(i)) = False Then Exit Function
            'Check address interval:
            If SplitedAddress(i) > 255 Or SplitedAddress(i) < 0 Then Exit Function
            'Check if interval is Integer:
            If CInt(SplitedAddress(i)) <> SplitedAddress(i) Then Exit Function
        Next i
        IsAddress = True
    End If
End Function


