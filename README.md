<div align="center">

## IP Resolver \- IP Address Convert to Hostname


</div>

### Description

This code here is 100% working IP resolver. This is use for finding the hostname by converting those IP numbers into words (Hostname). or words (Hostname) into IP number. EXAMPLE: this IP address "209.75.50.202" will convert to "access-50-202.ixpres.com" or you can use it to convert Hostname to IP address.
 
### More Info
 
Put these in the code module so that you can reuse it again.

(I surf around the net and found this cool code. It is originally create for ActiveX Control (.ocx) but I changed it to use in a code module. I am sure no one like to create an OCX just to do this short procedure. anyway have fun)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jay Tr\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-tr.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jay-tr-ip-resolver-ip-address-convert-to-hostname__1-5371/archive/master.zip)

### API Declarations

```
'PUT THIS IS A MODULE
Option Explicit
'      ||================================||
'      || Remember to use:        ||
'      || WSACleanup in Form_Unload()  ||
'      || IP_Initialize in Form_Load() ||
'      ||================================||
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128
Private Type HOSTENT
 hName As Long
 hAliases As Long
 hAddrType As Integer
 hLength As Integer
 hAddrList As Long
End Type
Private Type WSADATA
 wversion As Integer
 wHighVersion As Integer
 szDescription(0 To WSADescription_Len) As Byte
 szSystemStatus(0 To WSASYS_Status_Len) As Byte
 iMaxSockets As Integer
 iMaxUdpDg As Integer
 lpszVendorInfo As Long
End Type
Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" _
 (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, _
 addrType As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, _
 ByVal cbCopy As Long)
'checks if string is valid IP address
Private Function IsIP(ByVal strIP As String) As Boolean
 On Error Resume Next
 Dim t As String: Dim s As String: Dim i As Integer
 s = strIP
 While InStr(s, ".") <> 0
  t = Left(s, InStr(s, ".") - 1)
  If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then s = Mid(s, InStr(s, ".") + 1) _
   Else Exit Function
  i = i + 1
 Wend
 t = s
 If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(Val(t)))) And _
  Val(t) >= 0 And Val(t) <= 255 And strIP <> "255.255.255.255" And i = 3 Then IsIP = True
 If Err.Number > 0 Then
  MsgBox Err.Description, , Err.Number
  Err.Clear
 End If
End Function
'converts IP address from string to sin_addr
Private Function MakeIP(strIP As String) As Long
 On Error Resume Next
 Dim lIP As Long
 lIP = Left(strIP, InStr(strIP, ".") - 1)
 strIP = Mid(strIP, InStr(strIP, ".") + 1)
 lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256
 strIP = Mid(strIP, InStr(strIP, ".") + 1)
 lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256 * 256
 strIP = Mid(strIP, InStr(strIP, ".") + 1)
 If strIP < 128 Then
  lIP = lIP + strIP * 256 * 256 * 256
 Else
  lIP = lIP + (strIP - 256) * 256 * 256 * 256
 End If
 MakeIP = lIP
 If Err.Number > 0 Then
  MsgBox Err.Description, , Err.Number
  Err.Clear
 End If
End Function
'resolves IP address to host name
Function NameByAddr(strAddr As String) As String
 On Error Resume Next
 Dim nRet As Long
 Dim lIP As Long
 Dim strHost As String * 255: Dim strTemp As String
 Dim hst As HOSTENT
 If IsIP(strAddr) Then
  lIP = MakeIP(strAddr)
  nRet = gethostbyaddr(lIP, 4, 2)
  If nRet <> 0 Then
   RtlMoveMemory hst, nRet, Len(hst)
   RtlMoveMemory ByVal strHost, hst.hName, 255
   strTemp = strHost
   If InStr(strTemp, Chr(10)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
   strTemp = Trim(strTemp)
   NameByAddr = strTemp
  Else
   MsgBox "Host name not found", , "9003"
   Exit Function
  End If
 Else
  MsgBox "Invalid IP address", , "9002"
  Exit Function
 End If
 If Err.Number > 0 Then
  MsgBox Err.Description, , Err.Number
  Err.Clear
 End If
End Function
'resolves host name to IP address
Function AddrByName(ByVal strHost As String)
 On Error Resume Next
 Dim hostent_addr As Long
 Dim hst As HOSTENT
 Dim hostip_addr As Long
 Dim temp_ip_address() As Byte
 Dim i As Integer
 Dim ip_address As String
 If IsIP(strHost) Then
  AddrByName = strHost
  Exit Function
 End If
 hostent_addr = gethostbyname(strHost)
 If hostent_addr = 0 Then
  MsgBox "Can't resolve hst", , "9001"
  Exit Function
 End If
 RtlMoveMemory hst, hostent_addr, LenB(hst)
 RtlMoveMemory hostip_addr, hst.hAddrList, 4
 ReDim temp_ip_address(1 To hst.hLength)
 RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength
 For i = 1 To hst.hLength
  ip_address = ip_address & temp_ip_address(i) & "."
 Next
 ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
 AddrByName = ip_address
 If Err.Number > 0 Then
  MsgBox Err.Description, , Err.Number
  Err.Clear
 End If
End Function
Sub IP_Initialize()
 Dim udtWSAData As WSADATA
 If WSAStartup(257, udtWSAData) Then
  MsgBox Err.Description, , Err.LastDllError
 End If
End Sub
'      ||================================||
'      || Remember to use:        ||
'      || WSACleanup in Form_Unload()  ||
'      || IP_Initialize in Form_Load() ||
'      ||================================||
```


### Source Code

```
'How do you call these Functions?
Option Explicit
Private Sub Command1_Click()
  Text1.Text = NameByAddr(Text2)
End Sub
Private Sub Command2_Click()
  Text2.Text = AddrByName("www.yahoo.com")
End Sub
Private Sub Form_Load()
  IP_Initialize
End Sub
Private Sub Form_Unload(Cancel As Integer)
  WSACleanup
End Sub
```

