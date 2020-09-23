<div align="center">

## View/Change TCP Stack WITHOUT calling "netstat \-o"


</div>

### Description

A few of the firewalls written here on PSC shell to netstat.exe to map ports with processes. This article explains how to do the same thing using a couple undocumented Windows XP API calls.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-smith.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-smith-view-change-tcp-stack-without-calling-netstat-o__1-49227/archive/master.zip)





### Source Code

<p><b>Firewalls...</b> They're a great tool that can be used to keep a system from being murdered while on a non-trustable network, ie the Internet. They are, however, somewhat of a mystery (and a pain) to write.
<p><b>A Dilemma...</b> By using the API functions GetTcpTable and SetTcpEntry, a program can retrieve the active connections and close them if needed. This is fine if all you want to do is restrict port access, but what if you want to restrict program access?
<p>The poor-man's way of doing it is to simply shell netstat.exe -o and process the output. But this method is slow, and not always fool-proof.
<p>After doing some research, I came across the function AllocateAndGetTcpTableExFromStack; an undocumented function in iphlpapi.dll. After trying to figure out how to use it for a few weeks and my results coming up empty, my question was answered.
<p>My code can be viewed below. Note that this is a <b>class module</b>
<p><pre>Option Explicit
Private Type MIB_TCPROW
  dwState As Long
  dwLocalAddr As Long
  dwLocalPort As Long
  dwRemoteAddr As Long
  dwRemotePort As Long
End Type
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long
Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private pTablePtr As Long
Private pDataRef As Long
Private nRows As Long
Private nCurrentRow As Long
Private udtRow As MIB_TCPROW
Private nState As Long
Private nLocalAddr As Long
Private nLocalPort As Long
Private nRemoteAddr As Long
Private nRemotePort As Long
Private nProcId As Long
Public Event StackEntry(ByVal StackIndex As Long, ByVal LocalAddr As Long, ByVal LocalPort As Long, ByVal RemoteAddr As Long, ByVal RemotePort As Long, ByVal ProcessId As Long, ByVal State As Long)
Public Function GetIPAddress(dwAddr As Long) As String
  Dim arrIpParts(3) As Byte
  CopyMemory arrIpParts(0), dwAddr, 4
  GetIPAddress = CStr(arrIpParts(0)) & "." & _
          CStr(arrIpParts(1)) & "." & _
          CStr(arrIpParts(2)) & "." & _
          CStr(arrIpParts(3))
End Function
Public Function GetPort(ByVal dwPort As Long) As Long
  GetPort = htons(dwPort)
End Function
Public Function RefreshStack() As Boolean
  Dim nRet As Long
  pDataRef = 0
  nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)
  If nRet = 0 Then
    CopyMemory nRows, ByVal pTablePtr, 4
    RefreshStack = True
  Else
    RefreshStack = False
  End If
End Function
Public Function GetEntryCount() As Long
  GetEntryCount = nRows - 2  '// The last entry is always an EOF of sorts
End Function
Public Function EnumEntries() As Boolean
  Dim i As Long  '// loop counter
  'On Error Resume Next
  EnumEntries = True
  If nRows = 0 Or pTablePtr = 0 Then
    EnumEntries = False
    Exit Function
  End If
  For i = 0 To nRows '// read 24 bytes at a time
    CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
    CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
    CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
    CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
    CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
    CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4
    DoEvents
    If nRemoteAddr <> 0 Or nRemotePort <> 0 Or nLocalPort <> 0 Then
      RaiseEvent StackEntry(i, nLocalAddr, nLocalPort, nRemoteAddr, nRemotePort, nProcId, nState)
    End If
    pDataRef = pDataRef + 24
    DoEvents
  Next i
  'pDataRef = 0
End Function
Public Sub TerminateThisConnection()
  udtRow.dwLocalAddr = nLocalAddr
  udtRow.dwLocalPort = nLocalPort
  udtRow.dwRemoteAddr = nRemoteAddr
  udtRow.dwRemotePort = nRemotePort
  udtRow.dwState = 12
  SetTcpEntry udtRow
End Sub</pre>
<p>If you have any questions as to how this code works, please feel free to ask.

