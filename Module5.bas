Attribute VB_Name = "Module5"
'How to detect if another application has established
'the connection to Internet.
'Run Only in WIN95,98 (not tested in NT)
'If the connection is succesfull, the function
'IsConnected return True
'The Burn Poodle

Option Explicit
Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
'
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32
'
Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
'
Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
        


Public Function IsConnected() As Boolean
On Error Resume Next
Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95
'
TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize
'
RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
If RetVal <> 0 Then
                    MsgBox "ERROR"
                    Exit Function
                    End If
'
Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
If Tstatus.RasConnState = &H2000 Then
                         IsConnected = True
                         Else
                         IsConnected = False
                         End If

End Function

