Attribute VB_Name = "SystemPath"
Option Explicit
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public syspath As String
Function WindowsDirectory() As String 'get windows directory
Dim WinPath As String
Dim temp
    WinPath = String(145, Chr(0))
    temp = GetWindowsDirectory(WinPath, 145)
    WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
syspath = WindowsDirectory
End Function
