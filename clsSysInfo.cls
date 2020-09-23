VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public strOS As String
Public Sub OSSys_Type()
Dim typOS As OSVERSIONINFO
Dim lngword As Long

typOS.dwOSVersionInfoSize = 148
lngword = GetVersionEx(typOS)
If typOS.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
strOS9598 = True
strOS = " Windows 95/98/Me  "
ElseIf typOS.dwPlatformId = VER_PLATFORM_WIN32_NT Then
strOS = " Windows NT/2000 "
strOSNT20 = True
End If
strOS = strOS & "  Version" & typOS.dwMajorVersion & "/" & typOS.dwMinorVersion
End Sub

