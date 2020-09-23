Attribute VB_Name = "Modvoice"

'  To load Genie
Public strchar As String '
Public strGeniePath As String
'To configure operating system
Public strOS9598 As Boolean
Public strOSNT20 As Boolean
'
Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public strOS As String
