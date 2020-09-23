VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadAcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' This code can be written as Private in a Form
'Since I have a plan to reuse this class, therefore I preffered to
'create clsLoadacs class.
Public Sub Load_genie()
If strOS9598 = True Then
strchar = "C:\WINDOWS\MSAGENT\CHARS\genie.acs"
ElseIf strOSNT20 = True Then
strchar = "C:\WINNT\MSAGENT\CHARS\genie.acs"
Else
MsgBox "genie is notloaded"
End If
End Sub
