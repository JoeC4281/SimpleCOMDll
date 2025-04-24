VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'set CScript = New Console
'CScript.Echo "This is VBScript"
Private fso
Private stdout
Private stderr

Private Sub Class_Initialize()
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set stdout = fso.GetStandardStream(1)
  Set stderr = fso.GetStandardStream(2)
End Sub

Public Sub Echo(theArg)
  stdout.WriteLine theArg
End Sub

Private Sub class_Terminate()
  Set stderr = Nothing
  Set stdout = Nothing
  Set fso = Nothing
End Sub
