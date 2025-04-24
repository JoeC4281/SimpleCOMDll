VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  'This program requires the vbAdvance plugin,
  '  which allows the creation of a console .exe
  '
  'You can also use editbin.exe to create a console .exe
  '  after you have created comtest.exe
  '
  '  editbin.exe /SUBSYSTEM:CONSOLE comtest.exe
  '
  'editbin.exe is part of Visual Studio
  '
  'Build your own Application Mode Changer with VB6...
  '  https://www.nirsoft.net/vb/appmodechange.html
  '
  Dim cs As SimpleComClass
  Dim CScript As New Console
  
  Set cs = New SimpleComClass
  Set CScript = New Console
  
  If App.LogMode = 1 Then
    CScript.Echo cs.PPP(6.59)
  End If
  
  If App.LogMode = 0 Then
    Debug.Print cs.PPP(6.59)
  End If
  
  Set cs = Nothing
  
  Unload Me
End Sub
