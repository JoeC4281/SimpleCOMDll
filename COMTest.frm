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
  Dim cs As SimpleComClass
  Dim CScript As Console
  
  Set cs = New SimpleComClass
  Set CScript = New Console
  
  Debug.Print cs.PPP(6.59)
  CScript.Echo cs.PPP(6.59)
  
  Set cs = Nothing
  
  Unload Me
End Sub
