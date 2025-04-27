Dim obj
Dim CScript

Set obj = CreateObject("SimpleComDll.SimpleComClass")
Set CScript = New Console

obj.ShowMessage "Hello from COM DLL!"

' Test PPP function
Dim inputValue, resultValue
inputValue = 6.59 ' Example input value for PPP function
resultValue = obj.PPP(inputValue)

' Display the result
CScript.Echo "PPP(" & inputValue & ") = " & resultValue

Set CScript = nothing
Set obj = Nothing

'set CScript = New Console
'CScript.Echo "This is VBScript"
Class Console
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
End Class