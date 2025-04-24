Dim obj
Set obj = CreateObject("SimpleComDll.SimpleComClass")
obj.ShowMessage "Hello from COM DLL!"

' Test PPP function
Dim inputValue, resultValue
inputValue = 6.59 ' Example input value for PPP function
resultValue = obj.PPP(inputValue)

' Display the result
WScript.Echo "PPP(" & inputValue & ") = " & resultValue

Set ojb = Nothing
