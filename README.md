# SimpleCOMDll
A simple C# ActiveX COM dll 

This is a method of creating a 32-bit ActiveX COM dll from the command line,  
without the need to use Visual Studio.

This 32-bit ActiveX COM dll can then be used from;

Powershell  
thinBasic  
VBScript  
Visual Basic 6.0  
Visual FoxPro 9.0

...and other 32-bit applications.

Review the source code for 32bit.btm for detailed instructions.

32bit.btm is a TCC batch file

TCC is available from https://jpsoft.com/

Here's a thinBASIC example;
```vb script
Uses "Console"

dim Cs as iDispatch
dim result as String

CS = CreateObject("SimpleComDll.SimpleComClass")

If IsComObject(cs) then
  result = cs.PPP(6.59)

  printl result
  
Else
  Printl "Could not create SimpleComDll.SimpleComClass object"
end if

cs = Nothing
```
