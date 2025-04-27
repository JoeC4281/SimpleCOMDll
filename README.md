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

Review the source code for <a href="https://github.com/JoeC4281/SimpleCOMDll/blob/main/32bit.btm" target="_blank">32bit.btm</a> for detailed instructions.

For 64-bit, review the source for <a href="https://github.com/JoeC4281/SimpleCOMDll/blob/main/x64/64bit.btm" target="_blank">64bit.btm</a> for detailed instructions.

32bit.btm and 64bit.btm are TCC batch files.

TCC is available from https://jpsoft.com/

Here's a <a href="https://www.thinbasic.com/" target="_blank">thinBASIC</a> example;
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
