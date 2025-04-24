Creating a simple ActiveX COM DLL using C# from the command line involves several steps. Here's a concise guide to help you through the process:

1. Create the C# Class Library

First, create a new directory for your project and navigate into it:

mkdir SimpleComDll
cd SimpleComDll


Create a C# class file, for example, SimpleComClass.cs:

using System;
using System.Runtime.InteropServices;

namespace SimpleComDll
{
    [ComVisible(true)]
    [Guid("E7A9A5C7-3B6D-4C6A-9B7A-2B5D5F5D5F5D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISimpleComClass
    {
        void ShowMessage(string message);
    }

    [ComVisible(true)]
    [Guid("D7A9A5C7-3B6D-4C6A-9B7A-2B5D5F5D5F5D")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SimpleComClass : ISimpleComClass
    {
        public void ShowMessage(string message)
        {
            Console.WriteLine(message);
        }
    }
}

2. Compile the DLL

Use the csc (C# Compiler) to compile the class library:

csc /target:library /out:SimpleComDll.dll SimpleComClass.cs

3. Register the DLL

To register the DLL, you need to use regasm (Assembly Registration Tool):

regasm /codebase SimpleComDll.dll

4. Create a Type Library (Optional)

If you need a type library (.tlb file), you can generate it using regasm:

regasm /tlb:SimpleComDll.tlb SimpleComDll.dll

5. Using the COM DLL

You can now use this COM DLL in any COM-compatible environment. For example, you can create a simple VBScript to test it:

Dim obj
Set obj = CreateObject("SimpleComDll.SimpleComClass")
obj.ShowMessage "Hello from COM DLL!"


Save this script as test.vbs and run it to see the message.

Summary
Create the C# class library with COM attributes.
Compile the DLL using csc.
Register the DLL using regasm.
(Optional) Create a type library using regasm.

This process allows you to create and register a simple ActiveX COM DLL using C# from the command line without needing Visual Studio. If you encounter any issues, feel free to ask for further assistance!

To un-register the ActiveX COM .dll;

regasm SimpleComDll.dll /u
