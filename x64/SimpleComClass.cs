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
		double PPP(double PPK); // Adding PPP function to the interface
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
        
        public double PPP(double PPK)
        {
            return PPK * 0.454; // Implementing PPP function
        }
    }
}
