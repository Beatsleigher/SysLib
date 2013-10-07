''' <License>
''' This program is licensed under the LGPL 3.0. For more information, please visit:
''' http://gnu.org
''' </License>
''' 
''' <summary>
''' Do not forget to add SysLib as a reference to your project!
''' </summary>
''' <remarks></remarks>
Imports SysLib.GetInfo

Module Samples
    Private cpu As CPU = New CPU
    Private cd As CDROMDrive = New CDROMDrive
    Private cc As ComputerSystem = New ComputerSystem

    Sub Main()
        Console.WriteLine(cpu.GetAddressWidth)
        Console.WriteLine(cpu.GetCaption)
        Console.WriteLine(cd.IsMediaLoaded)
        Console.WriteLine(cc.GetAdminPasswordStatus)
        Console.WriteLine("Display more info? (Y/N)")
        Dim read As String = Console.ReadLine()
        If read = "y" Or read = "Y" Then
            Console.WriteLine(cpu.ShowInfo_Console)
        Else
            Console.WriteLine("No more info will be displayed.")
        End If
        Console.WriteLine("Hit ENTER to continue...")
        Console.ReadLine()
    End Sub
End Module
