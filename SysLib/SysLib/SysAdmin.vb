Imports System.Diagnostics
Imports System.IO


Namespace SysAdmin

    ''' <summary>
    ''' Provides methods and functions for easy use of shutdown.exe
    ''' 
    ''' Provides Functionaility:
    ''' Shutdown
    ''' Restart
    ''' Hibernate
    ''' Log off
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ShutDown

        ''' <summary>
        ''' Cancels any and all shutdown, hibernate, log off, standby operations.
        ''' </summary>
        ''' <returns>True when complete</returns>
        ''' <remarks></remarks>
        Public Function CancelOperation() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-a"
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
            Return True
        End Function

        ''' <summary>
        ''' Shuts down computer. 
        ''' time declares how much time the user has until the system shuts down in seconds. 
        ''' The default is 60 (Integer/Int32)
        ''' 
        ''' comments declares any comments to issue to the user.
        ''' Comments can be any string type.
        ''' For example: "Your computer will shut down in X seconds. Please save all your work and close all currently running programs."
        ''' </summary>
        ''' <param name="time"></param>
        ''' <param name="comments"></param>
        ''' <remarks></remarks>
        Public Sub Shutdown(time As Int32, comments As String)
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-s -t" & time & " -c""" & comments & """"
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub

        ''' <summary>
        ''' Restarts the computer.
        ''' 
        ''' time declares how much time the user has until the system shut down in seconds.
        ''' The default is 60 (Integer/Int32)
        ''' 
        ''' comments declares any comments to issue to the user.
        ''' Comments can be any string type.
        ''' For example: "Your computer will shut down in X seconds. Please save all your work and close all currently running programs."
        ''' </summary>
        ''' <param name="time"></param>
        ''' <param name="comments"></param>
        ''' <remarks></remarks>
        Public Sub Restart(time As Int32, comments As String)
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-r -t" & time & " -c""" & comments & """"
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub

        ''' <summary>
        ''' Puts computer in hibernation mode.
        ''' time declares how much time the user has until the system hibernates in seconds.
        ''' The default is 60 (Integer/Int32)
        ''' 
        ''' comments declares any comments to be issued to the user.
        ''' Comments can be any string type.
        ''' For example: "Your computer will shut down in X seconds. Please save all your work and close all currently running programs."
        ''' </summary>
        ''' <param name="time"></param>
        ''' <param name="comments"></param>
        ''' <remarks></remarks>
        Public Sub Hibernate(time As Int32, comments As String)
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-h -t" & time & " -c""" & comments & """"
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub

        ''' <summary>
        ''' Shows the shutdown executables GUI.
        ''' Allows user to input arguments and shutdown remote computers.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DisplayShutdownGUI()
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-s -i"
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub

        ''' <summary>
        ''' Shuts down a remote computer connected to the home network/domain.
        ''' Can be used as a mean prank or as an administrative task.
        ''' 
        ''' Disclaimer:
        ''' Beatsleigher, Team M4gkBeatz and/or their affiliates and contributors to this project
        ''' are not responsible for any damage to computers and/or people by using this function.
        ''' Please use with care.
        ''' </summary>
        ''' <param name="computerName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ShutdownRemoteComputer(computerName As String) As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .Arguments = "-s -m \\" & computerName
                .CreateNoWindow = True
                .FileName = "shutdown.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
            Return True
        End Function
    End Class

    ''' <summary>
    ''' Gives easy access to management software pre-installed on the Windows operating system
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Management
        ''' <summary>
        ''' Shows Disk Management software
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DiskManagement() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "Diskmgmt.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
            Return True
        End Function

        ''' <summary>
        ''' Shows Device Management software
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeviceManagement() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "devmgmt.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
            Return True
        End Function

        ''' <summary>
        ''' Shows certificate manager software
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CertificateManager() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "certmgr.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Function

        ''' <summary>
        ''' Shows filesystem management software pre-installed on the Windows operating system
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FileSystemManagement() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "fsmgmt.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
            Return True
        End Function

        ''' <summary>
        ''' Shows local user and groups manager software
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LocalUserManager() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "lusmgr.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Function

        ''' <summary>
        ''' Shows Management Console software
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ManagementConsole() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "mmc.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Function

        ''' <summary>
        ''' Shwos Windows Management Instrumentation console
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WMIManagment() As Boolean
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "WmiMgmt.msc"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Function
    End Class

    ''' <summary>
    ''' Gives easy access to Registry Editor software pre-installed on the Windows operating system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RegEdit
        ''' <summary>
        ''' Shows Registry Editor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Show()
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "regedit.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub
    End Class

    ''' <summary>
    ''' Gives easy access to command prompt
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CommandPrompt
        Public Sub Show()
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "cmd.exe"
                .UseShellExecute = False
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub
    End Class

    ''' <summary>
    ''' Shows Microsoft Windows Operating system license text file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class License
        Public Sub ShowLicense()
            Process.Start("license.rtf")
        End Sub
    End Class

    ''' <summary>
    ''' Shows Windows version reporter applet
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WinVer
        Public Sub Show()
            Dim pr As Process = New Process
            With pr.StartInfo
                .CreateNoWindow = True
                .FileName = "winver.exe"
                .UseShellExecute = True
            End With
            pr.Start()
            pr.WaitForExit()
        End Sub
    End Class

    ''' <summary>
    ''' Gives functionailty over the Windos Explorer.
    ''' Can be used for, in example, restarting Explorer.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WinExplorer
        Public Function RestartExplorer() As Boolean
            Dim pr As Process = New Process
            pr.StartInfo.FileName = "explorer.exe"
            pr.Kill()
        End Function
    End Class
End Namespace
