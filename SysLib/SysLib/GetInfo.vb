Option Strict On
Imports System.Management
Imports System.Console

Namespace GetInfo
    ''' <summary>
    ''' Gathers CPU information by querying for WMI data.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CPU

        Private searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_Processor")

        ''' <summary>
        ''' Gets CPU address width (How many bits the processor supports (32x (x86) or 64x)
        ''' </summary>
        ''' <Usage>Labelx.Text = GetAddressWidth()</Usage>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAddressWidth() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("AddressWidth") Is Nothing Then
                    Return "Address Width: " & obj("AddressWidth").ToString()
                Else : Return "Address Width not Available."
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU architecture()
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetArchitecture() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Architecture") Is Nothing Then
                    Return "Architecture: " & obj("Architecture").ToString()
                Else : Return "Architecture not Available..."
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets processor availability
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAvailability() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Availability") Is Nothing Then
                    Return "Availability: " & obj("Availability").ToString()
                Else : Return "Availability not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU caption
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCaption() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Caption") Is Nothing Then
                    Return "Caption: " & obj("Caption").ToString()
                Else : Return "Caption not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets configuration manager error codes from CPU.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetConfigManagerErrorCode() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ConfigManagerErrorCode") Is Nothing Then
                    Return "Config Manager Error Code: " & obj("ConfigManagerErrorCode").ToString()
                Else : Return "Config Manager Error Code not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets configuration manager user configuration.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetConfigManagerUserConfig() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ConfigManagerUserConfig") Is Nothing Then
                    Return "Config Manager User Config: " & obj("ConfigManagerUserConfig").ToString()
                Else : Return "Config Manager User Config not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets processor status
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCpuStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CpuStatus") Is Nothing Then
                    Return "CPU Status: " & obj("CpuStatus").ToString()
                Else : Return "CPU Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets current clock speed of CPU. 
        ''' Is not interactive, will stay fixed until new values are requested.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCurrentClockSpeed() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CurrentClockSpeed") Is Nothing Then
                    Return "Current Clock Speed (Static): " & obj("CurrentClockSpeed").ToString()
                Else : Return "Current Clock Speed not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets current voltage of CPU.
        ''' Is not interactive, will stay fixed until new values are requested.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCurrentVoltage() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CurrentVoltage") Is Nothing Then
                    Return "Current Voltage: " & obj("CurrentVoltage").ToString()
                Else : Return "Current Voltage not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets data width from CPU
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataWidth() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DataWidth") Is Nothing Then
                    Return "Data Width: " & obj("DataWidth").ToString()
                Else : Return "Data Width not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets processor description. Rather useless in most occasions
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDescription() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Description") Is Nothing Then
                    Return "Description: " & obj("Description").ToString()
                Else : Return "Description not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's ID.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDeviceID() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DeviceID") Is Nothing Then
                    Return "Device ID: " & obj("DeviceID").ToString()
                Else : Return "Device ID not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets if error is cleared or not.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetErrorCleared() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ErrorCleared") Is Nothing Then
                    Return "Error Cleared: " & obj("ErrorCleared").ToString()
                Else : Return "Error Cleared is not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets error descriptions if available.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetErrorDescription() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ErrorDescription") Is Nothing Then
                    Return "Error Description: " & obj("ErrorDescription").ToString()
                Else : Return "Error Description not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets FSB speeds (External Clock)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetExtClock() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ExtClock") Is Nothing Then
                    Return "External Clock: " & obj("ExtClock").ToString()
                Else : Return "External Clock not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU family
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFamily() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Family") Is Nothing Then
                    Return "Family: " & obj("Family").ToString()
                Else : Return "Family not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU install date
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstallDate() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("InstallDate") Is Nothing Then
                    Return "Install Date: " & obj("InstallDate").ToString()
                Else : Return "Install Date not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets Level2 Cache size
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetL2CacheSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("L2CacheSize") Is Nothing Then
                    Return "L2 Cache Size: " & obj("L2CacheSize").ToString()
                Else : Return "L2 Cache Size not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets level 2 cache speed
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetL2CacheSpeed() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("L2CacheSpeed") Is Nothing Then
                    Return "L2 Cache Speed: " & obj("L2CacheSpeed").ToString()
                Else : Return "L2 Cache Speed not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets level 3 cache size
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetL3CacheSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("L3CacheSize") Is Nothing Then
                    Return "L3 Cache Size: " & obj("L3CacheSize").ToString()
                Else : Return "L3 Cache Size not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets level 3 cache speed
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetL3CacheSpeed() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("L3CacheSpeed") Is Nothing Then
                    Return "L3 Cache Speed: " & obj("L3CacheSpeed").ToString()
                Else : Return "L3 Cache Speed not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's last error code
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastErrorCode() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("LastErrorCode") Is Nothing Then
                    Return "Last Error Code: " & obj("LastErrorCode").ToString()
                Else : Return "Last Error Code not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU level
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLevel() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Level") Is Nothing Then
                    Return "Level: " & obj("Level").ToString()
                Else : Return "Level not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's current load percentage
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLoadPercentage() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("LoadPercentage") Is Nothing Then
                    Return "Load Percentage: " & obj("LoadPercentage").ToString() & "%"
                Else : Return "Load Percentage not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU manufacturer
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetManufacturer() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Manufacturer") Is Nothing Then
                    Return "Manufacturer: " & obj("Manufacturer").ToString()
                Else : Return "Manufacturer not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets maximum clock speed of CPU
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxClockSpeed() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MaxClockSpeed") Is Nothing Then
                    Return "Max. Clock Speed: " & obj("MaxClockSpeed").ToString() & "MHz"
                Else : Return "Max. Clock Speed not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's name (E.G.: Pentium IV)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetName() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Name") Is Nothing Then
                    Return "Name: " & obj("Name").ToString()
                Else : Return "Name not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets number of CPU cores
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNumOfCores() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("NumberOfCores") Is Nothing Then
                    Return "Number of Cores: " & obj("NumberOfCores").ToString()
                Else : Return "Number of Cores not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets number of CPU threads (Logical processors)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getNumOfThread() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("NumberOfLogicalProcessors") Is Nothing Then
                    Return "Number of Logical Processors: " & obj("NumberOfLogicalProcessors").ToString()
                Else : Return "Number of Logical Processors not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets other family description from CPU
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOtherFamDescription() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("OtherFamilyDescription") Is Nothing Then
                    Return "Other Family Description: " & obj("OtherFamilyDescription").ToString()
                Else : Return "Other Family Description not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's Plug 'n' Play Device ID
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPNPDeviceID() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PNPDeviceID") Is Nothing Then
                    Return "PNP Device ID: " & obj("PNPDeviceID").ToString()
                Else : Return "PNP Device ID not available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets Power Management Supported value (True or False)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPowerMgmtSupported() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PowerManagementSupported") Is Nothing Then
                    Return "Power Mgmt. Supported: " & obj("PowerManagementSupported").ToString()
                Else : Return "Power Mgmt. Supported not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets processor type
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCpuType() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ProcessorType") Is Nothing Then
                    Return "CPU Type: " & obj("ProcessorType").ToString()
                Else : Return "CPU Type not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU revision
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCpuRev() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Revision") Is Nothing Then
                    Return "CPU Revision: " & obj("Revision").ToString()
                Else : Return "CPU Revision not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU role
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCpuRole() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Role") Is Nothing Then
                    Return "CPU Role: " & obj("Role").ToString()
                Else : Return "CPU Role not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's socket designation
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSocDesignation() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SocketDesignation") Is Nothing Then
                    Return "Socket Designation: " & obj("SocketDesignation").ToString()
                Else : Return "Socket Designation not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's status
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Status") Is Nothing Then
                    Return "Status: " & obj("Status").ToString()
                Else : Return "Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU stepping
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStepping() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Stepping") Is Nothing Then
                    Return "Stepping: " & obj("Stepping").ToString()
                Else : Return "Stepping not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's version
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetVersion() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Version") Is Nothing Then
                    Return "Version: " & obj("Version").ToString()
                Else : Return "Version not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets CPU's voltage caps
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetVoltageCaps() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("VoltageCaps") Is Nothing Then
                    Return "Voltage Caps: " & obj("VoltageCaps").ToString()
                Else : Return "Voltage Caps not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Queries for all data and displays it inside a console window.
        ''' Must be called from a console application!
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function ShowInfo_Console() As Int32
            Try
                For Each queryObj As ManagementObject In searcher.Get()

                    Console.WriteLine("-----------------------------------")
                    Console.WriteLine("Win32_Processor instance")
                    Console.WriteLine("-----------------------------------")
                    Console.WriteLine("AddressWidth: {0}", queryObj("AddressWidth"))
                    Console.WriteLine("Architecture: {0}", queryObj("Architecture"))
                    Console.WriteLine("Availability: {0}", queryObj("Availability"))
                    Console.WriteLine("Caption: {0}", queryObj("Caption"))
                    Console.WriteLine("ConfigManagerErrorCode: {0}", queryObj("ConfigManagerErrorCode"))
                    Console.WriteLine("ConfigManagerUserConfig: {0}", queryObj("ConfigManagerUserConfig"))
                    Console.WriteLine("CpuStatus: {0}", queryObj("CpuStatus"))
                    Console.WriteLine("CreationClassName: {0}", queryObj("CreationClassName"))
                    Console.WriteLine("CurrentClockSpeed: {0}", queryObj("CurrentClockSpeed"))
                    Console.WriteLine("CurrentVoltage: {0}", queryObj("CurrentVoltage"))
                    Console.WriteLine("DataWidth: {0}", queryObj("DataWidth"))
                    Console.WriteLine("Description: {0}", queryObj("Description"))
                    Console.WriteLine("DeviceID: {0}", queryObj("DeviceID"))
                    Console.WriteLine("ErrorCleared: {0}", queryObj("ErrorCleared"))
                    Console.WriteLine("ErrorDescription: {0}", queryObj("ErrorDescription"))
                    Console.WriteLine("ExtClock: {0}", queryObj("ExtClock"))
                    Console.WriteLine("Family: {0}", queryObj("Family"))
                    Console.WriteLine("InstallDate: {0}", queryObj("InstallDate"))
                    Console.WriteLine("L2CacheSize: {0}", queryObj("L2CacheSize"))
                    Console.WriteLine("L2CacheSpeed: {0}", queryObj("L2CacheSpeed"))
                    Console.WriteLine("L3CacheSize: {0}", queryObj("L3CacheSize"))
                    Console.WriteLine("L3CacheSpeed: {0}", queryObj("L3CacheSpeed"))
                    Console.WriteLine("LastErrorCode: {0}", queryObj("LastErrorCode"))
                    Console.WriteLine("Level: {0}", queryObj("Level"))
                    Console.WriteLine("LoadPercentage: {0}", queryObj("LoadPercentage"))
                    Console.WriteLine("Manufacturer: {0}", queryObj("Manufacturer"))
                    Console.WriteLine("MaxClockSpeed: {0}", queryObj("MaxClockSpeed"))
                    Console.WriteLine("Name: {0}", queryObj("Name"))
                    Console.WriteLine("NumberOfCores: {0}", queryObj("NumberOfCores"))
                    Console.WriteLine("NumberOfLogicalProcessors: {0}", queryObj("NumberOfLogicalProcessors"))
                    Console.WriteLine("OtherFamilyDescription: {0}", queryObj("OtherFamilyDescription"))
                    Console.WriteLine("PNPDeviceID: {0}", queryObj("PNPDeviceID"))

                    If queryObj("PowerManagementCapabilities") Is Nothing Then
                        Console.WriteLine("PowerManagementCapabilities: {0}", queryObj("PowerManagementCapabilities"))
                    Else
                        Dim arrPowerManagementCapabilities As UInt16()
                        arrPowerManagementCapabilities = CType(queryObj("PowerManagementCapabilities"), UShort())
                        For Each arrValue As UInt16 In arrPowerManagementCapabilities
                            Console.WriteLine("PowerManagementCapabilities: {0}", arrValue)
                        Next
                    End If
                    Console.WriteLine("PowerManagementSupported: {0}", queryObj("PowerManagementSupported"))
                    Console.WriteLine("ProcessorId: {0}", queryObj("ProcessorId"))
                    Console.WriteLine("ProcessorType: {0}", queryObj("ProcessorType"))
                    Console.WriteLine("Revision: {0}", queryObj("Revision"))
                    Console.WriteLine("Role: {0}", queryObj("Role"))
                    Console.WriteLine("SocketDesignation: {0}", queryObj("SocketDesignation"))
                    Console.WriteLine("Status: {0}", queryObj("Status"))
                    Console.WriteLine("StatusInfo: {0}", queryObj("StatusInfo"))
                    Console.WriteLine("Stepping: {0}", queryObj("Stepping"))
                    Console.WriteLine("SystemCreationClassName: {0}", queryObj("SystemCreationClassName"))
                    Console.WriteLine("SystemName: {0}", queryObj("SystemName"))
                    Console.WriteLine("UniqueId: {0}", queryObj("UniqueId"))
                    Console.WriteLine("UpgradeMethod: {0}", queryObj("UpgradeMethod"))
                    Console.WriteLine("Version: {0}", queryObj("Version"))
                    Console.WriteLine("VoltageCaps: {0}", queryObj("VoltageCaps"))
                Next
            Catch err As ManagementException
                ForegroundColor = ConsoleColor.Red
                BackgroundColor = ConsoleColor.Black
                Write("ERROR WHILE QUERYING FOR DATA" & vbNewLine & err.ToString())
            End Try
        End Function
    End Class

    ''' <summary>
    ''' Gathers motherboard information by querying for WMI data
    ''' </summary>
    ''' <remarks></remarks>
    Public Class BaseBoard

        Private searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_BaseBoard")

        ''' <summary>
        ''' Gets motherboard depth
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getDepth() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Depth") Is Nothing Then
                    Return "Depth: " & obj("Depth").ToString()
                Else : Return "Depth not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard height
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetHeight() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Height") Is Nothing Then
                    Return "Height: " & obj("Height").ToString()
                Else : Return "Height not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets if baseboard is hosting board.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsHostingBoard() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("HostingBoard") Is Nothing Then
                    If obj("HostingBoard").ToString() = "True" Then
                        Return True
                    Else : Return False
                    End If
                Else : Return False
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets if motherboard is hot-swappable
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsHotSwappable() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("HotSwappable") Is Nothing Then
                    If obj("HotSwappable").ToString = "True" Then
                        Return True
                    Else : Return False
                    End If
                Else : Return False
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's install date.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstallDate() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("InstallDate") Is Nothing Then
                    Return "Install Date: " & obj("InstallDate").ToString()
                Else : Return "Install Date not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's manufacturer
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetManufacturer() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Manufacturer") Is Nothing Then
                    Return "Manufacturer: " & obj("Manufacturer").ToString()
                Else : Return "Manufacturer not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard model
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetModel() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Model") Is Nothing Then
                    Return "Model: " & obj("Model").ToString()
                Else : Return "Model not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's name
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetName() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Name") Is Nothing Then
                    Return "Name: " & obj("Name").ToString()
                Else : Return "Name not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets other identifying information about motherboard
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOtherIdentInfo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("OtherIdentifyingInfo") Is Nothing Then
                    Return "Other Identifying Information: " & obj("OtherIdentifyingInfo").ToString()
                Else : Return "Other Identfying Information not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets part number of motherboard
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPartNo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PartNumber") Is Nothing Then
                    Return "Part Number: " & obj("PartNumber").ToString()
                Else : Return "Part Number not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Checks if motherboard is powered on.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsPoweredOn() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PoweredOn") Is Nothing Then
                    If obj("PoweredOn").ToString() = "True" Then Return True Else Return False
                Else : Return True
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard product
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetProduct() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Product") Is Nothing Then
                    Return "Product: " & obj("Product").ToString()
                Else : Return "Product not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Checks if motherboard is removeable
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsRemoveable() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Removeable") Is Nothing Then
                    If obj("Removeable").ToString() = "True" Then Return True Else Return False
                Else : Return True
                End If
            Next
        End Function

        ''' <summary>
        ''' Checks if motherboard is replaceable or not.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsReplaceable() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Replaceable") Is Nothing Then
                    If obj("Replaceable").ToString() = "True" Then Return True Else Return False
                Else : Return True
                End If
            Next
        End Function

        ''' <summary>
        ''' Checks if motherboard requires daughter board
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function RequiresDaughterBoard() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("RequiresDaughterBoard") Is Nothing Then
                    If obj("RequiresDaughterBoard").ToString() = "True" Then Return True Else Return False
                Else : Return False
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's serial number
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSerialNo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SerialNumber") Is Nothing Then
                    Return "Serial Number: " & obj("SerialNumber").ToString()
                Else : Return "Serial Number not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard SKU (What ever the hell that is...)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSKU() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SKU") Is Nothing Then
                    Return "SKU: " & obj("SKU").ToString()
                Else : Return "SKU not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's slot layout
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSlotLayout() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SlotLayout") Is Nothing Then
                    Return "SlotLayout: " & obj("SlotLayout").ToString()
                Else : Return "Slot Layout not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's special requirements
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSpecialRequirements() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SpecialRequirements") Is Nothing Then
                    Return "Special Requirements: " & obj("Special Requirements").ToString()
                Else : Return "Special Requirements not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's status
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If obj("Status") Is Nothing Then
                    Return "Status: " & obj("Status").ToString()
                Else : Return "Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets motherboard's version.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetVersion() As String
            For Each obj As ManagementObject In searcher.Get()
                If obj("Version") Is Nothing Then
                    Return "Version: " & obj("Version").ToString()
                Else : Return "Version not Available"
                End If
            Next
        End Function

        Public Overloads Function ShowInfo_Console() As Int32
            ForegroundColor = ConsoleColor.Green
            Try

                For Each queryObj As ManagementObject In searcher.Get()

                    Console.WriteLine("-----------------------------------")
                    Console.WriteLine("Win32_BaseBoard instance")
                    Console.WriteLine("-----------------------------------")
                    Console.WriteLine("Caption: {0}", queryObj("Caption"))

                    If queryObj("ConfigOptions") Is Nothing Then
                        Console.WriteLine("ConfigOptions: {0}", queryObj("ConfigOptions"))
                    Else
                        Dim arrConfigOptions As String()
                        arrConfigOptions = CType(queryObj("ConfigOptions"), String())
                        For Each arrValue As String In arrConfigOptions
                            Console.WriteLine("ConfigOptions: {0}", arrValue)
                        Next
                    End If
                    Console.WriteLine("CreationClassName: {0}", queryObj("CreationClassName"))
                    Console.WriteLine("Depth: {0}", queryObj("Depth"))
                    Console.WriteLine("Description: {0}", queryObj("Description"))
                    Console.WriteLine("Height: {0}", queryObj("Height"))
                    Console.WriteLine("HostingBoard: {0}", queryObj("HostingBoard"))
                    Console.WriteLine("HotSwappable: {0}", queryObj("HotSwappable"))
                    Console.WriteLine("InstallDate: {0}", queryObj("InstallDate"))
                    Console.WriteLine("Manufacturer: {0}", queryObj("Manufacturer"))
                    Console.WriteLine("Model: {0}", queryObj("Model"))
                    Console.WriteLine("Name: {0}", queryObj("Name"))
                    Console.WriteLine("OtherIdentifyingInfo: {0}", queryObj("OtherIdentifyingInfo"))
                    Console.WriteLine("PartNumber: {0}", queryObj("PartNumber"))
                    Console.WriteLine("PoweredOn: {0}", queryObj("PoweredOn"))
                    Console.WriteLine("Product: {0}", queryObj("Product"))
                    Console.WriteLine("Removable: {0}", queryObj("Removable"))
                    Console.WriteLine("Replaceable: {0}", queryObj("Replaceable"))
                    Console.WriteLine("RequirementsDescription: {0}", queryObj("RequirementsDescription"))
                    Console.WriteLine("RequiresDaughterBoard: {0}", queryObj("RequiresDaughterBoard"))
                    Console.WriteLine("SerialNumber: {0}", queryObj("SerialNumber"))
                    Console.WriteLine("SKU: {0}", queryObj("SKU"))
                    Console.WriteLine("SlotLayout: {0}", queryObj("SlotLayout"))
                    Console.WriteLine("SpecialRequirements: {0}", queryObj("SpecialRequirements"))
                    Console.WriteLine("Status: {0}", queryObj("Status"))
                    Console.WriteLine("Tag: {0}", queryObj("Tag"))
                    Console.WriteLine("Version: {0}", queryObj("Version"))
                    Console.WriteLine("Weight: {0}", queryObj("Weight"))
                    Console.WriteLine("Width: {0}", queryObj("Width"))
                Next
            Catch err As ManagementException
                ForegroundColor = ConsoleColor.Red
                BackgroundColor = ConsoleColor.Black
                WriteLine("An error occurred while querying for WMI data: " & err.Message)
                ForegroundColor = ConsoleColor.Green
            End Try
        End Function

    End Class

    ''' <summary>
    ''' Gets information about the battery by querying WMI for information
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Battery
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_Battery")

        ''' <summary>
        ''' Gets battery recharge time
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRechargeTime() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BatteryRechargeTime") Is Nothing Then
                    Return "Battery Recharge Time: " & obj("BatteryRechargeTime").ToString()
                Else : Return "Battery Recharge Time not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets battery status
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBatteryStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BatteryStatus") Is Nothing Then
                    Return "Battery Status: " & obj("BatteryStatus").ToString()
                Else : Return "Battery Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets battery chemistry
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetChemistry() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Chemistry") Is Nothing Then
                    Return "Chemistry: " & obj("Chemistry").ToString()
                Else : Return "Chemistry not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets battery capacity
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDesignCapacity() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DesignCapacity") Is Nothing Then
                    Return "Design Capacity: " & obj("DesignCapacity").ToString()
                Else : Return "Design Capacity not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets battery voltage
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDesignVoltage() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DesignVoltage") Is Nothing Then
                    Return "Design Voltage: " & obj("DesignVoltage").ToString()
                Else : Return "Design Voltage not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets battery's device ID
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDeviceID() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DeviceID") Is Nothing Then
                    Return "Device ID: " & obj("DeviceID").ToString()
                Else : Return "Device ID not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets estimated remaining charge
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRemainingCharge() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("EstimatedChargeRemaining") Is Nothing Then
                    Return "Estimated Charge Remaining: " & obj("EstimatedChargeRemaining").ToString()
                Else : Return "Estimated Charge Remaining not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Estimate in minutes of the time to battery charge depletion under the present load 
        ''' conditions if the utility power is off, or lost and remains off, or a laptop is disconnected from a power source
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRunTime() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("EstimatedRunTime") Is Nothing Then
                    Return "Estimated Run Time (In Minutes): " & obj("EstimatedRunTime").ToString()
                Else : Return "Estimated Run Time (In Minutes) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Amount of time it takes to completely drain the battery after it is fully charged. 
        ''' This property is no longer used and is considered obsolete.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBatteryLife() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ExpectedBatteryLife") Is Nothing Then
                    Return "Expected Battery Life (In Minutes): " & obj("ExpectedBatteryLife").ToString
                Else : Return "ExpectedBatteryLife not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Battery's expected lifetime in minutes, assuming that the battery is fully charged. 
        ''' The property represents the total expected life of the battery, not its current remaining life, 
        ''' which is indicated by the EstimatedRunTime property.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetExpectedLife() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ExpectedLife") Is Nothing Then
                    Return "Expected Battery Life (In Minutes): " & obj("ExpectedLife").ToString()
                Else : Return "Expected Battery Life (In Minutes) not Available)"
                End If
            Next
        End Function

        ''' <summary>
        ''' Full charge capacity of the battery in milliwatt-hours. 
        ''' Comparison of the value to the DesignCapacity property determines when the battery requires replacement. 
        ''' A battery's end of life is typically when the FullChargeCapacity property falls below 80% of the DesignCapacity property. 
        ''' If the property is not supported, enter 0 (zero). 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFullChargeCapacity() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("FullChargeCapacity") Is Nothing Then
                    Return "Full Charge Capacity (In mW/h): " & obj("FullChargeCapacity").ToString
                Else : Return "Full Charge Capacity (In mW/h) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Date and time the object was installed. This property does not need a value to indicate that the object is installed.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstallDate() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("InstallDate") Is Nothing Then
                    Return "Install Date: " & obj("InstallDate").ToString
                Else : Return "Install Date not Availabke"
                End If
            Next
        End Function

        ''' <summary>
        ''' Maximum time, in minutes, to fully charge the battery. 
        ''' The property represents the time to recharge a fully depleted battery, 
        ''' not the current remaining charge time, which is indicated in the TimeToFullCharge property.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxRechargeTime() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MaxRechargeTime") Is Nothing Then
                    Return "Max Recharge Time (In Minutes): " & obj("MaxRechargeTime").ToString
                Else : Return "Max Recharge Time (In Minutes) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Windows Plug and Play device identifier of the logical device.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPNPDeviceID() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PNPDeviceID") Is Nothing Then
                    Return "Plug 'n' Play Device ID: " & obj("PNPDeviceID").ToString
                Else : Return "Plug 'n' Play Device ID not Available"
                End If
            Next
        End Function
    End Class
End Namespace
