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

    ''' <summary>
    ''' Gets information about the computer's BIOS by querying WMI for information 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class BIOS

        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_BIOS")

        ''' <summary>
        ''' Gets BIOS Version
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBIOSVer() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BIOSVersion") Is Nothing Then
                    Return "Version: " & obj("BIOSVersion").ToString()
                Else : Return "Version not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Internal identifier for this compilation of this software element.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBuildNo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BuildNumber") Is Nothing Then
                    Return "Build Number: " & obj("BuildNumber").ToString()
                Else : Return "Build Number not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Code set used by this software element.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCodeSet() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CodeSet") Is Nothing Then
                    Return "Code Set: " & obj("CodeSet").ToString()
                Else : Return "Code Set not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Name of the current BIOS language.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCurrentLang() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CurrentLanguage") Is Nothing Then
                    Return "Current Language: " & obj("CurrentLanguage").ToString()
                Else : Return "Current Language not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Manufacturer's identifier for this software element. Often this will be a stock keeping unit (SKU) or a part number.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getIdentCode() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("IdentificationCode") Is Nothing Then
                    Return "Identification Code: " & obj("IdentificationCode").ToString()
                Else : Return "Identification Code not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Number of languages available for installation on this system. 
        ''' Language may determine properties such as the need for Unicode and bidirectional text.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInstallableLangs() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("InstallableLanguages") Is Nothing Then
                    Return "Installable Languages: " & obj("InstallableLanguages").ToString()
                Else : Return "Installable Languages not Available"
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
                    Return "Install Date: " & obj("InstallDate").ToString()
                Else : Return "Install Date not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Language edition of this software element. 
        ''' The language codes defined in ISO 639 should be used. Where the software element represents a 
        ''' multilingual or international version of a product, the string "multilingual" should be used.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLangEdition() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("LanguageEdition") Is Nothing Then
                    Return "Language Edition: " & obj("LanguageEdition").ToString()
                Else : Return "Language Edition not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Manufacturer of this software element.
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
        ''' Records the manufacturer and operating system type for a software element when the 
        ''' TargetOperatingSystem property has a value of 1 (Other). When TargetOperatingSystem has a value of 1, 
        ''' OtherTargetOS must have a nonnull value. For all other values of TargetOperatingSystem, OtherTargetOS is NULL.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOtherTargetOS() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("OtherTargetOS") Is Nothing Then
                    Return "Other Target OS: " & obj("OtherTargetOS").ToString()
                Else : Return "Other Target OS not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' If TRUE, this is the primary BIOS of the computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsPrimaryBIOS() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PrimaryBIOS") Is Nothing Then
                    Dim pb As String = obj("PrimaryBIOS").ToString()
                    Return Boolean.Parse(pb)
                Else : Return True
                End If
            Next
        End Function

        ''' <summary>
        ''' Release date of the Windows BIOS in the Coordinated Universal Time (UTC) format of YYYYMMDDHHMMSS.MMMMMM(+-)OOO.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetReleaseDate() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ReleaseDate") Is Nothing Then
                    Return "Release Date: " & obj("").ToString()
                Else : Return "Release Date not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Assigned serial number of the software element.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSerialNumber() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SerialNumber") Is Nothing Then
                    Return "Serial Number: " & obj("SerialNumber").ToString()
                Else : Return "Serial Number not Available"
                End If
            Next
        End Function
    End Class

    ''' <summary>
    ''' Gets information about the computer's boot configuration by querying WMI for information
    ''' </summary>
    ''' <remarks></remarks>
    Public Class BootConfig

        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_BootConfiguration")

        ''' <summary>
        ''' Path to the system files required for booting the system.
        ''' Example: "C:\Windows"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBootDir() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BootDirectory") Is Nothing Then
                    Return "Boot Directory: " & obj("BootDirectory").ToString()
                Else : Return "Boot Directory not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Path to the configuration files. This value may be similar to the value in the BootDirectory property.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetConfigPath() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ConfigurationPath") Is Nothing Then
                    Return "Configuration Path: " & obj("ConfigurationPath").ToString()
                Else : Return "Configuration Path not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Last drive letter to which a physical drive is assigned.
        ''' Example: "E:"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastDrive() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("LastDrive") Is Nothing Then
                    Return "Last Drive to be Assigned: " & obj("LastDrive").ToString()
                Else : Return "Last Drive to be Assigned not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Directory where temporary files can reside during boot time.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetScratchDir() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ScratchDirectory") Is Nothing Then
                    Return "Scratch Directory: " & obj("ScratchDirectory").ToString()
                Else : Return "Scratch Directory not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Directory where temporary files are stored.
        ''' Example: "C:\TEMP"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTempDir() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("TempDirectory") Is Nothing Then
                    Return "Temp Directory: " & obj("TempDirectory").ToString()
                Else : Return "Temp Directory not Available"
                End If
            Next
        End Function
    End Class

    ''' <summary>
    ''' Gets information about the computer's CD ROM Drive by querying WMI for information
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CDROMDrive

        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_CDROMDrive")

        ''' <summary>
        ''' Algorithm or tool used by the device to support compression. 
        ''' If it is not possible or not desired to describe the compression scheme 
        ''' (perhaps because it is not known), use the following words: "Unknown" to 
        ''' represent that it is not known whether the device supports compression capabilities; 
        ''' "Compressed" to represent that the device supports compression capabilities but either 
        ''' its compression scheme is not known or not disclosed; and "Not Compressed" to represent 
        ''' that the device does not support compression capabilities.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCompressionMethod() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("CompressionMethod") Is Nothing Then
                    Return "Compression Method: " & obj("CompressionMethod").ToString()
                Else : Return "Compression Method not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Default block size, in bytes, for this device.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDefBlockSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DefaultBlockSize") Is Nothing Then
                    Return "Default Block Size (In Bytes): " & obj("DefaultBlockSize").ToString()
                Else : Return "Default Block Size (In Bytes) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Unique identifier for a CD-ROM drive.
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
        ''' Drive letter of the CD-ROM drive.
        ''' Example: "d:\"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDriveLetter() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Drive") Is Nothing Then
                    Return "Drive Letter: " & obj("Drive").ToString()
                Else : Return "Drive Letter not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Access type: Read-only
        ''' Date and time the object is installed. 
        ''' This property does not need a value to indicate that the object is installed.
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
        ''' Manufacturer of the Windows CD-ROM drive.
        ''' Example: "PLEXTOR"
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
        ''' Maximum block size, in bytes, for media accessed by this device.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxBlockSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MaxBlockSize") Is Nothing Then
                    Return "Max. Block Size (In Bytes): " & obj("MaxBlockSize").ToString()
                Else : Return "Max. Block Size (In Bytes) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Maximum length of a filename component supported by the Windows CD-ROM drive. 
        ''' A file name component the portion of a file name between backslashes. 
        ''' The value can be used to indicate that long names are supported by the specified file system. 
        ''' For example, for a FAT file system supporting long names, the function stores the value 255, 
        ''' rather than the previous 8.3 indicator. Long names can also be supported on systems that use 
        ''' the NTFS file system.
        ''' Example: 255
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxComponentLength() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MaximumComponentLength") Is Nothing Then
                    Return "Max. Filename Length (In Chars): " & obj("MaximumComponentLength").ToString()
                Else : Return "Max Filename Length (In Chars) not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Maximum size, in kilobytes, of media supported by this device.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxMediaSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MaxMediaSize") Is Nothing Then
                    Return "Max. Media Size: " & obj("MaxMediaSize").ToString()
                Else : Return "Max. Media Size not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' If True, a CD-ROM is in the drive.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsMediaLoaded() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MediaLoaded") Is Nothing Then
                    Return Boolean.Parse(obj("MediaLoaded").ToString())
                Else : Return False
                End If
            Next
        End Function

        ''' <summary>
        '''  Type of media that can be used or accessed by this device. Possible values are:
        ''' CdRomOnly
        ''' CdRomWrite
        ''' DVDRomOnly
        ''' DVDRomWrite
        '''
        ''' Windows Server 2003 and Windows XP:  Possible values are:
        '''
        ''' Random Access
        ''' Supports Writing
        ''' Removable Media
        ''' CD-ROM 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMediaType() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MediaType") Is Nothing Then
                    Return "Media Type: " & obj("MediaType").ToString()
                Else : Return "Media Type not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Firmware revision level that is assigned by the manufacturer.
        ''' Windows Server 2003 and Windows XP:  This property is not available.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMfrAssignedRevLevel() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MfrAssignedRevisionLevel") Is Nothing Then
                    Return "Revision Level: " & obj("MfrAssignedRevisionLevel").ToString()
                Else : Return "Revision Level not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Minimum block size, in bytes, for media accessed by this device.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMinBlockSize() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("MinBlockSize") Is Nothing Then
                    Return "Min. Block Size: " & obj("MinBlockSize").ToString()
                Else : Return "Min. Block Size not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Maximum number of media that can be supported or inserted, when the media access device supports multiple individual media.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNoOfMediaSupported() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("NumberOfMediaSupported") Is Nothing Then
                    Return "Number of Media Supported: " & obj("NumberOfMediaSupported").ToString()
                Else : Return "Number of Media Supported not Available"
                End If
            Next
        End Function
    End Class

    ''' <summary>
    ''' Gets highly administrative information from Computer.
    ''' 
    ''' Please use with care.
    ''' Only use this with complete agreement of the user.
    ''' At best you use some sort of captcha or password to seal data like this.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ComputerSystem

        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_ComputerSystem")

        ''' <summary>
        ''' 
        ''' Gets status of administrator password. 
        ''' For security reasons, no strings will be output.
        ''' 
        ''' <returns>
        ''' Value	Meaning
        ''' 1 (0x1) = Disabled
        ''' 2 (0x2) = Enabled
        ''' 3 (0x3) = Not Implemented
        ''' 4 (0x4) = Unknown
        ''' </returns>
        ''' </summary>
        ''' <remarks></remarks>
        Public Function GetAdminPasswordStatus() As Int16
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("AdminPasswordStatus") Is Nothing Then
                    Return Int16.Parse(obj("AdminPasswordStatus").ToString())
                End If
            Next
        End Function

        ''' <summary>
        ''' System hardware security settings for administrator password status.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsPagefileAutoManaged() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Return Boolean.Parse(obj("AutomaticManagedPagefile").ToString)
            Next
        End Function

        ''' <summary>
        ''' If True, the system manages the page file.
        ''' Windows Server 2003 and Windows XP:  This property is not available.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AutomaticResetBootOption() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Boolean.Parse(obj("AutomaticResetBootOption").ToString())
            Next
        End Function

        ''' <summary>
        ''' If True, the automatic reset boot option is enabled.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AutomaticResetCapability() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Boolean.Parse(obj("AutomaticResetCapability").ToString())
            Next
        End Function

        ''' <summary>
        ''' If True, indicates whether a boot ROM is supported.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsBootROMSupported() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Boolean.Parse(obj("BootROMSupported").ToString())
            Next
        End Function

        ''' <summary>
        ''' System is started. Fail-safe boot bypasses the user startup files—also called SafeBoot.
        ''' The following list contains the required values:
        '''
        ''' "Normal boot"
        ''' "Fail-safe boot"
        ''' "Fail-safe with network boot" 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBootupState() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("BootupState") Is Nothing Then
                    Return "Bootup State: " & obj("BootupState").ToString()
                Else : Return "Bootup State not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' If True, the daylight savings mode is ON.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsDaylightSavingActive() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Boolean.Parse(obj("DaylightInEffect").ToString())
            Next
        End Function

        ''' <summary>
        ''' Name of local computer according to the domain name server (DNS).
        ''' Windows XP:  This property is not available
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDNSHostName() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DNSHostName") Is Nothing Then
                    Return "DNS Host Name: " & obj("DNSHostName").ToString()
                Else : Return "DNS Host Name not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Name of the domain to which a computer belongs.
        ''' Note  If the computer is not part of a domain, then the name of the workgroup is returned.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDomain() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("Domain") Is Nothing Then
                    Return "Domain: " & obj("Domain").ToString()
                Else : Return "Domain not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Role of a computer in an assigned domain workgroup. 
        ''' A domain workgroup is a collection of computers on the same network. 
        ''' For example, a DomainRole property may show that a computer is a 
        ''' member workstation.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDomainRole() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("DomainRole") Is Nothing Then
                    Dim o As Int32 = Int32.Parse(obj("DomainRole").ToString())
                    If o = 0 Then
                        Return "Domain Role: Standalone Workstation"
                    ElseIf o = 1 Then
                        Return "Domain Role: Member Workstation"
                    ElseIf o = 2 Then
                        Return "Domain Role: Standalone Server"
                    ElseIf o = 3 Then
                        Return "Domain Role: Member Server"
                    ElseIf o = 4 Then
                        Return "Domain Role: Backup Domain Controller"
                    ElseIf o = 5 Then
                        Return "Domain Role: Primary Domain Controller"
                    End If
                Else : Return "Domain Role not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Gets front panel reset status.
        ''' (Whether the front panel has a reset button or not)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFrontPanelResetStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("FrontPanelResetStatus") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("FrontPanelResetStatus").ToString())
                    If o = 0 Then
                        Return "Front Panel Reset Status: Disabled"
                    ElseIf o = 1 Then
                        Return "Front Panel Reset Status: Enabled"
                    ElseIf o = 2 Then
                        Return "Front Panel Reset Status: Not Implemented"
                    ElseIf o = 3 Then
                        Return "Front Panel Reset Status: Unknown"
                    End If
                Else : Return "Front Panel Reset Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' If True, an infrared (IR) port exists on a computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsInfraredSupported() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Return Boolean.Parse(obj("InfraredSupported").ToString())
            Next
        End Function

        ''' <summary>
        ''' Data required to find the initial load device or boot service to request that the operating system start up.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInitialLoadInfo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("InitialLoadInfo") Is Nothing Then
                    Return "Initial Load Info: " & obj("InitialLoadInfo").ToString()
                Else : Return "Initial Load Info not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Object is installed. An object does not need a value to indicate that it is installed.
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
        ''' System hardware security settings for Keyboard Password Status.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetKeyboardPassStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("KeyboardPasswordStatus") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("KeyboardPasswordStatus").ToString())
                    If o = 0 Then
                        Return "Keyboard Password Status: Disabled"
                    ElseIf o = 1 Then
                        Return "Keyboard Password Status: Enabled"
                    ElseIf o = 2 Then
                        Return "Keyboard Password Status: Not Implemented"
                    ElseIf o = 3 Then
                        Return "Keyboard Password Status: Unknown"
                    End If
                Else : Return "Keyboard Password Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Array entry of the InitialLoadInfo property that contains the data to start the loaded operating system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastLoadInfo() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("LastLoadInfo") Is Nothing Then
                    Return "Last Load Info: " & obj("LoadLoadInfo").ToString()
                Else : Return "Loast Load Info not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Name of a computer manufacturer.
        ''' Example: Adventure Works
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
        ''' Product name that a manufacturer gives to a computer. This property must have a value.
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
        ''' Computer system Name value that is generated automatically. 
        ''' The CIM_ComputerSystem object and its derivatives are top-level 
        ''' objects of the Common Information Model (CIM). They provide the 
        ''' scope for several components. Unique CIM_System keys are required, 
        ''' but you can define a heuristic to create the CIM_ComputerSystem name 
        ''' that generates the same name, and is independent from the discovery 
        ''' protocol. This prevents inventory and management problems when the 
        ''' same asset or entity is discovered multiple times, but cannot be 
        ''' resolved to one object. Using a heuristic is recommended, but not required.
        ''' 
        ''' The heuristic is outlined in the CIM V2 Common Model specification, 
        ''' and assumes that the documented rules are used to determine and assign a name. 
        ''' The NameFormat values list defines the order to assign a computer system name. 
        ''' Several rules map to the same value.
        ''' 
        ''' The CIM_ComputerSystem Name value that is calculated using the heuristic 
        ''' is the key value of the system. However, use aliases to assign a different 
        ''' name for CIM_ComputerSystem, which can be more unique to your company. 
        ''' This property is inherited from CIM_System.
        ''' 
        ''' The following list identifies the values for this property.
        ''' "IP"
        ''' "Dial"
        ''' "HID"
        ''' "NWA"
        ''' "HWA"
        ''' "X25"
        ''' "ISDN"
        ''' "IPX"
        ''' "DCC"
        ''' "ICD"
        ''' "E.164"
        ''' "SNA"
        ''' "OID/OSI"
        ''' "Other" 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNameFormat() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("NameFormat") Is Nothing Then
                    Return "Name Format: " & obj("NameFormat").ToString()
                Else : Return "Name Format not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' If True, the network Server Mode is enabled.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsNetworkServerModeEnabled() As Boolean
            For Each obj As ManagementObject In searcher.Get()
                Return Boolean.Parse(obj("NetworkServerModeEnabled").ToString())
            Next
        End Function

        ''' <summary>
        ''' Type of the computer in use, such as laptop, desktop, or Tablet.
        ''' Windows Server 2003 and Windows XP:  This property is not available. 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPCSystemType() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PCSystemType") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("PCSystemType").ToString())
                    If o = 0 Then
                        Return "PC System Type: Unspecified"
                    ElseIf o = 1 Then
                        Return "PC System Type: Desktop"
                    ElseIf o = 2 Then
                        Return "PC System Type: Mobile"
                    ElseIf o = 3 Then
                        Return "PC System Type: Workstation"
                    ElseIf o = 4 Then
                        Return "PC System Type: Enterprise Server"
                    ElseIf o = 5 Then
                        Return "PC System Type: Small Office and Home Office (SOHO) Server"
                    ElseIf o = 6 Then
                        Return "PC System Type: Appliance PC"
                    ElseIf o = 7 Then
                        Return "PC System Type: Performance Server"
                    ElseIf o = 8 Then
                        Return "PC System Type: Maximum"
                    End If
                Else : Return "PC System Type not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' System hardware security settings for Power-On Password Status.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPowerOnPasswordStatus() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PowerOnPasswordStatus") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("PowerOnPasswordStatus").ToString())
                    If o = 0 Then
                        Return "Power On Password Status: Disabled"
                    ElseIf o = 1 Then
                        Return "Power On Password Status: Enabled"
                    ElseIf o = 2 Then
                        Return "Power On Password Status: Not Implemented"
                    ElseIf o = 3 Then
                        Return "Power On Password Status: Unknown"
                    End If
                Else : Return "Power On Password Status not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Current power state of a computer and its associated operating system. 
        ''' The power saving states have the following values: Value 4 (Unknown) 
        ''' indicates that the system is known to be in a power save mode, 
        ''' but its exact status in this mode is unknown; 2 (Low Power Mode) 
        ''' indicates that the system is in a power save state, but still 
        ''' functioning and may exhibit degraded performance; 3 (Standby) indicates 
        ''' that the system is not functioning, but could be brought to full power 
        ''' quickly; and 7 (Warning) indicates that the computer system is in a 
        ''' warning state and a power save mode.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPowerState() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PowerState") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("PowerState").ToString())
                    If o = 0 Then
                        Return "Power State: Unknown"
                    ElseIf o = 1 Then
                        Return "Power State: Full Power"
                    ElseIf o = 2 Then
                        Return "Power State: Power Save (Low Power Mode)"
                    ElseIf o = 3 Then
                        Return "Power State: Power Save (Standby)"
                    ElseIf o = 4 Then
                        Return "Power State: Power Save (Unknown)"
                    ElseIf o = 5 Then
                        Return "Power State: Power Cycle"
                    ElseIf o = 6 Then
                        Return "Power State: Power Off"
                    ElseIf o = 7 Then
                        Return "Power State: Power Save (Warning)"
                    End If
                Else : Return "Power State not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' State of the power supply or supplies when last booted
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPowerSupplyState() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PowerSupplyState") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("PowerSupplyState").ToString())
                    If o = 1 Then
                        Return "Power Supply State: Other"
                    ElseIf o = 2 Then
                        Return "Power Supply State: Unknown"
                    ElseIf o = 3 Then
                        Return "Power Supply State: Safe"
                    ElseIf o = 4 Then
                        Return "Power Supply State: Warning"
                    ElseIf o = 5 Then
                        Return "Power Supply State: Critical"
                    ElseIf o = 6 Then
                        Return "Power Supply State: Nonrecoverable"
                    End If
                Else : Return "Power Supply State not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Contact information for the primary system owner, for example, phone number, email address, and so on.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOwnerContact() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PrimaryOwnerContact") Is Nothing Then
                    Return "Owner Contact Details: " & obj("PrimaryOwnerContact").ToString()
                Else : Return "Owner Contact Details not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Name of the primary system owner.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetOwnerName() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("PrimaryOwnerName") Is Nothing Then
                    Return "Owner Name: " & obj("PrimaryOwnerName").ToString()
                Else : Return "Owner Name not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Number of automatic resets since the last reset. A value of –1 (minus one) indicates that the count is unknown.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetResetCount() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ResetCount") Is Nothing Then
                    If Not Int16.Parse(obj("ResetCount").ToString()) = -1 Then
                        Return "Reset Count: " & obj("ResetCount").ToString()
                    Else
                        Return "Reset Count: Unknown"
                    End If
                Else : Return "Reset Count not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Current status of an object. Various operational and nonoperational statuses can be defined. 
        ''' Operational statuses include: OK, Degraded, and Pred Fail, which is an element such as a 
        ''' SMART-enabled hard disk drive that may be functioning properly, but predicts a failure in 
        ''' the near future. Nonoperational statuses include: Error, Starting, Stopping, and Service, 
        ''' which can apply during mirror-resilvering of a disk, reloading a user permissions list, 
        ''' or other administrative work. Not all status work is online, but the managed element is 
        ''' not OK or in one of the other states. 
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
        ''' System running on the Windows-based computer. This property must have a value.
        ''' The following list identifies some of the possible values for this property.
        ''' 
        ''' "x64-based PC"
        ''' "X86-based PC"
        ''' "MIPS-based PC"
        ''' "Alpha-based PC"
        ''' "Power PC"
        ''' "SH-x PC"
        ''' "StrongARM PC"
        ''' "64-bit Intel PC"
        ''' "64-bit Alpha PC"
        ''' "Unknown"
        ''' "X86-Nec98 PC" 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSysType() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("SystemType") Is Nothing Then
                    Return "System Type: " & obj("SystemType").ToString()
                Else : Return "System Type not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Thermal state of the system when last booted.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetThermalState() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("ThermalState") Is Nothing Then
                    Dim o As Int16 = Int16.Parse(obj("ThermalState").ToString())
                    If o = 1 Then
                        Return "Thermal State: Other"
                    ElseIf o = 2 Then
                        Return "Thermal State: Unknown"
                    ElseIf o = 3 Then
                        Return "Thermal State: Safe"
                    ElseIf o = 4 Then
                        Return "Thermal State: Warning"
                    ElseIf o = 5 Then
                        Return "Thermal State: Critical"
                    ElseIf o = 6 Then
                        Return "Thermal State: Nonrecoverable"
                    End If
                Else : Return "Thermal State not Available"
                End If
            Next
        End Function

        ''' <summary>
        ''' Name of a user that is logged on currently. This property must have a value. 
        ''' In a terminal services session, UserName returns the name of the user that 
        ''' is logged on to the console—not the user logged on during the terminal service session.
        ''' Example: jeffsmith
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserName() As String
            For Each obj As ManagementObject In searcher.Get()
                If Not obj("UserName") Is Nothing Then
                    Return "Username: " & obj("UserName").ToString()
                Else : Return "Username not Available"
                End If
            Next
        End Function
    End Class
End Namespace
