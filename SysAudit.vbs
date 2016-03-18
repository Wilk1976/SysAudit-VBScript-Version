On Error Resume Next
'Copyright David B. Wilkerson 2016
'Written 17-March-2016.
'SysAudit Version 1.0.0.0
const HKEY_LOCAL_MACHINE = &H80000002
Dim env
Dim fso
Dim opf
Dim objShell
Dim prgReg, prgSubKey, prgSubKeyAry
Dim prgRegPath
Dim strComputer
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set env = objShell.environment("Process")
pcName = env.Item("Computername")
Set opf = fso.CreateTextFile(pcName & ".txt", ForWriting, True)
msgbox"Wait for the text file to open, do not open manually.",48,"SysAudit Tool V 1.0.0.0"
strComputer = "."
memType = array("Unknown","Other","DRAM","Synchronous DRAM","Cache DRAM","EDO","EDRAM","VRAM","SRAM","RAM","ROM","Flash","EEPROM","FEPROM","EPROM","CDRAM","3DRAM","SDRAM","SGRAM","RDRAM","DDR","DDR-2","DDR-3","DDR-4")
strMemory = ""
strMemtype = ""
i = 1      
set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set biosSettings = objWMIService.ExecQuery("Select * from Win32_BIOS")
set pcinfo = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
set cpuinfo = objWMIService.ExecQuery("Select * from Win32_Processor")
set memItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
set hdinfo = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = '3'")
set drvinfo = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
set osinfo = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
set laninfo = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")
set tskinfo = objWMIService.ExecQuery("Select * from Win32_Process")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("c:\scripts\software.tsv", True)
Set prgReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
opf.WriteLine "SysAudit Tool Version 1.5"
opf.WriteLine "Audit created; " & Date & " " & Time
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "BIOS and PC Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objBIOS in biosSettings 
    	opf.WriteLine "       BIOS Date: " & left(objBIOS.releaseDate,8)
	opf.WriteLine "    BIOS Version: " & objBIOS.Version
	opf.WriteLine " PC Manufacturer: " & objBIOS.Manufacturer
	opf.WriteLine "PC Serial Number: " & objBIOS.serialNumber
	opf.WriteLine "         PC Name: " & pcName
Next
For Each objPC in pcinfo
	opf.WriteLine "        PC Model: " & objPC.Model
Next
For Each roleinfo in pcinfo
    Select Case roleinfo.DomainRole 
        Case 0 
            strComputerRole = "Standalone Workstation"
        Case 1        
            strComputerRole = "Member Workstation"
        Case 2
            strComputerRole = "Standalone Server"
        Case 3
            strComputerRole = "Member Server"
        Case 4
            strComputerRole = "Backup Domain Controller"
        Case 5
            strComputerRole = "Primary Domain Controller"
    End Select
    	opf.WriteLine "         PC Role: " & strComputerRole
Next
opf.WriteLine ""
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Operating System Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objOS in osinfo
	opf.WriteLine "OS Name and Location: " & objOS.Name
	opf.WriteLine "          OS Version: " & objOS.Version
	opf.WriteLine "        Service Pack: " & objOS.ServicePackMajorVersion & "." & objOS.ServicePackMinorVersion
Next
For Each objCPU in cpuinfo
	opf.WriteLine "     OS Architecture: " & objCPU.AddressWidth & " Bit"
Next
opf.WriteLine ""
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Processor Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objCPU in cpuinfo
	opf.WriteLine "Processor Make and Model: " & objCPU.Name
	opf.WriteLine "  Processor Architecture: " & objCPU.DataWidth & " Bit"
Next
opf.WriteLine ""
For Each objItem In memItems
	if strMemory <> "" then
		strMemory = strMemory & vbcrlf
	End If
	if memType(objItem.MemoryType).ToString = "0" then
		if objItem.Speed < 1033 then
		strMemtype = "DDR-2"
		else
			strMemtype = "DDR-3"
		End If
	else
		strMemtype = memType(objItem.MemoryType).ToString
	End If
	strMemory = strMemory &  "Slot" & i & ": " & (objItem.Capacity / 1048576) & " Mb" & " Type: " & strMemtype & " Speed: " & (objItem.Speed)
	i = i + 1
Next
installedModules = i - 1
Set memItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")
For Each objItem in memItems
	totalSlots = objItem.MemoryDevices
Next
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Memory Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
opf.WriteLine "Total Slots: " & totalSlots
opf.WriteLine " Free Slots: " & (totalSlots - installedModules)
opf.WriteLine ""
For Each objOS in osinfo
	opf.WriteLine "Available Phyisical Memory: " & left(objOS.FreePhysicalMemory / 1024,6) & " MB"
Next
opf.WriteLine ""
opf.WriteLine strMemory
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Network Connection Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objLAN in laninfo
	opf.WriteLine "     IP Address: " & objLAN.IPAddress(0)
	opf.WriteLine "    MAC Address: " & objLAN.MACAddress(0)
Next
opf.WriteLine ""
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "General Drive Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objDRV in drvinfo
	opf.WriteLine "Hard Drive Type: " & objDRV.InterfaceType
	opf.WriteLine "Hard Drive Model: " & objDRV.Model
	opf.WriteLine "Hard Drive Size: " & left(objDRV.Size / 1073741824,5) & " GB"
	opf.WriteLine "Number of Partitions: " & objDRV.Partitions
	opf.WriteLine ""
Next
opf.WriteLine ""
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Hard Drive Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objHD in hdinfo
	totalspce = left(objHD.size/1073741824,6)	
	freespace = left(objHD.freespace/1073741824,6)
	usedspace = left(totalspce-freespace,6)
	freeperct = left(freespace/totalspce*100,4)
	opf.WriteLine "   Drive ID: " & objHD.DeviceID
	opf.WriteLine " FileSystem: " & objHD.Filesystem
	opf.WriteLine "Total Space: " & totalspce & " GB"
	opf.WriteLine " Free Space: " & freespace & " GB"
	opf.writeLine " Free Perct: " & freeperct & "%"
	opf.WriteLine " Used Space: " & usedspace & " GB"
	opf.WriteLine ""
Next
opf.WriteLine ""
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Process Information"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
For Each objPRC in tskinfo
	opf.WriteLine "Name: " & left(objPRC.Name,12) & "        Location: " & objPRC.ExecutablePath
Next
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Installed Software"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
prgRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
prgReg.EnumKey HKEY_LOCAL_MACHINE, prgRegPath, prgSubKeyAry
For Each prgSubkey In prgSubKeyAry
	prgReg.GetStringValue HKEY_LOCAL_MACHINE, prgRegPath & prgSubkey, "DisplayName" , Name
	If Name <> "" Then 
		prgReg.GetStringValue HKEY_LOCAL_MACHINE, prgRegPath & prgSubkey,"DisplayVersion",Version
		opf.WriteLine Name&"____"& Version    
	End If 
Next
Dim prgRegWOW, prgSubKeyWOW, prgSubKeyAryWOW
Dim prgRegPathWOW
prgRegPathWOW = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
if prgReg.EnumKey(HKEY_LOCAL_MACHINE, prgRegPathWOW, prgSubKeyAryWOW) = 0 Then
opf.WriteLine "-----------------------------------------------"
opf.WriteLine "Installed Software (WOW)"
opf.WriteLine "-----------------------------------------------"
opf.WriteLine ""
prgReg.EnumKey HKEY_LOCAL_MACHINE, prgRegPathWOW, prgSubKeyAryWOW
	For Each prgSubkeyWOW In prgSubKeyAryWOW
		prgReg.GetStringValue HKEY_LOCAL_MACHINE, prgRegPathWOW & prgSubkeyWOW, "DisplayName" , Name
		If Name <> "" Then 
			prgReg.GetStringValue HKEY_LOCAL_MACHINE, prgRegPathWOW & prgSubkeyWOW,"DisplayVersion",Version
			opf.WriteLine Name&"____"& Version
		End If
	Next
End If
objShell.Run(notepad & pcName & ".txt")
Set biosSettings = nothing
Set cpuinfo = nothing
Set drvinfo = nothing
Set env = nothing
Set fso = nothing
Set hdinfo = nothing
Set instlldSoftware = nothing
Set laninfo = nothing
Set memItems = nothing
Set objShell = nothing
Set objWMIService = nothing
Set opf = nothing
Set osinfo = nothing
Set pcinfo = nothing
Set pcName = nothing
Set tskinfo = nothing