'Check_sg_status.vbs
' Version 0.3 13th January 2016 - FirmwareVersion and LANStatus checks
' Version 0.4 14th January 2016 - Application Server service checks
' Version 0.5 14th January 2016 - FanStatus check
' Version 0.51 14th January 2016 - Temperature Check, Voltage Check, Boot Source, Disk Space, UCB Disk Status
' Version 0.52 15th January 2016 - Service Appliance service checks, Uptime
CONST plugin_version = "0.51 14/01/2016"
Dim objFields
Dim InputFile 
Dim FSO, oFile 
Dim strSQL
DIM strHQServer
Dim row
Dim SQLcn, SQLrs, i

strCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890"
intLength = 16

strCheck = Lcase(WScript.Arguments.Item(0))
strSGIPAddress = WScript.Arguments.Item(1)
dblOKThreshold = CDbl(WScript.Arguments.Item(2))
dblCRITICALThreshold = CDbl(WScript.Arguments.Item(3))

Select Case strCheck
    Case "status" Check_status strSGIPAddress
    Case "inventory" inventory
    Case "checkmem" check_mem strSGIPAddress, dblOKThreshold, dblCRITICALThreshold
	case "version" plugin_ver
	Case "firmware" firmware strSGIPAddress
	Case "lanstatus" LANStatus strSGIPAddress
	Case "services" Services strSGIPAddress
	Case "fanstatus" FanStatus strSGIPAddress
	Case "temperature" TemperatureStatus strSGIPAddress
	Case "voltage" VoltageStatus strSGIPAddress
	Case "bootsource" BootSource strSGIPAddress
	Case "diskspace" DiskSpace strSGIPAddress, dblOKThreshold, dblCRITICALThreshold
	Case "diskstatus" DiskStatus strSGIPAddress
	Case "ucbservice" UCBServiceStatus strSGIPAddress
	Case "uptime" UpTime strSGIPAddress
	Case "lastconnect" LastConnect strSGIPAddress
	Case "lastboot" LastBoot strSGIPAddress
	Case "lastdisconnect" LastDisconnect strSGIPAddress
	Case "initialise" initialise
	Case "initialize" initialise
    Case Else
		print_exit_code "No Command", 2, strCheckName
End Select

Sub Initialise
	wscript.echo "beginning setup of the ShoreTel plugin..."
	wscript.echo "reading ini files..."
	strPasscode = ReadIni("stconfig.ini","NRDP","token")
	
	If strPasscode <> "" Then
		wscript.echo "Existing Passcode = " & strPasscode
		wscript.echo "Delete this on the token line in the stconfig.ini to create a new one"
		strPasscode = GeneratePassword(strCharacters, intLength)
		wscript.echo strPasscode
	Else
		strConfigFile = ReadIni("stconfig.ini","parameters","strConfigFile")
		strSQL = "SELECT shoreware.switches.HostName FROM shoreware.switches WHERE shoreware.switches.type = 'SGHQ'"
		SQLExecute strSQL
		SQLrs.MoveFirst
		strHQServer=SQLrs.Fields(0)
		SQLrs.close
		SQLcn.close
		
		wscript.echo "Generating new passcode"
		strPasscode = GeneratePassword(strCharacters, intLength)
		wscript.echo strPasscode
		WriteIni ".\stconfig.ini","NRDP","token", strPasscode
		writeini strConfigFile, "NRDP", "token",strPasscode
		writeini strConfigFile, "NRDP", "Parent", "{Nagios server url}"
		writeini strConfigFile, "NRDP", "Hostname", strHQServer
		writeini strConfigFile, "api", "community_string",strPasscode
		writeini strConfigFile, "plugin directives","plugin_path","C:\\Program Files (x86)\\Nagios\\NCPA\\plugins\\"
		
	End if
	wscript.echo "Run check_st.vbs inventory 0 0 0  to generate or refresh the checks in NRDP configuration"
	intExitCode = 0
	strStatusText = strPasscode
	strCheckName = "Initialise"
	print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub plugin_ver()
	strCheckName = "Plugin Version"
	intExitCode = 0
	strStatusText = plugin_version
	print_exit_code strStatusText, intExitCode, strCheckName
End Sub


Sub LANStatus(strSGIPAddress)
	strCheckName = "LANStatus"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT shoreware.switches.type, shorewarestatus.heapswitchstatus.EthernetAddress, shorewarestatus.heapswitchstatus.ActiveEthernetLink, shorewarestatus.heapswitchstatus.EthernetLink1Status, shorewarestatus.heapswitchstatus.EthernetLink2Status FROM shorewarestatus.heapswitchstatus INNER JOIN shoreware.switches ON shorewarestatus.heapswitchstatus.SwitchID = shoreware.switches.SwitchID where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	
	strType=SQLrs.Fields(0)
	strEthernetAddress=SQLrs.Fields(1)
	intActiveEthernetLink=CInt(SQLrs.Fields(2))
	intEthernetLink1Status=CInt(SQLrs.Fields(3))
	intEthernetLink2Status=CInt(SQLrs.Fields(4))
	
	SQLrs.close
	SQLcn.close
	' Confirm that the type of device actually has a LAN interface we can query	
	Select Case strType
		CASE "SGHQ" blnValid = False
		CASE "SGDVS" blnValid = False
		CASE "VIRTUALPHONESWITCH" blnValid = False
		CASE "VIRTUALTRUNKSWITCH" blnValid = False
		CASE "VIRTUALSA" blnValid = False
		Case Else
			blnValid = True
	End Select
	' Supposing the device isnt something we can report the LAN status then bale out here
	If blnValid = False Then
		print_exit_code "Check is invalid for this model type", 1, strCheckName
	End If
	'	Otherwise proceed with returning the requested information.
	
	If intActiveEthernetLink = 1 Then
		strStatusText = "LAN1: " & PortSpeed(intEthernetLink1Status) &" : "& strEthernetAddress
	Else
		strStatusText = "LAN2: " & PortSpeed(intEthernetLink2Status) &" : "& strEthernetAddress
	End If
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub Services(strSGIPAddress)
	strCheckName="ShoreWare Services"
	If strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2, strCheckName
	End If
	'First confirm that we are attempting to check this is a server and not a ShoreGear switches
	strType = whatis(strSGIPAddress)
	Set SQLrs = Nothing
	Set SQLcn = Nothing
	Select Case strType
		Case "SGHQ"
		Case "SGDVS"
		Case Else
			intExitCode = 1
			strStatusText = "This is a " & strType & ", this check can only be made on ShoreTel HQ Server or DVS Server"
			print_exit_code strStatusText, intExitCode, strCheckName
	End Select
	strSQL = "SELECT heapservicestatus.DisplayName, heapservicestatus.ServiceName, heapservicestatus.CurrentState, heapswitchstatus.InternetAddress FROM heapswitchstatus LEFT OUTER JOIN heapservicestatus ON heapswitchstatus.ServerID = heapservicestatus.ServerID WHERE heapservicestatus.CurrentState <> 4 AND heapswitchstatus.InternetAddress = '"&strSGIPAddress&"';"
	SQLExecute strSQL
	
	' evaluate whether any records were returned.
	If SQLrs.EOF <> True Then
		' If there are records returned, then step through and get the names of the services, return a CRITICAL error and list the services that are not running.
		intExitCode = 2 
		strStatusText = "Services Not Running - "
		SQLrs.MoveFirst
		do until SQLrs.EOF
			strStatusText=strStatusText & SQLrs.Fields(0) & "; "
			SQLrs.MoveNext
		loop
	Else
		' If nothing was returned then everything is OK and we can return an OK status
		intExitCode=0
		strStatusText = "All Services are running"
	End If
	
	SQLrs.close
	SQLcn.close
	print_exit_code strStatusText, intExitCode, strCheckName
End Sub
Sub UCBServiceStatus(strSGIPAddress)
	strCheckName = "UCB Service Status"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT UCASE(shoreware.switches.type) as type, shoreware.switches.IPAddress, shorewarestatus.heapucbstatus.CMCAStatus, shorewarestatus.heapucbstatus.TMSStatus, shorewarestatus.heapucbstatus.SttsStatus, shorewarestatus.heapucbstatus.WebServerStatus, shorewarestatus.heapimstatus.IMStatus, shoreware.switches.HostName FROM shoreware.switches INNER JOIN shorewarestatus.heapucbstatus ON shoreware.switches.SwitchID = shorewarestatus.heapucbstatus.SwitchID LEFT OUTER JOIN shorewarestatus.heapimstatus ON shorewarestatus.heapucbstatus.ServerID = shorewarestatus.heapimstatus.ServerID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			CASE "SA100"
			CASE "SA400"
			CASE "VIRTUALSA"
		Case Else
			intExitCode=1
			strStatusText = "This check isn't valid for this kind of device"
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		If Not IsNull(SQLrs.Fields(2)) Then
			intCMCAStatus= 1
		Else
			intCMCAStatus= 0
		End If
		If Not IsNull(SQLrs.Fields(3)) Then
			intTMSStatus = 1
		Else
			intTMSStatus= 0
		End If
		If Not IsNull(SQLrs.Fields(4)) Then
			intSttsStatus= 1
		Else
			intSttsStatus= 0
		End If
		If Not IsNull(SQLrs.Fields(5)) Then
			intWebServerStatus= 1
		Else
			intWebServerStatus= 0
		End If
		If Not IsNull(SQLrs.Fields(6)) Then
			intIMStatus= 1
		Else
			intIMStatus= 0
		End If
		
		'wscript.echo "CMCAStatus " & intCMCAStatus 
		'wscript.echo "TMSStatus "  & intTMSStatus 
		'wscript.echo "SttsStatus "  & intSttsStatus 
		'wscript.echo "WebServerStatus "&intWebServerStatus
		'wscript.echo "IMStatus "&intIMStatus
		
		
		SQLrs.close
		SQLcn.close
		strStatusText = ""
		If intCMCAStatus + intTMSStatus + intSttsStatus + intWebServerStatus + intIMStatus <> 5 Then
			intExitCode = 1
			If intCMCAStatus <> 1 Then
				strStatusText = strStatusText & "CMCA Service Failed; "
			End IF
			If intTMSStatus <> 1 Then
				strStatusText = strStatusText & "TMS Service Failed; "
			End IF
			If intSttsStatus <> 1 Then
				strStatusText = strStatusText & "Softswitch Failed; "
			End IF
			If intWebServerStatus <> 1 Then
				strStatusText = strStatusText & "Web Server Failed; "
			End IF
			If intIMStatus <> 1 Then
				strStatusText = strStatusText & "IM Server Failed; "
			End IF
		Else
			intExitCode = 0
			strStatusText = "All Services Running"
		End If
	End If
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub DiskStatus(strSGIPAddress)
	strCheckName = "Disk Status"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT Ucase(shoreware.switches.type) as type, shoreware.switches.IPAddress, shorewarestatus.heapswitchhardwarestatus.Disk1Status, shorewarestatus.heapswitchhardwarestatus.Disk2Status FROM shorewarestatus.heapswitchhardwarestatus INNER JOIN shoreware.switches ON shorewarestatus.heapswitchhardwarestatus.SwitchID = shoreware.switches.SwitchID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			CASE "SA100"
			CASE "SA400"
		Case Else
			intExitCode=1
			strStatusText = "This kind of device doesn't have disks"
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		intDisk1Status=CInt(SQLrs.Fields(2))
		intDisk2Status=CInt(SQLrs.Fields(3))
		SQLrs.close
		SQLcn.close
		If intDisk1Status + intDisk2Status <> 2 then
			intExitCode=1
			strStatusText="Bad Disk"
		Else
			intExitCode=0
		End If
		If intDisk1Status = 1 Then
			strStatusText=strStatusText&"Disk1 OK"
		Else
			strStatusText=strStatusText&"Disk1 Failed"
		End If
		If intDisk2Status = 1 Then
			strStatusText=strStatusText&"; Disk2 OK"
		Else
			strStatusText=strStatusText&"; Disk2 Failed"
		End If
	End If
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub
Sub FanStatus(strSGIPAddress)
	strCheckName = "Fan Status"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT UCASE(shoreware.switches.type), shorewarestatus.heapswitchhardwarestatus.Fan FROM shorewarestatus.heapswitchhardwarestatus LEFT OUTER JOIN shoreware.switches ON shoreware.switches.SwitchID = shorewarestatus.heapswitchhardwarestatus.SwitchID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			CASE "SA100"
			CASE "SA400"
			CASE "SG120"
			CASE "SG220E1"
			CASE "SG24A"
			CASE "SG30"
			CASE "SG30BRI"
			CASE "SG40"
			CASE "SG50"
			CASE "SG50V"
			CASE "SG60"
			CASE "SG90"
			CASE "SG90BRI"
			CASE "SG90BRIV"
			CASE "SG90V"
			CASE "SGE1"
			CASE "SGE1K"
		Case Else
			intExitCode=1
			strStatusText = "This kind of device doesn't have a fan"
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		intFanStatus=CInt(SQLrs.Fields(1))
		SQLrs.close
		SQLcn.close
		Select Case intFanStatus
			CASE 1
				intExitCode=0
				strStatusText="OK"
			CASE 2
				intExitCode=2
				strStatusText="BAD"
			CASE 3
				intExitCode=1
				strStatusText="SLOW"
		Case Else
			intExitCode = 1
			strStatusText=intFanStatus &": Is NULL or Unknown"
		End Select
	End If
	
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub
Sub BootSource(strSGIPAddress)
	strCheckName = "Boot Source"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT UCASE(shoreware.switches.type), shorewarestatus.heapswitchstatus.BootSource FROM shorewarestatus.heapswitchstatus LEFT OUTER JOIN shoreware.switches ON shoreware.switches.SwitchID = shorewarestatus.heapswitchstatus.SwitchID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			'CASE "SA100"
			'CASE "SA400"
			CASE "SG120"
			CASE "SG220E1"
			CASE "SG24A"
			CASE "SG30"
			CASE "SG30BRI"
			CASE "SG40"
			CASE "SG50"
			CASE "SG50V"
			CASE "SG60"
			CASE "SG90"
			CASE "SG90BRI"
			CASE "SG90BRIV"
			CASE "SG90V"
			CASE "SGE1"
			CASE "SGE1K"
		Case Else
			intExitCode=1
			strStatusText = "This check cant be done on this kind of device"
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		intBootSource=CInt(SQLrs.Fields(1))
		SQLrs.close
		SQLcn.close
		Select Case intBootSource
			CASE 1
				intExitCode=1
				strStatusText="FTP Boot"
			CASE 2
				intExitCode=0
				strStatusText="FLASH"
			CASE 3
				intExitCode=2
				strStatusText="Firmware Update Failed"
		Case Else
			intExitCode = 1
			strStatusText=intBootSource &": Is NULL or Unknown"
		End Select
	End If
	
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub TemperatureStatus(strSGIPAddress)
	strCheckName = "Temperature Status"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT UCASE(shoreware.switches.type), shorewarestatus.heapswitchhardwarestatus.Temperature FROM shorewarestatus.heapswitchhardwarestatus LEFT OUTER JOIN shoreware.switches ON shoreware.switches.SwitchID = shorewarestatus.heapswitchhardwarestatus.SwitchID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			CASE "SA100"
			CASE "SA400"
			CASE "SG120"
			CASE "SG220E1"
			CASE "SG24A"
			CASE "SG30"
			CASE "SG30BRI"
			CASE "SG40"
			CASE "SG50"
			CASE "SG50V"
			CASE "SG60"
			CASE "SG90"
			CASE "SG90BRI"
			CASE "SG90BRIV"
			CASE "SG90V"
			CASE "SGE1"
			CASE "SGE1K"
		Case Else
			intExitCode=1
			strStatusText = "This kind of device cannot be checked for temperature status."
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		intTemperatureStatus=CInt(SQLrs.Fields(1))
		SQLrs.close
		SQLcn.close
		Select Case intTemperatureStatus
			CASE 1
				intExitCode=0
				strStatusText="Green"
			CASE 2
				intExitCode=1
				strStatusText="Yellow"
			CASE 3
				intExitCode=2
				strStatusText="Red"
			CASE 4
				intExitCode=2
				strStatusText="Blue"
			CASE 5
				intExitCode=2
				strStatusText="White"
		Case Else
			intExitCode = 1
			strStatusText=intTemperatureStatus &": Is NULL or Unknown"
		End Select
	End If
	
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub VoltageStatus(strSGIPAddress)
	strCheckName = "Voltage Status"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT UCASE(shoreware.switches.type), shorewarestatus.heapswitchhardwarestatus.VoltageStatus FROM shorewarestatus.heapswitchhardwarestatus LEFT OUTER JOIN shoreware.switches ON shoreware.switches.SwitchID = shorewarestatus.heapswitchhardwarestatus.SwitchID where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	Else
		SQLrs.MoveFirst
		strType=SQLrs.Fields(0)
		Select Case strType
			CASE "SA100"
			CASE "SA400"
			CASE "SG120"
			CASE "SG220E1"
			CASE "SG24A"
			CASE "SG30"
			CASE "SG30BRI"
			CASE "SG40"
			CASE "SG50"
			CASE "SG50V"
			CASE "SG60"
			CASE "SG90"
			CASE "SG90BRI"
			CASE "SG90BRIV"
			CASE "SG90V"
			CASE "SGE1"
			CASE "SGE1K"
		Case Else
			intExitCode=1
			strStatusText = "This kind of device cannot be checked for Voltage status."
			SQLrs.close
			SQLcn.close
			print_exit_code strStatusText, intExitCode, strCheckName
		End Select
		intVoltageStatus=CInt(SQLrs.Fields(1))
		SQLrs.close
		SQLcn.close
		Select Case intVoltageStatus
			CASE 0
				intExitCode=0
				strStatusText="OK"

			Case Else
			intExitCode = 2
			strStatusText="FAIL: "& intVoltageStatus &": Is NULL or Unknown"
		End Select
	End If
	
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub UpTime(strSGIPAddress)
	strCheckName = "Uptime"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT TIMESTAMPDIFF( SECOND, shorewarestatus.heapswitchstatus.LastBoot, UTC_TIMESTAMP ()) AS Uptime, heapswitchstatus.LastDisconnect FROM heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	strStatusText=SplitSec(CLng(SQLrs.Fields(0)))
	SQLrs.close
	SQLcn.close
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub LastBoot(strSGIPAddress)
	strCheckName = "Last Boot"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT heapswitchstatus.LastBoot FROM heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	strStatusText=ConvertUTCToLocal(SQLrs.Fields(0))
	SQLrs.close
	SQLcn.close
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub LastConnect(strSGIPAddress)
	strCheckName = "Last Connect"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT heapswitchstatus.LastConnect FROM heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	strStatusText=ConvertUTCToLocal(SQLrs.Fields(0))
	SQLrs.close
	SQLcn.close
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub LastDisConnect(strSGIPAddress)
	strCheckName = "Last Disconnect"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "SELECT heapswitchstatus.LastDisconnect FROM heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	If IsNull(SQLrs.Fields(0)) Then
		strStatusText="Yay! it has never disconnected!"
	Else
		strStatusText=ConvertUTCToLocal(SQLrs.Fields(0))
	End If
	
	SQLrs.close
	SQLcn.close
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Function SplitSec(pNumSec)
  Dim d, h, m, s
  Dim h1, m1

  d = int(pNumSec/86400)
  h1 = pNumSec - (d * 86400)
  h = int(h1/3600)
  m1 = h1 - (h * 3600)
  m = int(m1/60)
  s = m1 - (m * 60)

  SplitSec = cStr(d) & "d," & cStr(h) & "h," & cStr(m) & "m," & cStr(s) & "s"
End Function

Function ConvertUTCToLocal( varTime )
    Dim myObj, MyDate
    MyDate = CDate( varTime )
    Set myObj = CreateObject( "WbemScripting.SWbemDateTime" )
    myObj.Year = Year( MyDate )
    myObj.Month = Month( MyDate )
    myObj.Day = Day( MyDate )
    myObj.Hours = Hour( MyDate )
    myObj.Minutes = Minute( myDate )
    myObj.Seconds = Second( myDate )
    ConvertUTCToLocal = myObj.GetVarDate( True )
End Function

Sub Firmware(strSGIPAddress)
	strCheckName = "Firmware"
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2
	End If
	strSQL = "Select FirmwareVersion from shorewarestatus.heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	strStatusText=SQLrs.Fields(0)
	SQLrs.close
	SQLcn.close
	intExitCode=0
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub inventory()
    strSQL = "SELECT LCASE(shoreware.switches.HostName) AS HostName, shoreware.switches.IPAddress, LCASE(shoreware.switches.type) AS Type, shoreware.switches.SwitchName AS Description, shoreware.switches.EthernetAddress, shoreware.switches.SerialNum FROM shoreware.switches"
    strConfigFile = ReadIni("stconfig.ini","parameters","strConfigFile")
    SQLExecute strSQL
    'fieldCount = SQLrs.Fields.Count
        
    SQLrs.MoveFirst
    do until SQLrs.EOF
        '  WriteIni( myFilePath, mySection, myKey, myValue )'
        ' Output some handy inventory information to the INI file
        strComment="#"
        strHostname=SQLrs.Fields(0).value
        strIPAddress=SQLrs.Fields(1).value
        strType=SQLrs.Fields(2).value
        strDescription=SQLrs.Fields(3).value
        strMACAddress=SQLrs.Fields(4).value
        strSerialNo=SQLrs.Fields(5).value
        strRow = strHostname & vbTab & strIPAddress & vbTab & strType & vbTab & strMACAddress & vbTab & strSerialNo
		wscript.echo strRow
        WriteIni  strConfigFile, "Inventory", strRow, strSerialNo

       ' 'ts-tsc-wlg-sge1k-01|ShoreGear Status=/agent/plugin/check_sg_status.vbs/status/192.168.6.2/0/0'
       ' strCommand = strHostname &"|ShoreGear Status"
	   '	strParameters = "/agent/plugin/check_sg_status.vbs/status/" & strIPAddress &"0/0"
       ' WriteIni  strConfigFile, "passive checks", strCommand, strParameters
        Select Case UCASE(strType)
		CASE "SA100"
			Initialise_SA100 strHostname, strIPAddress, strConfigFile	
		CASE "SA400"
			Initialise_SA400 strHostname, strIPAddress, strConfigFile
		CASE "SG120"
			Initialise_SG120 strHostname, strIPAddress, strConfigFile
		CASE "SG220E1"
			Initialise_SG220E1 strHostname, strIPAddress, strConfigFile
		CASE "SG24A"
			Initialise_SG24A strHostname, strIPAddress, strConfigFile
		CASE "SG30"
			Initialise_SG30 strHostname, strIPAddress, strConfigFile
		CASE "SG30BRI"
			Initialise_SG30BRI strHostname, strIPAddress, strConfigFile
		CASE "SG40"
			Initialise_SG40 strHostname, strIPAddress, strConfigFile
		CASE "SG50"
			Initialise_SG50 strHostname, strIPAddress, strConfigFile
		CASE "SG50V"
			Initialise_SG50V strHostname, strIPAddress, strConfigFile
		CASE "SG60"
			Initialise_SG60 strHostname, strIPAddress, strConfigFile
		CASE "SG90"
			Initialise_SG90 strHostname, strIPAddress, strConfigFile
		CASE "SG90BRI"
			Initialise_SG90BRI strHostname, strIPAddress, strConfigFile
		CASE "SG90BRIV"
			Initialise_SG90BRIV strHostname, strIPAddress, strConfigFile
		CASE "SG90V"
			Initialise_SG90V strHostname, strIPAddress, strConfigFile
		CASE "SGE1"
			Initialise_SGE1 strHostname, strIPAddress, strConfigFile
		CASE "SGE1K"
			Initialise_SGE1K strHostname, strIPAddress, strConfigFile
		CASE "SGHQ"
			Initialise_SGHQ strHostname, strIPAddress, strConfigFile
			WriteIni  strConfigFile, "Inventory", strRow, "HQ Server"
			
		CASE "SGDVS"
			Initialise_SGDVS strHostname, strIPAddress, strConfigFile
			WriteIni  strConfigFile, "Inventory", strRow, "DVS Server"
		CASE "VIRTUALPHONESWITCH"
			Initialise_VIRTUALPHONESWITCH strHostname, strIPAddress, strConfigFile
		CASE "VIRTUALSA"
			Initialise_VIRTUALSA strHostname, strIPAddress, strConfigFile
		CASE "VIRTUALTRUNKSWITCH"
			Initialise_VIRTUALTRUNKSWITCH strHostname, strIPAddress, strConfigFile
		CASE Else
			wscript.echo strType & " :wasnt any of the recognised ShoreTel Model types"
		
		End Select
		strRow=""
        SQLrs.MoveNext
    loop
    SQLrs.close
    SQLcn.close
    print_exit_code "Done", 0, "Inventory Complete"

End Sub
Sub DiskSpace(strSGIPAddress, dblOKThreshold, dblCRITICALThreshold)
	strCheckName = "Disk Space"
	
	if strSGIPAddress = "" Then
		print_exit_code "No IP Address", 2, strCheckName
	End If
	
    If dblOKThreshold = "" Then
        print_exit_code "No OK Threshold Defined", 2, strCheckName
    End If
    
	If dblCRITICALThreshold = "" Then
        print_exit_code "No CRITICAL Threshold Defined", 2, strCheckName
    End If
	
	strType = whatis(strSGIPAddress)
	
	Select Case strType
			CASE "SA100"
			CASE "SA400"
			CASE "SG50V"
			CASE "SG90BRIV"
			CASE "SG90V"
			CASE "SGHQ"
			CASE "SGDVS"
	Case Else
			intExitCode=1
			strStatusText = "This kind of device cannot be checked for Disk Space."
			print_exit_code strStatusText, intExitCode, strCheckName
	End Select
	
	strSQL = "SELECT UCASE(type), InternetAddress, PerCentUsed, SpaceUsed, SpaceCapacity FROM ( SELECT shorewarestatus.heapswitchstatus.InternetAddress, shoreware.switches.type, shorewarestatus.heapswitchstatus.`Name`, TRUNCATE ((( shorewarestatus.heapmailservstatus.DiskspaceUsed / ( shorewarestatus.heapmailservstatus.DiskspaceFree + shorewarestatus.heapmailservstatus.DiskspaceUsed )) * 100 ), 2 ) AS PerCentUsed, shorewarestatus.heapmailservstatus.DiskspaceUsed AS SpaceUsed, ( shorewarestatus.heapmailservstatus.DiskspaceFree + shorewarestatus.heapmailservstatus.DiskspaceUsed ) AS SpaceCapacity FROM shorewarestatus.heapmailservstatus INNER JOIN shorewarestatus.heapserverstatus ON shorewarestatus.heapmailservstatus.ServerID = shorewarestatus.heapserverstatus.ServerID INNER JOIN shorewarestatus.heapswitchstatus ON shorewarestatus.heapserverstatus.SwitchID = shorewarestatus.heapswitchstatus.SwitchID INNER JOIN shoreware.switches ON shorewarestatus.heapswitchstatus.SwitchID = shoreware.switches.SwitchID UNION SELECT shorewarestatus.heapswitchstatus.InternetAddress, shoreware.switches.type, shorewarestatus.heapswitchstatus.`Name`, TRUNCATE ((( shorewarestatus.heapucbstatus.SpaceUsed / shorewarestatus.heapucbstatus.SpaceCapacity ) * 100 ), 2 ) AS PerCentUsed, shorewarestatus.heapucbstatus.SpaceUsed, shorewarestatus.heapucbstatus.SpaceCapacity FROM shorewarestatus.heapucbstatus INNER JOIN shorewarestatus.heapswitchstatus ON shorewarestatus.heapucbstatus.SwitchID = shorewarestatus.heapswitchstatus.SwitchID INNER JOIN shoreware.switches ON shorewarestatus.heapucbstatus.SwitchID = shoreware.switches.SwitchID ) AS DiskUsage WHERE InternetAddress = '"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	strType=SQLrs.Fields(0)
	strIPAddress=SQLrs.Fields(1)
	dblPercentUsed=Cdbl(SQLrs.Fields(2))
	lngSpacedUsed=Clng(SQLrs.Fields(3))
	lngSpaceCapacity=Clng(SQLrs.Fields(4))
	SQLrs.close
	SQLcn.close
	If dblPercentUsed <= dblOKThreshold Then  ' OK Status - life is good here...
        strStatusText = dblPercentUsed &_
			"%|'percent'="&dblPercentUsed&"%,"&_
			"Used="&lngSpaceUsed&_
			",Capacity="&lngSpaceCapacity&";"&_
			dblOKThreshold &";"&dblCRITICALThreshold&";"
        intExitCode = 0
        ElseIf dblPercentUsed >= dblCRITICALThreshold Then  ' CRITICAL Status - DefCon 1; I hope you packed a lunch...
            strStatusText = dblPercentUsed & "%|'percent'="&dblPercentUsed&"%,"&"Used="&lngSpaceUsed&",Capacity="&lngSpaceCapacity&";"& dblOKThreshold &";"&dblCRITICALThreshold&";"
            intExitCode = 2
        Else                                                    ' WARNING Status - not *yet* cause for alarm... but our eyes are upon ye...
        strStatusText = dblPercentUsed & "%|'percent'="&dblPercentUsed&"%,"&"Used="&lngSpaceUsed&",Capacity="&lngSpaceCapacity&";"& dblOKThreshold &";"&dblCRITICALThreshold&";"
        intExitCode = 1
    End If
    	
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub Check_mem(strSGIPAddress, dblOKThreshold, dblCRITICALThreshold)
    strCheckName = "Memory"
    if strSGIPAddress = "" Then
        print_exit_code "No IP Address", 2, strCheckName
    End If

    If dblOKThreshold = "" Then
        print_exit_code "No OK Threshold Defined", 2, strCheckName
    End If
    If dblCRITICALThreshold = "" Then
        print_exit_code "No CRITICAL Threshold Defined", 2, strCheckName
    End If

    strSQL = "Select truncate((memused/MemTotal)*100,2) as MemUsedPC from heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"


    SQLExecute strSQL
    SQLrs.MoveFirst
    row=CDbl(SQLrs.Fields(0))
    SQLrs.close
    SQLcn.close

    If row <= dblOKThreshold Then  ' OK Status - life is good here...
        strStatusText = row & "%|'percent'="&row&"%;"& dblOKThreshold &";"&dblCRITICALThreshold&";"
        intExitCode = 0
        ElseIf row >= dblCRITICALThreshold Then  ' CRITICAL Status - DefCon 1; I hope you packed a lunch...
            strStatusText = row & "%|'percent'="&row&"%;"& dblOKThreshold &";"&dblCRITICALThreshold&";"
            intExitCode = 2
        Else                                                    ' WARNING Status - not *yet* cause for alarm... but our eyes are upon ye...
        strStatusText = row & "%|'percent'="&row&"%;"& dblOKThreshold &";"&dblCRITICALThreshold&";"
        intExitCode = 1
    End If
    print_exit_code strStatusText, intExitCode, strCheckName


End Sub

Sub Check_status(strSGIPAddress)
    strCheckName = "Status"
	if strSGIPAddress = "" Then
        print_exit_code "No IP Address", 2, strCheckName
    End If

    strSQL = "SELECT shorewarestatus.heapswitchstatus.Service from heapswitchstatus where InternetAddress='"&strSGIPAddress&"';"
    intDChannelDown = CInt(ReadIni("stConfig.ini","choices","intDChannelDown"))

    SQLExecute strSQL
    SQLrs.MoveFirst
    row=CInt(SQLrs.Fields(0))
    SQLrs.close
    SQLcn.close

    intExitCode = 1

    SELECT Case Row
        CASE 0 strStatusText="Unknown"
            intExitCode=2
        CASE 110 strStatusText="In Service"
            intExitCode = 0
        CASE 210 strStatusText="Firmware Update Available"
        CASE 220 strStatusText="Restart Pending"
        CASE 230 strStatusText="Upgrade In Progress"
        CASE 325 strStatusText="Port Out Of Service"
        CASE 330 strStatusText="Hunt Group Out Of Service"
            intExitCode=0
        CASE 340 strStatusText="Sip Trunks Out Of Service"
            intExitCode=2
        CASE 343 strStatusText="Sip Trunks Out Of Service Operational"
            intExitCode=2
        CASE 345 strStatusText="Sip Trunks Out Of Service Administrative"
        CASE 347 strStatusText="Ports Out Of Service Busy"
        CASE 349 strStatusText="Ports Out Of Service Not Aquired"
        CASE 350 strStatusText="Soft Phones Out Of Service"
        CASE 354 strStatusText="Soft Phones Out Of Service Operational"
        CASE 357 strStatusText="Soft Phones Out Of Service Administrative"
        CASE 360 strStatusText="Ip Phones Out Of Service"
            intExitCode=0
        CASE 364 strStatusText="Ip Phones Out Of Service Operational"
            intExitCode=0
        CASE 367 strStatusText="Ip Phones Out Of Service Administrative"
            intExitCode=0
        CASE 370 strStatusText="Some Ports Out Of Service"
        CASE 380 strStatusText="All Ports Out Of Service"
        CASE 383 strStatusText="Platform Version Mismatch"
            intExitCode=2
        CASE 387 strStatusText="Booting From Ftp"
            intExitCode=2
        CASE 390 strStatusText="Firmware Mismatch"
            intExitCode=2
        CASE 410 strStatusText="Configuration Mismatch"
        CASE 430 strStatusText="D Channel Down"
            intExitCode = intDChannelDown
        CASE 440 strStatusText="Fan Failure"
            intExitCode=2
        CASE 450 strStatusText="Temperature Failure"
            intExitCode=2
        CASE 460 strStatusText="Voltage Failure"
            intExitCode=2
        CASE 470 strStatusText="Firmware Update Failure"
            intExitCode=2
        CASE 480 strStatusText="Disk Failure"
            intExitCode=2
        CASE 490 strStatusText="Lost Communication"
            intExitCode=2
        CASE ELSE strStatusText="Unknown"
            intExitCode=2
    End Select
    print_exit_code strStatusText, intExitCode, strCheckName
End Sub

Sub print_exit_code(strStatusText, intExitCode, strCheckName)
    If intExitCode = 0 Then  ' OK Status - life is good here...
        wscript.echo strCheckName & " OK: "& strStatusText & "|" & strStatusText
        WScript.Quit(0)
    End If
    If intExitCode = 2 Then  ' CRITICAL Status - DefCon 1; I hope you packed a lunch...
        wscript.echo strCheckName & " CRITICAL: "& strStatusText & "|" & strStatusText
        WScript.Quit(2)
    End If

    If intExitCode = 1 Then ' WARNING Status - not *yet* cause for alarm... but our eyes are upon ye...
        wscript.echo strCheckName & " WARNING: "& strStatusText & "|" & strStatusText
        WScript.Quit(1)
    End If

end Sub

Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim(CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)&"\\"& myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["

                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )

                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "

                            End If

                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )

                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
		strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
		WScript.Echo strPath
        Wscript.Quit 1
    End If
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue )
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni, wshShell
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
    Set wshShell = CreateObject( "WScript.Shell" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
    Set wshShell  = Nothing
End sub
    
Sub SQLExecute (strSQL)
    ' This subroutine connects to the database and runs the SQL query and retrieves a recordset containing the results
    ' after you are done with the connection and recordset you need to close it elsewhere.
    ' Arguments:
    ' strSQL      [string]  the SQL Select statement that you are going to run.
    

    set SQLcn = CreateObject("ADODB.Connection")
    set SQLrs = CreateObject("ADODB.Recordset")

    ' Read in some parameters from the INI file about the database that we are going to connect to.'
    strODBCDriver = ReadIni("stConfig.ini","parameters","odbc")
    strHQServer = ReadIni("stConfig.ini","parameters","shoretelhq")

    SQLcn.connectionString = "Driver={"&strODBCDriver&"}; Server="& strHQServer &"; Port=4308; " & _
                       "Database=shorewarestatus;User=shoreadmin;" & _
                       "Password=passwordshoreadmin;"&_
                       "option=2;"&_
                       "AllowUserVariables=True;"
    ' Open the connection to the database
    SQLcn.open
    ' Execute the query
    SQLrs.open strSQL, SQLcn, 0
	If SQLrs.EOF = True Then
		intExitCode = 2 
		strStatusText = "Device not found - "
	End If
End Sub

Function PortSpeed(intPortSpeed)
	Select Case intPortSpeed
		CASE 0 PortSpeed= "Unknown"
		CASE 1 PortSpeed= "10M Half Duplex (Auto)"
		CASE 2 PortSpeed= "10M Full Duplex (Auto)"
		CASE 3 PortSpeed= "100M Half Duplex (Auto)"
		CASE 4 PortSpeed= "100M Full Duplex (Auto)"
		CASE 5 PortSpeed= "10M Half Duplex (Link)"
		CASE 6 PortSpeed= "100M Half Duplex (Link)"
		CASE 7 PortSpeed= "10M Half Duplex (Manual)"
		CASE 8 PortSpeed= "10M Full Duplex (Manual)"
		CASE 9 PortSpeed= "100M Half Duplex (Manual)"
		CASE 10 PortSpeed= "100M Full Duplex (Manual)"
		CASE 11 PortSpeed= "1G Half Duplex (Auto)"
		CASE 12 PortSpeed= "1G Full Duplex (Auto)"
		CASE 13 PortSpeed= "1G Half Duplex (Link)"
		CASE 14 PortSpeed= "1G Half Duplex (Manual)"
		CASE 15 PortSpeed= "1G Full Duplex (Manual)"
		CASE 16 PortSpeed= "10G Full Duplex (Auto)"
		CASE 17 PortSpeed= "10G Full Duplex (Manual)"
		Case Else PortSpeed = "Unknown"
	End Select
End Function

Function whatis(strSGIPAddress)
	strSQL = "Select shoreware.switches.type from shoreware.switches where shoreware.switches.IPAddress='"&strSGIPAddress&"';"
	SQLExecute strSQL
	SQLrs.MoveFirst
	whatis=SQLrs.Fields(0)
	SQLrs.close
	SQLcn.close
	set SQLcn = Nothing
    set SQLrs = Nothing
End Function

Function GeneratePassword(strCharacters, intLength)
     Randomize
     Dim strS, intI
     For intI = 1 To intLength
          strS = strS + Mid(strCharacters, Int(Rnd() * Len(strCharacters))+1, 1)
     Next
     GeneratePassword=strS
End Function

SUB Initialise_SA100(strHostName, strSGIPAddress, strConfigFile)
	make_check_ucbservice strHostName, strSGIPAddress, strConfigFile
	make_check_diskstatus strHostName, strSGIPAddress, strConfigFile
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
	
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
	
	
	
End Sub
SUB Initialise_SA400(strHostName, strSGIPAddress, strConfigFile)
	make_check_ucbservice strHostName, strSGIPAddress, strConfigFile
	make_check_diskstatus strHostName, strSGIPAddress, strConfigFile
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
	
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG120(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG220E1(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG24A(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG30(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG30BRI(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG40(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG50(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG50V(strHostName, strSGIPAddress, strConfigFile)
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
' General checks that can be done on this model,
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG60(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG90(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG90BRI(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG90BRIV(strHostName, strSGIPAddress, strConfigFile)
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
' General checks that can be done on this model, 
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SG90V(strHostName, strSGIPAddress, strConfigFile)
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SGE1(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SGE1K(strHostName, strSGIPAddress, strConfigFile)
' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_bootsource strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_voltage strHostName, strSGIPAddress, strConfigFile
	make_check_checkmem strHostName, strSGIPAddress	, strConfigFile
	make_check_fanstatus strHostName, strSGIPAddress, strConfigFile
	make_check_temperature strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_SGHQ(strHostName, strSGIPAddress, strConfigFile)
	'Model Specific checks
	make_check_services strHostName, strSGIPAddress, strConfigFile
	make_plugin_version strHostName, strSGIPAddress, strConfigFile
	make_CPU strHostName, strSGIPAddress, strConfigFile
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
	make_memory strHostName, strSGIPAddress, strConfigFile
	
	' General checks that can be done on this model
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	
End Sub
SUB Initialise_SGDVS(strHostName, strSGIPAddress, strConfigFile)
End Sub
SUB Initialise_VIRTUALPHONESWITCH(strHostName, strSGIPAddress, strConfigFile)
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_VIRTUALSA(strHostName, strSGIPAddress, strConfigFile)
	make_check_ucbservice strHostName, strSGIPAddress, strConfigFile
	make_check_diskspace strHostName, strSGIPAddress, strConfigFile
End Sub
SUB Initialise_VIRTUALTRUNKSWITCH(strHostName, strSGIPAddress, strConfigFile)
	make_check_status strHostName, strSGIPAddress, strConfigFile
	make_check_firmware strHostName, strSGIPAddress, strConfigFile
	make_check_lastboot strHostName, strSGIPAddress, strConfigFile
	make_check_lastconnect strHostName, strSGIPAddress, strConfigFile
	make_check_lastdisconnect strHostName, strSGIPAddress, strConfigFile
	make_check_uptime strHostName, strSGIPAddress, strConfigFile
	make_check_lanstatus strHostName, strSGIPAddress, strConfigFile
End Sub

Sub make_check_status(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "ShoreGear Status"
	strCommandName = "Status"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_checkmem(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "ShoreGear Memory Usage"
	strCommandName = "checkmem"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"80/85"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_firmware(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Firmware"
	strCommandName = "firmware"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_lanstatus(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "LANStatus"
	strCommandName = "lanstatus"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_services(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "ShoreWare Services"
	strCommandName = "services"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_fanstatus(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Fan Status"
	strCommandName = "fanstatus"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_temperature(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Temperature Status"
	strCommandName = "temperature"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_voltage(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Voltage Status"
	strCommandName = "voltage"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_bootsource(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Boot Source"
	strCommandName = "bootsource"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_diskspace(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Disk Space"
	strCommandName = "diskspace"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"90/95"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_diskstatus(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Disk Status"
	strCommandName = "diskstatus"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_ucbservice(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "UCB Service Status"
	strCommandName = "ucbservice"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_uptime(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Uptime"
	strCommandName = "uptime"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_lastconnect(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Last Connect"
	strCommandName = "lastconnect"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_lastboot(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Last Boot"
	strCommandName = "lastboot"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_check_lastdisconnect(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Last Disconnect"
	strCommandName = "lastdisconnect"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_plugin_version(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Plugin Version"
	strCommandName = "version"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/agent/plugin/check_sg_status.vbs/"&strCommandName&"/" & strSGIPAddress &"/"&"0/0"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_memory(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "Memory"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/memory/virtual/percent?warning=80&critical=92&check=true"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub

Sub make_CPU(strHostName, strSGIPAddress, strConfigFile)
	strCheckName = "CPU"
	strCommand = strHostname &"|"&strCheckName
	strParameters = "/cpu/percent?warning=90&critical=95&check=true"
    WriteIni  strConfigFile, "passive checks", strCommand, strParameters
    strCheckName = ""
	strCommand = ""
	strParameters = ""
End Sub


