'#########################################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								  		      			   			##
'##																													  		      			   			##
'## December, 2014																									  		      			   			##
'##																													  		      		       			##
'## Version 1.0																										  		      			   			##
'##																													  		      			   			##
'## DESCRIPTION: Monitor windows services status and performance (CPU and memory utilization of each service)  			          			   			##
'##																													  		      			   			##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "Services.vbs" <HOST> <METRIC_STATE> <CIR_IDS> <CIRS> <USERNAME> <PASSWORD> <DOMAIN>    			##
'##																													  		      			   			##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "Services.vbs" "10.10.10.1" "2147,2148" "MpsSvc,wuauserv" "1,1,1,0,1,0,0" "user" "pwd" "domain" ##
'##																													              			   			##
'## README:	<METRIC_STATE> - (internal): only used by tellki default monitors. 						  			   										##
'##         1 - metric is on ; 0 - metric is off					              												  			   			##
'## 																												              			   			##
'## 		CIR_IDS - (internal): only used by tellki default monitors. Service unique cmdb ID.															##
'## 																												              			   			##
'## 		CIRS - (internal): only used by tellki default monitors. Service name.											                            ##
'## 																												              			   			##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this 			   			##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   			   			##
'##			pass them to the script.																						      			   			##
'## 																												              			   			##
'#########################################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 7 Then 
	CALL ShowError(3, 0) 
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const Status = "140:Status:9"
Const CPUUsage = "137:CPU Usage:6"
Const MemUsage = "45:Memory Usage:4"
Const MemUsagePerc = "134:% Memory Usage:6"
Const VirtualBytes = "50:Virtual Bytes:4"
Const ProcessorTimePerc = "56:% Processor Time:6"
Const ElapsedTime = "25:Elapsed Time:7"

'INPUTS
Dim Host, MetricState, CIR_IDS, CIRS, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
CIR_IDS = WScript.Arguments(2)
CIRS = WScript.Arguments(3)
Username = WScript.Arguments(4)
Password = WScript.Arguments(5)
Domain = WScript.Arguments(6)


Dim arrCIRS, arrCIRSIDs, arrMetrics
arrCIRS = Split(CIRS,",")
arrCIRSIDs = Split(CIR_IDS,",")
arrMetrics = Split(MetricState,",")

Dim objSWbemLocator, objSWbemServices, colItems, colItems2, colItems3
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, objItem, objItem2, FullUserName, strService, aux1, aux2, objItem3
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
	objSWbemServices.Security_.ImpersonationLevel = 3
	
	Set colItems = objSWbemServices.ExecQuery( _
			"SELECT TotalVirtualMemorySize,TotalVisibleMemorySize FROM Win32_OperatingSystem",,16) 
	For Each objItem in colItems
		aux1 = objItem.TotalVisibleMemorySize
	Next
	Dim OS, TellkiCPUUsage,TellkiMemoryUsage,TellkiMemoryUsagePerc,TellkiVirtualBytes,TellkiProcessorTimePerc, TellkiElapsedTime

	For Each strService In arrCIRS
		if IsEmpty(strService) = False Then
        
            Set colItems = objSWbemServices.ExecQuery( "SELECT State,ProcessId from Win32_Service Where Name = '" & strService & "'",,16)
				If colItems.Count <> 0 Then
				For Each objItem in colItems 
					TellkiCPUUsage = 0
					TellkiMemoryUsage = 0
					TellkiMemoryUsagePerc = 0
					TellkiVirtualBytes = 0
					TellkiProcessorTimePerc = 0
			
					'Status Process
					If objItem.State = "Running" Then
						OS = GetOSVersion(objSWbemServices)
						if OS >= 3000 Then
							Set colItems2 = objSWbemServices.ExecQuery("SELECT PercentProcessorTime,WorkingSet,VirtualBytes,PercentProcessorTime,ElapsedTime FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess="& objItem.ProcessId,,16)
						If colItems2.Count <> 0 Then
						
						'Service Status
						if arrMetrics(0) = 1 then CALL Output(Status,arrCIRSIDs(Counter),"1",strService)
						
						For Each objItem2 in colItems2 
							'CPU Usage
							if arrMetrics(1) = 1 then CALL Output(CPUUsage,arrCIRSIDs(Counter),FormatNumber(objItem2.PercentProcessorTime+TellkiCPUUSage),strService)
							'Memory Usage
							if arrMetrics(2) = 1 then CALL Output(MemUsage,arrCIRSIDs(Counter),FormatNumber((objItem2.WorkingSet/1048576)+TellkiMemoryUsage),strService)
							'% Memory Usage
							if arrMetrics(3) = 1 then CALL Output(MemUsagePerc,arrCIRSIDs(Counter),FormatNumber(((objItem2.WorkingSet/1024/aux1)*100)+TellkiMemoryUsagePerc),strService)
						    'Virtual Bytes
							if arrMetrics(4) = 1 then CALL Output(VirtualBytes,arrCIRSIDs(Counter),FormatNumber((objItem2.VirtualBytes/1048576)+TellkiVirtualBytes),strService)
						    'PercentProcessorTime
						 	if arrMetrics(5) = 1 then CALL Output(ProcessorTimePerc,arrCIRSIDs(Counter),FormatNumber(objItem2.PercentProcessorTime+TellkiProcessorTimePerc),strService)
						    'ElapsedTime
							if arrMetrics(6) = 1 then CALL Output(ElapsedTime,arrCIRSIDs(Counter),objItem2.ElapsedTime,strService)
						Next
					Else
						CALL getProcDetails(objSWbemServices, objItem.ProcessId, arrMetrics(1), arrMetrics(2), arrMetrics(3), arrMetrics(4), arrMetrics(5), arrMetrics(6), arrCIRSIDs(Counter), strService, aux1)
					End If
                
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
				Else
                    CALL Output(Status,arrCIRSIDs(Counter),"0",strService)
                
				End If
                If Err.number <> 0 Then
		           CALL ShowError(5, Host)
               
                End if
		    Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
        End if
	Counter = Counter + 1
	Next
	End If
If Err Then 
	CALL ShowError(1)
Else
	WScript.Quit(0)
End If

' Services
Function getProcDetails(SWbem,pid, m1, m2, m3, t, s, total)
	Dim sumPCT, musage, i, N1, D1, N2, D2, PercentProcessorTime, cpuPid, memusageMB
	Dim objInstance1, objInstance2
   	sumPCT=0
	musage=0
	For i = 1 to 5
		Set colItems = SWbem.ExecQuery("Select PercentProcessorTime,WorkingSet,TimeStamp_Sys100NS from Win32_PerfRawData_PerfProc_Process where IDProcess="&pid,,16)
		For Each objInstance1 in colItems
			musage=objInstance1.WorkingSet
			N1 = objInstance1.PercentProcessorTime
			D1 = objInstance1.TimeStamp_Sys100NS
		Next
		WScript.Sleep(1000)
		Set colItems = SWbem.ExecQuery("Select PercentProcessorTime,WorkingSet,TimeStamp_Sys100NS from Win32_PerfRawData_PerfProc_Process where IDProcess="&pid,,16)
		For Each objInstance2 in colItems
			musage=objInstance2.WorkingSet
			N2 = objInstance2.PercentProcessorTime
			D2 = objInstance2.TimeStamp_Sys100NS
		Next
			PercentProcessorTime = (((N2 - N1)/(D2-D1)))*100
			sumPCT=PercentProcessorTime+sumPCT
	Next
	'CPU Usage
	if m1 = 1 then CALL Output(CPUUsage,t,Round((sumPCT/10),2),strService)
	'Memory Usage
	if m2 = 1 then CALL Output(MemUsage,t,Round((musage/1048576)/1024,2),strService)
	'% Memory Usage
	if m3 = 1 then CALL Output(MemUsagePerc,t,FormatNumber((Round((musage/1048576)/1024,2)/total)*100),strService)
End Function

Function GetOSVersion(SWbem)
	Dim colItems, objItem
	Set colItems = SWbem.ExecQuery("select BuildVersion from Win32_WMISetting",,48)
	For Each objItem in colItems
		GetOSVersion = CInt(objItem.BuildVersion)
	Next
End Function

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(MetricID, CIR_ID, MetricValue, MetricObject)
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo CIR_ID & "|" & MetricID & "|" & MetricValue & "|" & MetricObject & "|" 		
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo CIR_ID & "|" & MetricID & "|" & MetricValue & "|" 		
		Else
			CALL ShowError(5, Host) 
		End If
	End If
End Sub


