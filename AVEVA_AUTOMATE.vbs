'============================================================================================
'Instructions:
'			For arguments:
'			1. To pass an argument use format "-argumentName argumentValue"
'			2. Only passing the "argumentValue" will save the argumentValue as UNKNOWARG
'			3. Shuffling of arguments is possible, do not have to maintain sequence
'			4. For running macro, module can be defined with "-module"(Ex. -module tags), but not mandatory, process will run in Design module as default
'			5. Mandatory arguments: PDMSCommand, projectCode, mdb
'Need Changes:
'			1. Change path for 'automatedServiceFile' for RVM export process
'			2. Change path for process monitor vbs
'			3. Change path of log file(RVM)
'			3. Change path of log file(MACRO)
'============================================================================================

Option Explicit
Dim objFSO, objShell, objEnvironment, dateToday, projectCode, automatedServices, logLine, _
	objFile, logFile, arguments, objWMIService, objXMLDoc, name, path, mdbName, timeNow, _
	project, logLinewithTime, description, strCommand, strTitle, colItems, selfPID, objItem, _
	startTime, endTime, macro, arglist, pdmsCommand, modul, product, mainarglist, scriptPath, _
	scriptFolder, scrptFile

startTime = Timer()
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objEnvironment = objShell.Environment("PROCESS")
Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
scriptPath = Wscript.ScriptFullName
Set scrptFile = objFSO.GetFile(scriptPath)
scriptFolder = objFSO.GetParentFolderName(scrptFile) 
Set arguments = Wscript.Arguments
Set arglist = CreateObject("Scripting.Dictionary")
Set mainarglist = CreateObject("Scripting.Dictionary")
arglist.CompareMode = VBTextCompare
mainarglist.CompareMode = VBTextCompare
dateToday = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)

Call service

Sub cleanPdms
	On Error Resume Next
	objFSO.DeleteFile "X:\PDMSUSER\srvPlantPLMUser1\E3D2\session_srvplantplmuser1" 
	On Error Goto 0
End Sub

Sub service
	Dim numOfargs, i, j, n, k, argkey, argval, arg, sizebef, sizeaft, mdblist
	Set mdblist = CreateObject("Scripting.Dictionary")
	mainarglist.Add "-PDMScommand", ""
	mainarglist.Add "-projectCode", ""
	mainarglist.Add "-mdb", ""
	sizebef = mainarglist.Count
	
	Set mdblist = readconfig("C:\ScheludedTasks\NwdExport\mdbconfig.txt")
	If mdblist.Exists(mdbName) Then
		Wscript.Quit
	End If
	'=====================================================================================
	'logFile = "X:\PDMSUSER\jklhasanma\NWD\ORIGINALS\MODIFIED070921\TESTLOGS\TEST_RVM.txt"
	'---- Get all arguments and load into dictionary 'arglist'
	numOfargs = arguments.Count
	i = 0
	n = 1
	For k = 0 to (numOfargs-1)
		If i > numOfargs-1 Then
			Exit For
		End IF
		If InStr(1, arguments(i), "-") = 1 Then
			argkey = UCase(arguments(i))
			On Error Resume Next
			argval = arguments(i + 1)
			If Err.Number <> 0 Then
				Err.Clear
				argval = ""
			End If
			If InStr(1, argval, "-") = 1 Then
				argval = ""
			Else
				i = i + 1
			End If
			arglist.Add argkey, argval
		Else
			argkey = "-UNKNOWNARG" & n
			argval = arguments(i)
			arglist.Add argkey, argval
			n = n + 1
		End If
		i = i + 1
	Next
	'---- Load mandatory arguments into dictionary 'mainarglist'
	For Each arg in mainarglist.keys()
		If arglist.Exists(UCase(arg)) And Len(arglist(UCase(arg))) > 0 Then
			mainarglist(arg) = arglist(UCase(arg))
		Else
			mainarglist.Remove arg
		End If
	Next
	'---- Check whether all mandatory arguments are available, otherwise quite the process
	sizeaft = mainarglist.Count
	If (sizeaft < sizebef) Then
		writeToLog logFile, "Mandatory arguments are not available", 8
		Wscript.Quit
	Else
		pdmsCommand = mainarglist("-PDMScommand")
		projectCode = mainarglist("-projectCode")
		mdbName = mainarglist("-mdb")
	End If
	
	'=====================================================================================
	'---- Get self process ID
	strTitle   = Rnd( Second( Now ) ) & " " & FormatDateTime( Now, vbShortTime )
	strCommand = "cmd.exe /k title " & strTitle
	objShell.Run strCommand, 7, False
	On Error Resume Next
	Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_Process WHERE CommandLine LIKE '%cmd.exe% /k title " & strTitle & "'" )
	For Each objItem In colItems
		selfPID = objItem.ParentProcessId
		On Error Resume Next
		' Terminate the spawned process
		objItem.Terminate
	Next
	'---- Run service based on pdms command (RVM/MAC)
	runService(pdmsCommand)
	
	endTime = Timer()
	writeToLog logFile, "- time elapsed " & FormatNumber(endTime - startTime, 5), 8
End Sub
Function runService(pdmsCommand)
	Dim automatedServiceFile, automatedServiceWithArgs, nwdFilePath, j, n, m, numOfargs, f, mdl, obj
	'---- Run process for RVM Export
	If pdmsCommand = "RVM" Then
		'::[CHANGE]-PATH CHANGE 
		'automatedServiceFile = "C:\ScheludedTasks\NwdExport\VALIT_AutomatedServices.mac"
		automatedServiceFile = "X:\PDMSUSER\jklhasanma\NWD\PROD\VALIT_AutomatedServices.mac"
		nwdFilePath = arglist("-NWDPATH")
		mainarglist.Add "-NWDPath", nwdFilePath
		'::[CHANGE]-PATH CHANGE
		logFile = "X:\PDMSUSER\jklhasanma\NWD\NavisService\Logs\" & mdbName & "_"  & selfPID &  "_PlantExport_" & projectCode & "_RVM.txt"
		argInfoToLog()
		If Len(nwdFilePath) > 0 And objFSO.FileExists(automatedServiceFile) Then
			automatedServiceWithArgs = automatedServiceFile & " " & nwdFilePath & " " & logFile & ""
			automatedServices = runAutomatedServices("model", automatedServiceWithArgs)
		Else 
			writeToLog logFile, "NWD path not valid", 8
			Wscript.Quit
		End IF
	'---- Run process for running macro
	ElseIf pdmsCommand = "MAC" Then
		automatedServiceFile = arglist("-MACRO")
		mdl = "model"
		If arglist.Exists("-MODULE") Then
			mdl = arglist("-MODULE")
			mainarglist.Add "-module", mdl
		End IF
		'::[CHANGE]-PATH CHANGE 
		logFile = "X:\PDMSUSER\jklhasanma\NWD\NavisService\Logs\" & mdbName & "_"  & selfPID &  "_RunMacro_" & projectCode & ".txt"
		n = 1
		automatedServiceWithArgs = automatedServiceFile
		'---- Get all macro argument
		For j = 0 to arglist.Count - 1
			If (arglist.keys()(j) = "-MACRO") Then
				mainarglist.Add "-macro", arglist("-MACRO")
				For m = j + 1 to arglist.Count - 1
					If InStr(1, arglist.keys()(m), "-UNKNOWNARG") = 1 Then
						mainarglist.Add "-macroArg" & n, arglist.items()(m)
						arglist.key(arglist.keys()(m)) = "-macroArg" & n
						automatedServiceWithArgs = automatedServiceWithArgs + " " + arglist.items()(m)
						n = n + 1
					End IF
				Next
			End IF
		Next
		' For Each obj in mainarglist.keys
			' writeToLog logFile, obj & "|" & mainarglist(obj) , 8
		' Next
		argInfoToLog()
		If (Len(automatedServiceFile) > 0) And objFSO.FileExists(automatedServiceFile) Then
			'Wscript.Quit
			automatedServices = runAutomatedServices(mdl, automatedServiceWithArgs)
		Else
			writeToLog logFile, "Macro path not valid", 8
			Wscript.Quit
		End IF
	Else
		Wscript.Quit
	End If
End Function
Function argInfoToLog()
	Dim j, args, elem
	args = "Arguments: "
	writeToLog logFile, "-processID", 8
	writeToLog logFile, selfPID, 8
	For Each elem in arglist.keys
		If not (mainarglist.Exists(elem)) Then
			mainarglist.Add elem, arglist(elem)
		End If
		
	Next
	For Each elem in mainarglist.keys
		writeToLog logFile, elem, 8
		writeToLog logFile, mainarglist(elem), 8
		args = args & " " & elem & " " & mainarglist(elem)
	Next
	writeToLog logFile, args, 8
End Function
Function e3dmodule(e3dmod)
	'Dim modul
	If UCase(e3dmod) = "MODEL" Or UCase(e3dmod) = "DESIGN" Then
		modul = "C:\Program Files (x86)\AVEVA\Everything3D2.10\des.exe"
		product = "E3D"
	ElseIf UCase(e3dmod) = "DRAW" Then
		modul = "C:\Program Files (x86)\AVEVA\Everything3D2.10\draw.exe"
		product = "E3D"
	ElseIf UCase(e3dmod) = "TAGS" Then
		modul = "C:\Program Files (x86)\AVEVA\Engineering14.1.1\tags.exe"
		product = "ENGINEERING"
	ElseIf UCase(e3dmod) = "CATALOGUE" Then
		modul = "C:\Program Files (x86)\AVEVA\Everything3D2.10\cata.exe"
		product = "CATALOGUE"
	End If
	'e3dmodule =  modul
End Function

Function runAutomatedServices(mdl, automatedServiceWithArgs)
	Dim automatedServiceFile, startFile, commandToRun, objTextFile, millId, username, password, perfData, availableMemory, _
		totalmemory, pdmsProcesses, pid, err, credentials, cred, credentialSuccess, numOfargs, _
		i, arglog, sec, setproduct
	writeToLog logFile, "_______________________________________________________automatedServiceFile", 8

		'--Credentials-----------------------------------------
		Set credentials = CreateObject("Scripting.Dictionary")
		credentials.Add "SYSTEM", "XXXXXX"
		credentials.Add projectCode, "BOILER"		
		credentials.Add "ONE", "VALMET"
		'------------------------------------------------------
			credentialSuccess = false
			Do
				Set pdmsProcesses = objWMIService.ExecQuery ("Select * From Win32_Process Where Name = 'des.exe' Or Name = 'mon.exe' Or Name = 'adm.exe'")
				Set perfData = objWMIService.ExecQuery ("Select * From Win32_OperatingSystem")
				availablememory = perfData.Itemindex(0).FreePhysicalMemory / (1024 * 1024)
				totalmemory = perfData.Itemindex(0).TotalVisibleMemorySize / (1024 * 1024)
				
				
				If availableMemory / totalmemory > 0.3 And availableMemory > 7 And pdmsProcesses.Count < 9 Then
					Call cleanPdms					
					setenvvariables(projectCode)
					For Each cred in credentials
						writeToLog logFile, "_______________________________________________________automatedServiceFile", 8
						username = cred
						password = credentials(cred)
						e3dmodule(mdl)
						objEnvironment("AVEVA_PRODUCT") = product
						commandToRun = """" & modul & """ PROD "& product &" init X:\AVEVA\STARTUP\E3D2\master_init.init" & " TTY " & projectCode & " " & username & "/" & password & " /" & mdbName & " $M/" & automatedServiceWithArgs & ""
						 writeToLog logFile, commandToRun, 8
						 'Wscript.Quit
						err = rundesprocess(commandToRun)
						If (err = 0) Then
							credentialSuccess = true
							Exit For
						End IF
						
					Next
					If not(credentialSuccess) Then
						writeToLog logFile, "Invalid Username or Password, unable to run export process", 8
					End If
					writeToLog logFile, "Process finished.", 8
					exit do
				Else
					writeToLog logFile, "Running low on memory or more than 3 PDMS processes running, waiting...", 8
					WScript.Sleep 60000
				End If
			Loop Until availableMemory / totalmemory > 0.3 And availableMemory > 7 And pdmsProcesses.Count < 9
End Function

Function rundesprocess(commandline)
	Dim proc, strOutput, errexist, loglines, var, duration, processmonitorpath
	writeToLog logFile, "Executed command|" & commandline, 8
	Set proc = objShell.Exec(commandline)
	rundesprocess = proc.ProcessID
	
	'::[CHANGE]-PATH CHANGE
	processmonitorpath = scriptFolder + "\" + "processmonitor.vbs"
	If not objFSO.FileExists(processmonitorpath) Then
		processmonitorpath = "X:\PDMSUSER\jklhasanma\NWD\PROD\processmonitor.vbs"
	End If
	
	On Error Resume Next
	objShell.Run processmonitorpath + " " + Cstr(rundesprocess) + " " + "480"
	If Err.Number <> 0 Then
		writeToLog logFile, "Unable to run Process Monitor", 8	
		Err.Clear
	End If
	'objShell.Run "C:\ScheludedTasks\NwdExport\processmonitor.vbs " + Cstr(rundesprocess) + " " + "150"
	writeToLog logFile, "Process Monitor: " + processmonitorpath + " " + Cstr(rundesprocess) + " " + "480", 8
	'objShell.Run ProcessControl + ProcessID + runtime duration(in minute)
	
	Do Until proc.StdOut.AtEndOfStream
		var = proc.StdOut.ReadLine
		loglines = loglines & vbNewLine & var
		writeToLog logFile, var , 8		
	Loop
	errexist = checkerror(loglines, "Invalid Username or Password")
	
	rundesprocess = errexist
End Function
Function checkerror(logContent, substr)
	Dim filecontent, alllines, i
	checkerror = False
	alllines = Split(logContent, vbNewline)
	For i = UBound(alllines) To 1 Step -1
		If InStr(alllines(i), substr) Then
			checkerror = True
			Exit For
		End If
	Next
End Function
Function readconfig(configfilepath)
	Dim strLine, configdict, configfile
	Set configdict = CreateObject("Scripting.Dictionary")
	If not objFSO.FileExists(configfilepath) Then
		'Wscript.Echo "File not Exists"
		Set readconfig = configdict
		Exit Function
	End If
	On Error Resume Next
	Set configfile = objFSO.OpenTextFile(configfilepath, 1)
	If Err.Number <> 0 Then
		Set readconfig = configdict
		Exit Function
	End If
	While Not configfile.AtEndOfStream
		strLine = configfile.ReadLine
		configdict.Add strLine, ""
	Wend
	configfile.Close
	Set configfile = Nothing
	Set readconfig = configdict
End Function
Sub writeToLog(logFile, logLine, ioMode)
	On Error Resume Next
	Set objFile = objFSO.OpenTextFile(logFile, ioMode, True)
	If Err.Number <> 0 Then
		Wscript.Echo "Error in log file. Desc: " & Err.Description
		Wscript.Quit
	End If
	timeNow = Hour(Now) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
	logLinewithTime = dateToday & " " & timeNow & "|" & logLine
	objFile.Write(logLinewithTime & vbCrLf)
	objFile.Close
End Sub

Function setenvvariables(project)
	Dim var, envVariables
	'Set wshProcessEnv = objShell.Environment( "Process" )
	Set envVariables = projenvvariables(project)
	For Each var in envVariables
		objEnvironment(var) = envVariables(var)
		'writeToLog logFile, envVariables(var), 8
	Next
	
End Function
Function projenvvariables(project)
	Dim projectCode, xmlFoldPath, xmlPath, depNodeList, root, projCode, var, depProjXmlPath, depEnvVariables, projEnvVariable, depProjects
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set projEnvVariable = CreateObject("Scripting.Dictionary")
	projectCode = project
	xmlFoldPath = "\\v0473a\PlantPLM\Integrations\PLM\07 NavisWorks\ProjectXmlList\PerProjectXML\"
	xmlPath = xmlFoldPath & projectCode & ".xml"
	objXMLDoc.async = False 
	If objFSO.FileExists(xmlPath) Then
		Set projEnvVariable = getenvvariables(xmlPath)
		writeToLog logFile, "Description: " & description(0).text, 8
		writeToLog logFile, "Setting environment variables from|" & xmlPath, 8
		objXMLDoc.load(xmlPath)
		Set root = objXMLDoc.documentElement 
		Set depNodeList = root.getElementsByTagName("DependencyProjectCodes")(0)
		depProjects = Split(depNodeList.text, ";")
		For Each projCode in depProjects
			If Len(Trim(projCode)) > 0 then
				depProjXmlPath = xmlFoldPath & projCode & ".xml"
				writeToLog logFile, "Setting environment variables from|" & depProjXmlPath, 8
				Set depEnvVariables = getenvvariables(depProjXmlPath)
				For Each var in depEnvVariables
					On Error Resume Next
					projEnvVariable.Add var, depEnvVariables(var)
				Next				
			End If
		Next
	Else
		writeToLog logFile, "Project xml file not found |" & projectCode & "|" & mdbName & "|" & xmlPath, 8
		WScript.Quit
	End If
	Set projenvvariables = projEnvVariable
End Function

Function getenvvariables(xmlPath)
	Dim projectEnvVariable, root, nodelist, name, path, objectXMLDoc, Elem
	Set objectXMLDoc = CreateObject("Microsoft.XMLDOM") 
	Set projectEnvVariable = CreateObject("Scripting.Dictionary")
	objectXMLDoc.load(xmlPath)
	Set root = objectXMLDoc.documentElement 
	Set nodeList = root.getElementsByTagName("ProjectEnviromentVariable") 
	Set description = root.getElementsByTagName("CommonDescription")
	For Each Elem In nodeList 
		SET name = Elem.getElementsByTagName("Name")(0)
		SET path = Elem.getElementsByTagName("Path")(0)
		On Error Resume Next
		projectEnvVariable.Add name.text, path.text
	Next
	Set getenvvariables = projectEnvVariable
End Function 