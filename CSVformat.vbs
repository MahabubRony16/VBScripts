Dim line
'---------------------------------------------------
Set arguments = Wscript.Arguments
inputtextfile = arguments(0)
outputcsvfile = arguments(1)
'---------------------------------------------------
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(inputtextfile,1)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(outputcsvfile,2,true)
msgbox("Formatter used")
do while not objFileToRead.AtEndOfStream
    strLine = objFileToRead.ReadLine()

	if InStr(1,strLine,"\~") = 1 Then
		line = line & "," & replace(strLine,"\~","")
	else
		if len(line) <> 0 Then
			objFileToWrite.WriteLine(line)
		end if
		line = strLine
	end if
loop
objFileToWrite.WriteLine(line)
objFileToRead.Close
objFileToWrite.Close
Set objFileToRead = Nothing
Set objFileToWrite = Nothing
