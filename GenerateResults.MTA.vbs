' GenerateResults.vbs
'
'
' Usage:	GenerateResults.vbs SourceDirectory ResultsDirectory debug|release CutNumber
' Example:  GenerateResults.vbs C:\Source\NGEN C:\UnitTestResults release 0120.28
'
'---------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------
' Global Variables
'----------------------------------------------------------------------------------------------

Set wFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set wShell = WScript.CreateObject("WScript.Shell")
dQ = String(1, 34)	' Double qoute

lngPassDeveloper = 0
lngPassDebugger = 0
lngPassWindowsRuntime = 0
lngPassWindowsDatabase = 0
lngPassCLRBuilder = 0
lngPassMCPBuilder = 0
'lngPassJ2EERuntime = 0
lngPassChangeAnalysis = 0
lngPassClientTools = 0

'Added for ATT Unit test
lngPassATT = 0

'Added for ClientFrameWork Unit test
lngPassClientFrameWork = 0

lngTotalDeveloper = 0
lngTotalDebugger = 0
lngTotalWindowsRuntime = 0
lngTotalWindowsDatabase = 0
lngTotalCLRBuilder = 0
lngTotalMCPBuilder = 0
'lngTotalJ2EERuntime = 0
lngTotalChangeAnalysis = 0
lngTotalClientTools = 0

'Added for ATT Unit test
lngTotalATT = 0

'Added for ClientFrameWork
lngTotalClientFrameWork = 0

public const eAreaDeveloper = 0
public const eAreaDebugger = 1
public const eAreaWindowsRuntime = 2
public const eAreaWindowsDatabase = 3
public const eAreaCLRBuilder = 4
public const eAreaMCPBuilder = 5
'public const eAreaJ2EEBuilder = 6
'public const eAreaJ2EERuntime = 7

'Added for ATT Unit test
public const eAreaATT = 8

' Added for Change Analysis

public const eAreaChangeAnalysis = 9

public const eAreaClientTools = 10

public const eAreaClientFrameWork = 11

'----------------------------------------------------------------------------------------------
' Main logic
'----------------------------------------------------------------------------------------------

If wscript.arguments.count = 5 Then
	SourceDirectory = WScript.Arguments(0) ' e.g. C:\Source\NGEN
	ResultsDirectory = WScript.Arguments(1) ' e.g. C:\UnitTestResults
	Configuration = WScript.Arguments(2) ' e.g. release
	CutNumber = WScript.Arguments(3) ' e.g. 0128.28
	LogUrlFormat = Wscript.Arguments(4) ' e.g. local | web

	SummaryFile = ResultsDirectory + "\UnitTest_summary.html"
	ErrorFile = ResultsDirectory + "\UnitTest_errors.html"
	LogFile = ResultsDirectory + "\UnitTest_logs.html"
	
	CompareTestRunFile = ResultsDirectory + "\UnitTest_compare_"&CutNumber&".txt"

	If Not (LCase(Configuration) = "release" Or LCase(Configuration) = "debug") Then
		Usage
		Return
	End If

	GenerateResults
Else
	Usage
End If

'----------------------------------------------------------------------------------------------
' Create Results to HTML file
' Examine all text files ending with testresult.txt in the input directory
' (passed in as arg to this function) and generate a HTML output
sub GenerateResults()
    WScript.Echo("Generating summary results with the following result files.")
    Set rootFolder = wFSO.GetFolder(ResultsDirectory)
    WriteOutputHeaders
    for each file in rootFolder.Files
        If Right(file.Name, 14) = "TestResult.txt" then
            testName = Left(file.Name, len(file.Name) - 14)
            WScript.Echo(file.Path)
            WriteOutputBody file.Path, testName
        End If
    next
    WriteOutputFooters
	'ExtractPerformanceMeasures
End sub

' Application to Log performance measures under the perfmeasure SQL database
sub ExtractPerformanceMeasures()
    perfCommand = dQ + SourceDirectory + "\Cut Scripts\Unit Test Scripts\PerformanceMeasure\CollectorApp" + dQ + _
				  + " " + dQ + ResultsDirectory + dQ + " /v " + "Cut" + CutNumber
    WScript.Echo "Running Extract Performance Measures."
    wShell.run perfCommand, 1, True
End sub

'----------------------------------------------------------------------------------------------
' Create and write header values to the main summary file, error file and log file.
sub WriteOutputHeaders()

	' Write Summary File Header
	wFSO.CreateTextFile SummaryFile, True 'overwrite if existing
	Set f = wFSO.OpenTextFile(SummaryFile, 2) '2 for writing
	f.writeLine("<html>")
	f.writeLine("<body link=black alink=black vlink=black>")
	f.writeLine("")
	f.writeLine("<font size=-2 face=verdana>")
	f.writeLine("<center>")
	f.writeLine("<h3>Unit Test Summary for NGEN: Cut " + CutNumber + ", Configuration " + Configuration + "</h3>")
	f.writeLine(FormatDateTime(now, vbGeneralDate))
	f.writeLine("<table width=85% border=0 bordercolor=white RULES=ALL FRAME=VOID cellpadding=4 cellspacing=1>")
	f.writeLine("<tr>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Project</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Area</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Result</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Pass</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Fail</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Total</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Pass Rate %</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Error Log</b></font></td>")
	f.writeLine("	<td bgcolor=#99CCFF><font size=-1><b>Complete Log</b></font></td>")
	f.close

	' Write Error File Header
	wFSO.CreateTextFile ErrorFile, True 'overwrite if existing
	Set f2 = wFSO.OpenTextFile(ErrorFile, 2) '2 for writing
	f2.writeLine("<html>")
	f2.writeLine("")
	f2.writeLine("<body>")
	f2.writeLine("<center>")
	f2.writeLine("<h3>Unit Test Error Log for NGEN: Cut " + CutNumber + ", Configuration " + Configuration + "</h3>")
	f2.writeLine(FormatDateTime(now, vbGeneralDate))
	f2.writeLine("</center>")
	f2.writeLine("<font face=verdana>")
	f2.close

	' Write Log File Header
	wFSO.CreateTextFile LogFile, True 'overwrite if existing
	Set f3 = wFSO.OpenTextFile(LogFile, 2) '2 for writing
	f3.writeLine("<html>")
	f3.writeLine("")
	f3.writeLine("<body>")
	f3.writeLine("<center>")
	f3.writeLine("<h3>Unit Test Log for NGEN: Cut " + CutNumber + ", Configuration " + Configuration + "</h3>")
	f3.writeLine(FormatDateTime(now, vbGeneralDate))
	f3.writeLine("</center>")
	f3.writeLine("<font face=verdana>")
	f3.close

	' Write Compare File Header
	wFSO.CreateTextFile CompareTestRunFile, True 'overwrite if existing
	Set f5 = wFSO.OpenTextFile(CompareTestRunFile, 2) '2 for writing
	f5.writeLine("Comparison file for Unit Test for NGEN")
	f5.writeLine("Cut " + CutNumber)
	f5.writeLine("Configuration " + Configuration)
	f5.writeLine("-----------------------------------------------------------------------------------------------")
	f5.close
End sub

'----------------------------------------------------------------------------------------------
' Write footer values to the main summary file, error file and log file
sub WriteOutputFooters()

	' Write Summary File Footer
	Set f = wFSO.OpenTextFile(SummaryFile, 8) '8 for appending

	'Set f1 = wFSO.OpenTextFile(CompareTestRunFile, 8) '9 for appending

	' Write Totals
    Call WriteAreaTotal(f, "DEVELOPER", lngPassDeveloper, lngTotalDeveloper)
    Call WriteAreaTotal(f, "DEBUGGER", lngPassDebugger, lngTotalDebugger)
    Call WriteAreaTotal(f, "WINDOWS RUNTIME", lngPassWindowsRuntime, lngTotalWindowsRuntime)
    'Call WriteAreaTotal(f, "J2EE RUNTIME", lngPassJ2EERuntime, lngTotalJ2EERuntime)
    Call WriteAreaTotal(f, "WINDOWS DATABASE", lngPassWindowsDatabase, lngTotalWindowsDatabase)
    Call WriteAreaTotal(f, "CLR BUILDER", lngPassCLRBuilder, lngTotalCLRBuilder)
    Call WriteAreaTotal(f, "MCP BUILDER", lngPassMCPBuilder, lngTotalMCPBuilder)
	' Added for ChangeAnalysis
	Call WriteAreaTotal(f, "CHANGE ANALYSIS", lngPassChangeAnalysis, lngTotalChangeAnalysis)

	'Added for ATT Unit test
	Call WriteAreaTotal(f, "ATT", lngPassATT, lngTotalATT)

	Call WriteAreaTotal(f, "CLIENTTOOLS", lngPassClientTools, lngTotalClientTools)
	'Added for ClientFrameWork Unit test
	Call WriteAreaTotal(f, "CLIENTFRAMEWORK", lngPassClientFrameWork, lngTotalClientFrameWork)

	' Total Line
	lngPassTotal = lngPassDeveloper + lngPassDebugger + lngPassWindowsRuntime + lngPassWindowsDatabase + lngPassCLRBuilder + lngPassMCPBuilder + lngPassJ2EERuntime + lngPassATT + lngPassChangeAnalysis + lngPassClientTools + lngPassClientFrameWork
	lngCompleteTotal = lngTotalDeveloper + lngTotalDebugger + lngTotalWindowsRuntime + lngTotalWindowsDatabase + lngTotalCLRBuilder + lngTotalMCPBuilder + lngTotalJ2EERuntime + lngTotalATT + lngTotalChangeAnalysis + lngTotalClientTools + lngTotalClientFrameWork

	lngTotalPassRate = 0.00
	If lngCompleteTotal > 0 Then
		lngTotalPassRate = FormatNumber(lngPassTotal / lngCompleteTotal * 100, 2)
	End If
	f.writeLine("<tr>")
	f.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>TOTAL</b></font></td>")
	f.writeLine("	<td></td><td></td>")
	f.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=green><b>"&lngPassTotal&"</b></font></td>")
	f.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=red><b>"&lngCompleteTotal - lngPassTotal&"</b></font></td>")
	f.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&lngCompleteTotal&"</b></font></td>")
	f.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&lngTotalPassRate&"</b></font></td>")
	f.writeLine("	<td></td><td></td>")
	f.writeLine("</tr>")


	f.writeLine("</table>")
	f.writeLine("</center>")
	f.writeLine("")
	f.writeLine("<p>")
	f.writeLine("</body>")
	f.writeLine("</html>")
	f.close

	' Write Error File Footer
	Set f2 = wFSO.OpenTextFile(ErrorFile, 8) '8 for appending
	f2.writeLine("</body>")
	f2.writeLine("</html>")
	f2.close

	' Write Log File Footer
	Set f3 = wFSO.OpenTextFile(LogFile, 8) '8 for appending
	f3.writeLine("</body>")
	f3.writeLine("</html>")
	f3.close
End sub

'----------------------------------------------------------------------------------------------
' Write unit test result content to the main output file, error file and log file.
' Remove temporary files created.
sub WriteOutputBody(file, testName)
	'Wscript.echo file & " is unicode: " & isUnicode(file)

	' Check if its Unicode
	if isUnicode(file) = True Then
		Set f = wFSO.OpenTextFile(file,1,True,-1)
	Else
		Set f = wFSO.OpenTextFile(file)
	End If

	' open the log files for appending because the headers are already written
	Set f2 = wFSO.OpenTextFile(SummaryFile, 8)
	Set f3 = wFSO.OpenTextFile(ErrorFile, 8)
	Set f4 = wFSO.OpenTextFile(LogFile, 8)
	
	Set f5 = wFSO.OpenTextFile(CompareTestRunFile, 8)

	Set FailedRe = New RegExp
	FailedRe.Pattern = "(.*=+ Failed =+)"						'Matches - ==========Failed=========

	Set StartFail = New RegExp
	StartFail.Pattern = "(^\s*1\))"								'Matches - 1) This test has failed
	Set StartFail2 = New RegExp
	StartFail2.Pattern = "(\!\!\!FAILURES\!\!\!)"

	Set EndFail = New RegExp
	EndFail.Pattern = "(.*)Total(\s+\d+)|(\w+::\w+\s+-\s+\w+)"	'Matches - Total 123 OR Class::Method - fail

	Set ResultRe = New RegExp									'Matches - Total Run: 123, Failures: 23, Not run: 12,
	ResultRe.Pattern = "([\w\s]*[rR]un:(\s+\d+)[\s,]*Failures:(\s+\d+)[\s,]*Not run:(\s+\d+)[\s,]*)"

	Set NUnitResultRe = New RegExp							    'Matches - Tests run: 123, Errors: 123, Failures: 23,
	NUnitResultRe.Pattern = "([\w\s]*[rR]un:(\s+\d+)[\s,]*Errors:(\s+\d+)[\s,]*Failures:(\s+\d+)[\s,]*)"
	
	Set MSUnitResultRe = New RegExp									'Matches - 212/212 test(s) Passed
	MSUnitResultRe.Pattern = "([\s]*(\d+)\/(\d+)[\s]*test\(s\) Passed[\s]*)"
	Set SuccessRe = New RegExp
	
	'CT changes
	Set JUnitResultRe = New RegExp                              'Matches - Tests run: 45, Failures: 0, Errors: 0,
	JUnitResultRe.Pattern = "([\w\s]*[rR]un:(\s+\d+)[\s,]*Failures:(\s+\d+)[\s,]*Errors:(\s+\d+)[\s,]*)"

	Set UnitTestCSResultRe = New RegExp							    'Matches - Tests run: 123, Errors: 123, Failures: 23,
	UnitTestCSResultRe.Pattern = "([\w\s]*[Total tests]:(\s+\d+)[\s\.]*Passed:(\s+\d+)[\s\.]*Failed:(\s+\d+)[\s,]*)"

	'SuccessRe.Pattern = "(.*)Succeeded(.*)|(.*)Success(.*)"		'Matches - This test succeeded
	'This fix below will solve the problem of getting wrong result
	'for tests like DebuggerRuntimeIntegrationTest, MCPBuuilderTest etc.
	SuccessRe.Pattern = "(.*)Succeeded\s+(.*)"		'Matches - This test succeeded 

	Set OKRe = New RegExp
	OKRe.Pattern = "(OK \((\d+) tests\))"
	Set AsyncRe = New RegExp
	AsyncRe.Pattern = "(Run\:  \d+   Failures\: (\d+)   Random\: \d+   Expected\: \d+   Errors\: \d+)"

Set testRe = New RegExp
	testRe.Pattern = "(Run\:.*)"

	' For log files
	failedLines = ""
	resultLine = ""
	logLines = ""

	' counter for fails and stuff
	lngTotal = 0
	lngPass = 0
	lngFail = 0
	lngError = 0
	blnFail = False
	blnParseFail = False
	bNUnitTest = False
	testNameNew = Split(TestName, "-", -1, 1)
	eTestArea = GetTestArea(testNameNew(UBound(testNameNew)))

	'This bit works out how many tests have failed or passed, all unit tests except UnitTestCS has same format so
	filename = wFSO.GetFileName(file)
	Do While f.AtEndOfStream = False
		' read the line
		line = f.ReadLine

		' add the line to logfile
		logLines = logLines + line + vbNewLine

		' parse the results from 'Test Results' line for this file only
		' Added Painter and ATT Unit Test Result file
		If wFSO.GetFileName(file) = "DBMigrateTestResult.txt"    Or _
		   wFSO.GetFileName(file) = "ClientFrameWorkTestResult.txt"  Or _
		   wFSO.GetFileName(file) = "DBReorg1TestResult.txt"     Or _
		   wFSO.GetFileName(file) = "DBReorg2TestResult.txt"     Or _
		   wFSO.GetFileName(file) = "PainterUnitTestTestResult.txt" Or _
		   wFSO.GetFileName(file) = "PresentationClientTestResult" Or _
		   wFSO.GetFileName(file) = "CTUnitTest_dotnetTestResult.txt" Or _
		   wFSO.GetFileName(file) = "ManagedDataTypesTestResult.txt" Or _
		    wFSO.GetFileName(file) = "RuntimeAPITestResult.txt" Then

			bNUnitTest = True

			If ResultRe.test(line) = True Then
				Set Matches = ResultRe.Execute(line)
				sTotal = Matches(0).SubMatches(1)		'[rR]un:(.*)
				sFail = Matches(0).SubMatches(2)		'Failures:(.*)
				sError = Matches(0).SubMatches(3)		'Errors:(.*)

			    ' WScript.Echo "Total " & sTotal & " Error " & sError & " Fail " & sFail
				If isNumeric(sTotal) And isNumeric(sFail) And isNumeric(sError) Then
					lngTotal = CInt(sTotal)
					lngFail = CInt(sError) + CInt(sFail)
					lngPass = lngTotal - lngFail
				Else
					blnParseFail = True
				End If
			End If
		End If


			' parse the nunit results
		If Right(filename, 21) = "PainterTestResult.txt" Or _
		   wFSO.GetFileName(file) = "PresentationClientTestResult.txt" Or _
		   wFSO.GetFileName(file) = "CTUnitTest_dotnetTestResult.txt" Or _
		    wFSO.GetFileName(file) = "RuntimeAPITestResult.txt" Then
			If NUnitResultRe.test(line) = True Then
				Set Matches = NUnitResultRe.Execute(line)
				sTotal = Matches(0).SubMatches(1)		'[rR]un:(.*)
				sFail = Matches(0).SubMatches(2)		'Failures:(.*)
				sError = Matches(0).SubMatches(3)		'Errors:(.*)

			    ' WScript.Echo "Total " & sTotal & " Error " & sError & " Fail " & sFail
				If isNumeric(sTotal) And isNumeric(sFail) And isNumeric(sError) Then
					lngTotal = CInt(sTotal)
					lngFail = CInt(sError) + CInt(sFail)
					lngPass = lngTotal - lngFail
					If lngFail > 0 Then
					failedLines = failedLines + line + vbNewLine
					End If
				Else
					blnParseFail = True
				End If
			End If
		End If

		If Right(filename, 24) = "UnitTestCSTestResult.txt" Or _
		   Right(filename, 28) = "VersionControlTestResult.txt" Or _
		   Right(filename, 17) = "ATTTestResult.txt" Then
			If UnitTestCSResultRe.test(line) = True Then
				Set Matches = UnitTestCSResultRe.Execute(line)
				sTotal = Matches(0).SubMatches(1)		'[rR]un:(.*)
				sPass = Matches(0).SubMatches(2)		'Failures:(.*)
				sFail = Matches(0).SubMatches(3)		'Errors:(.*)

			    lngTotal = CInt(sTotal)
				lngFail = CInt(sFail)
				lngPass = CInt(sPass)
				If lngFail > 0 Then
					blnFail = true
				End If
			End If
		End If
		
		' parse the MS Unit results
		If wFSO.GetFileName(file) = "SMUnitTestTestResult.txt" Then
			If MSUnitResultRe.test(line) = True Then
				Set Matches = MSUnitResultRe.Execute(line)
				sPassed = Matches(0).SubMatches(1)		'212/
				sTotal = Matches(0).SubMatches(2)		'/212

			    ' WScript.Echo "Total " & sTotal & " Error " & sError & " Fail " & sFail
				If isNumeric(sTotal) And isNumeric(sPassed) Then
					lngTotal = CInt(sTotal)
					lngPass = CInt(sPassed)
					lngFail = lngTotal - lngPass
					If lngFail > 0 Then 
					failedLines = failedLines + line + vbNewLine
					End If
				
				Else
					blnParseFail = True
				End If
			End If
		End If
		
		'CT changes
		' parse the junit results
		
		If wFSO.GetFileName(file) = "PresentationClientTestResult.txt" Then
			
			If JUnitResultRe.test(line) = True Then
				Set Matches = JUnitResultRe.Execute(line)
				sTotal = Matches(0).SubMatches(1)		'[rR]un:(.*)
				sFail = Matches(0).SubMatches(2)		'Failures:(.*)
				sError = Matches(0).SubMatches(3)		'Errors:(.*)

			    ' WScript.Echo "Total " & sTotal & " Error " & sError & " Fail " & sFail
				If isNumeric(sTotal) And isNumeric(sFail) And isNumeric(sError) Then
					lngTotal = CInt(sTotal)
					lngFail = CInt(sError) + CInt(sFail)
					lngPass = lngTotal - lngFail
					If lngFail > 0 Then
					failedLines = failedLines + line + vbNewLine
					End If
				Else
					blnParseFail = True
				End If
			End If
		End If
		
		
		' determine whether to print out error log
		If blnFail = True And EndFail.test(line) = True Then
			blnFail = False
		Elseif StartFail.test(line) = True Then
			blnFail = True
			failedLines = failedLines + vbNewLine
		Elseif StartFail2.test(line) = True Then
			blnFail = True
			failedLines = failedLines + vbNewLine
		End If

		' determine whether to print out error log
		If blnFail = True And EndFail.test(line) = True Then
			blnFail = False
		Elseif StartFail.test(line) = True Then
			blnFail = True
			failedLines = failedLines + vbNewLine
		Elseif StartFail2.test(line) = True Then
			blnFail = True
			failedLines = failedLines + vbNewLine
		End If

		' Write to logs
		If FailedRe.test(line) = True Then	' put it into error log file and increase counter
			failedLines = failedLines + line + vbNewLine
			lngFail = lngFail + 1
			lngTotal = lngTotal + 1
			
		Elseif blnFail = True Then			' put in log but dont increase counter.
			failedLines = failedLines + line + vbNewLine
		Elseif SuccessRe.Test(line) = True Then	' increase pass counter
			lngPass = lngPass + 1
			lngTotal = lngTotal + 1
		Elseif AsyncRe.Test(line) = True Then
			Set Matches = AsyncRe.Execute(line)
			lngFail = lngFail + CInt(Matches(0).SubMatches(1))
		Elseif OKRe.Test(line) = True Then
			Set Matches = OKRe.Execute(line)
			lngTotal = Matches(0).SubMatches(1)
			lngPass = lngTotal - lngFail
		End If

	Loop

	' Write the results

	'Determine Pass Rate


	' Add to project area totals
	Call UpdateTotals(eTestArea, lngPass, lngTotal)

	' Draw the Summary Table, Log Files, Error Log Files
	f2.writeLine("<tr>")

	' Project Column
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2>" + TestName + "</font></td>")

	' Area Column
	sAreaName = GetProjectAreaName(eTestArea)
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2>" + sAreaName + "</font></td>")

	' Result Column and Comments
	If (blnParseFail) Then		                  '  Failed to parse results
		f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=orange><b>ERROR - Failed to read results</b></font></td>")
		
		f5.writeLine("TEST="&TestName)
		f5.writeLine("TOTAL=" & " " & lngTotal)
		f5.WriteLine("")
		

	ElseIf (logLines = "") Or (lngTotal = 0) Then ' If log is empty or zero tests then run is incomplete
		f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=orange><b>ERROR - Test Incomplete</b></font></td>")

		
		f5.writeLine("TEST="&TestName)
		f5.writeLine("TOTAL=" & " " & lngTotal)
		f5.WriteLine("")

	ElseIf failedLines = "" Then	              ' If the fail log is empty then successful
		f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=green><b>PASSED</font></td>")

		
		f5.writeLine("TEST="&TestName)
		f5.writeLine("TOTAL=" & " " & lngTotal)
		f5.WriteLine("")
						
	Else
		f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=red><b>FAILURES</b></font></td>")

		
		f5.writeLine("TEST="&TestName)
		f5.writeLine("TOTAL=" & " " & lngTotal)
		f5.WriteLine("")

	End	If
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=green><b>"&lngPass&"</b></font></td>")
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=red><b>"&lngFail&"</b></font></td>")
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&lngTotal&"</b></font></td>")
	f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&PassRateAsString(lngPass, lngTotal)&"</b></font></td>")
	
	Set TestNameFormat = New RegExp
	TestNameFormat.Pattern = "((\d+\-\d+\-\d+\-.*\-.*)\-UnitTestResults\-.*)"
	TestNameVersion = ""
	If TestNameFormat.test(TestName) = True Then
		Set Matches = TestNameFormat.Execute(TestName)
		TestNameVersion = Matches(0).SubMatches(1)
	End If

	' Error Log Column
	If failedLines = "" Then
		f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2>-</font></td>")
	Else
		f3.writeLine("<h3><a Name="+dQ+TestName+dQ+">"+TestName+"</h3>")
		f3.writeLine("<pre>")
		f3.write(failedLines)
		f3.writeLine("</pre>")
		f3.writeLine("<p>")
		f3.writeLine("<hr size=1>")
		f3.writeLine("<p>")
		If (LogUrlFormat = "local") Then
			f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2><a href=""UnitTest_errors.html#"&TestName&""" target=_blank>Error File</a></font></td>")
		Else
			f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2><a href=""..\\UnitTest_errors\\"&TestNameVersion&"-UnitTest_errors.html#"&TestName&""" target=_blank>Error File</a></font></td>")
		End If
	End If

	' Complete Log Column
     If logLines = "" Then
          f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2>-</font></td>")
     Else
          f4.writeLine("<h3><a Name="+dQ+TestName+dQ+">"+TestName+"</h3>")
          f4.writeLine("<pre>")
          f4.write(logLines)
          f4.writeLine("</pre>")
          f4.writeLine("<p>")
          f4.writeLine("<hr size=1>")
          f4.writeLine("<p>")
		  If (LogUrlFormat = "local") Then
			f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2><a href=""UnitTest_logs.html#"&TestName&""" target=_blank>Log File</a></font></td>")
		  Else
			f2.writeLine("	<td bgcolor=#E8E8E8><font size=-2><a href=""..\\UnitTest_logs\\"&TestNameVersion&"-UnitTest_logs.html#"&TestName&""" target=_blank>Log File</a></font></td>")
		  End If
     End If

	f4.close
	f3.close
	f2.close
		
	f5.close

	f.close
End Sub

'--------------------------------------------------------------------------------------------
' This function checks whether the file is a Unicode File
'--------------------------------------------------------------------------------------------
Function isUnicode(file)
	Set UnicodeRe = New RegExp
	UnicodeRe.Pattern = "\xFF\xFE"		' Matches the Unicode BOM

	'Make sure the size of file is bigger than 2 bytes
	Set fSizeTest = wFSO.GetFile(file)
	If fSizeTest.size > 2 Then
		' Open the file and read the 1st 2 characters
		Set fTest = wFSO.OpenTextFile(file)
		strTest = fTest.Read(2)

		' Test and return results
		If UnicodeRe.Test(strTest) then
			isUnicode = True
		Else
			isUnicode = False
		End If

		fTest.close
	Else
		' If the size is smaller than 2 bytes then it doesnt matter
		isUnicode = False
	End If
End Function


Function WriteAreaTotal(fStream, sArea, lngPass, lngTotal)
	fStream.writeLine("<tr>")
	fStream.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&UCase(sArea)&" TOTAL</b></font></td>")
	fStream.writeLine("	<td></td><td></td>")
	fStream.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=green><b>"&lngPass&"</b></font></td>")
	fStream.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=red><b>"&lngTotal - lngPass&"</b></font></td>")
	fStream.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&lngTotal&"</b></font></td>")
	fStream.writeLine("	<td bgcolor=#E8E8E8><font size=-2 color=black><b>"&PassRateAsString(lngPass, lngTotal)&"</b></font></td>")
	fStream.writeLine("	<td></td><td></td>")
	fStream.writeLine("</tr>")

End Function

Function PassRateAsString(lngPass, lngTotal)
	If lngTotal > 0 Then
		PassRateAsString = CStr(FormatNumber(lngPass / lngTotal * 100, 2))
    Else
        PassRateAsString = "0.00"
	End If
End Function

Function UpdateTotals(eTestArea, lngPass, lngTotal)
	If eTestArea = eAreaDeveloper Then
		lngPassDeveloper = lngPassDeveloper + lngPass
		lngTotalDeveloper = lngTotalDeveloper + lngTotal
	ElseIf eTestArea = eAreaDebugger Then
		lngPassDebugger = lngPassDebugger + lngPass
		lngTotalDebugger = lngTotalDebugger + lngTotal
	ElseIf eTestArea = eAreaWindowsRuntime Then
		lngPassWindowsRuntime = lngPassWindowsRuntime + lngPass
		lngTotalWindowsRuntime = lngTotalWindowsRuntime + lngTotal
	ElseIf eTestArea = eAreaWindowsDatabase Then
		lngPassWindowsDatabase = lngPassWindowsDatabase + lngPass
		lngTotalWindowsDatabase = lngTotalWindowsDatabase + lngTotal
	ElseIf eTestArea = eAreaCLRBuilder Then
		lngPassCLRBuilder = lngPassCLRBuilder + lngPass
		lngTotalCLRBuilder = lngTotalCLRBuilder + lngTotal
	ElseIf eTestArea = eAreaMCPBuilder Then
		lngPassMCPBuilder = lngPassMCPBuilder + lngPass
		lngTotalMCPBuilder = lngTotalMCPBuilder + lngTotal
	'commented J2ee runtime
	'ElseIf eTestArea = eAreaJ2EERuntime Then
		'lngPassJ2EERuntime = lngPassJ2EERuntime + lngPass
		'lngTotalJ2EERuntime = lngTotalJ2EERuntime + lngTotal
	
	'Added ChangeAnalysis.

	ElseIf eTestArea = eAreaChangeAnalysis Then
		lngPassChangeAnalysis = lngPassChangeAnalysis + lngPass
		lngTotalChangeAnalysis = lngTotalChangeAnalysis + lngTotal
		
	ElseIf eTestArea = eAreaClientTools Then
		lngPassClientTools = lngPassClientTools + lngPass
		lngTotalClientTools = lngTotalClientTools + lngTotal

'Added for ClientFrameWork Unit test

	ElseIf eTestArea = eAreaClientFrameWork Then
		lngPassClientFrameWork = lngPassClientFrameWork + lngPass
		lngTotalClientFrameWork = lngTotalClientFrameWork + lngTotal

	'Added for ATT
	ElseIf eTestArea = eAreaATT Then
		lngPassATT = lngPassATT + lngPass
		lngTotalATT = lngTotalATT + lngTotal
	End If
End Function

'--------------------------------------------------------------------------------------------
' This function returns the project area of the test
'--------------------------------------------------------------------------------------------
Function GetTestArea(test)
'added Painter/XMIIm/ExPort/Performance tests
'added PModel

	If test = "Builder" Or _
	   test = "DepconReport" Or _
	   test = "Language" Or _
	   test = "LCIFImport" Or _
	   test = "Licensing" Or _
	   test = "ModelBenchmarkSuite" Or _
	   test = "Model" Or _
   	   test = "PModel" Or _
	   test = "Other" Or _
	   test = "Search" Or _
	   test = "UnitTestCS" Or _
	   test = "Painter" Or _
	   test = "Converter" Or _
   	   test = "OS2200Builder" Or _
   	   test = "VersionControl" Or _
	   test = "XMI" Or _ 
	   test = "XMIExporter" Or _
	   test = "XMIImporter" Or _
	   test = "XMIPerformance"  Then
            GetTestArea = eAreaDeveloper
	ElseIf test = "Debugger" Or _
	       test = "DebuggerRuntimeIntegration" Then
            GetTestArea = eAreaDebugger
	ElseIf test = "CLRBuilder" Or _
           test = "BuilderSystem" Then
            GetTestArea = eAreaCLRBuilder
	ElseIf test = "MCPBuilder" Then
            GetTestArea = eAreaMCPBuilder
    ElseIf test = "DBReorg1" Or _
           test = "DBReorg2" Or _
           test = "DBMigrate" Or _
           test = "DBConfiguration" Then
            GetTestArea = eAreaWindowsDatabase
	'ElseIf test = "J2EERuntime" Then
	    'GetTestArea = eAreaJ2EERuntime

	ElseIf test = "ChangeAnalysisMCP" Or _
		   test = "ChangeAnalysisWindows" Or _
		   test = "ChangeAnalysisClientToolMCP" Or _
		   test = "ChangeAnalysisClientToolWindows" Or _
		   test = "ChangeAnalysisDebugger" Then
	    GetTestArea = eAreaChangeAnalysis
	
	'Added for ATT
	ElseIf test = "ATT" Then
		GetTestArea = eAreaATT
	' Added for CT
	ElseIf 	test = "CTUnitTest_dotnet" Or _
			test = "PresentationClient" Then
		GetTestArea = eAreaClientTools
'Added for ClientFrameWork Unit test
	ElseIf test = "ClientFrameWork" Then
		GetTestArea = eAreaClientFrameWork

	Else
		    GetTestArea = eAreaWindowsRuntime
	End If
End Function


Function GetProjectAreaName(eTestArea)
    If eTestArea = eAreaDeveloper Then
        GetProjectAreaName = "Developer"
    ElseIf eTestArea = eAreaDebugger Then
        GetProjectAreaName = "Debugger"
    ElseIf eTestArea = eAreaWindowsRuntime Then
        GetProjectAreaName = "Windows Runtime"
    'ElseIf eTestArea = eAreaJ2EERuntime Then
        'GetProjectAreaName = "J2EE Runtime"
    ElseIf eTestArea = eAreaWindowsDatabase Then
        GetProjectAreaName = "Windows Database"
    ElseIf eTestArea = eAreaCLRBuilder Then
        GetProjectAreaName = "CLR Builder"
    ElseIf eTestArea = eAreaMCPBuilder Then
        GetProjectAreaName = "MCP Builder"
    'ElseIf eTestArea = eAreaJ2EEBuilder Then
        'GetProjectAreaName = "J2EE Builder"
	
'Added for ChangeAnalysis.
ElseIf eTestArea = eAreaChangeAnalysis Then
        GetProjectAreaName = "ChangeAnalysis"

	'Added for ATT
	ElseIf eTestArea = eAreaATT Then
        GetProjectAreaName = "ATT"
    
	ElseIf eTestArea = eAreaClientTools Then
        GetProjectAreaName = "ClientTools"
'Added for ClientFrameWork Unit test
	ElseIf eTestArea = eAreaClientFrameWork Then
        GetProjectAreaName = "ClientFrameWork"
	Else
        GetProjectAreaName = "Unknown"
    End If
End Function

'--------------------------------------------------------------------------------------------
' Show the usage for this script and quit
'--------------------------------------------------------------------------------------------
Sub Usage
	WScript.Echo "Usage: " + WScript.ScriptName + " <SourceDirectory> <ResultsDirectory> <Configuration> <cut number>" + vbNewLine + _
				 "Example: " + WScript.ScriptName + " C:\Source\NGEN C:\UnitTestResults release 0120.28"
	WScript.Quit
End Sub

