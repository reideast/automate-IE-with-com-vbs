Option Explicit

If countIEWindows() <> 0 Then
	MsgBox "Internet Explorer is already open! Please close any IE windows before running this script. Quitting now.", 48, "Error"
	WScript.Quit
End If

'*****************************************************************************************
'************************              Setup             *********************************
'*****************************************************************************************

'Definintion of program parameters
Const PRINT_MODE = True

Const strURL = "http://www.fake_web_site.com/login.php"
Const strUsername = "user1"
Const strPassword = "password1"

Const strFilenameCSVtoArray = "CSVtoArray.vbs"
Const strFilenameModelCSV = "reference_model_num.csv"

'Get command line arguments: http://technet.microsoft.com/en-us/library/ee156618.aspx
Dim useCSV
Dim strFilenameCSV
If Wscript.Arguments.Named.Exists("ImportCSV") Then
	strFilenameCSV = Wscript.Arguments.Named.Item("ImportCSV")
	useCSV = True
Else
	useCSV = False
End If
'DEBUG: If useCSV Then MsgBox "DEBUG: CSV input file name: " & strFilenameCSV

'Technique to present a Yes/No Error msgbox. See: http://msdn.microsoft.com/en-us/library/139z2azd(v=vs.90).aspx
'These values are built-in to VBS:
'Const vbYesNo = 4
'Const vbYesNoCancel = 3
'Const vbExclamation = 48
'Const vbQuestion = 32
'Const vbNo = 7
'Const vbYes = 6
'Const vbCancel = 2
Dim msgValue
'EXAMPLE: msgValue = MsgBox("MSG BOX TEXT", vbYesNo + vbExclamation, "Error: MSG BOX TITLE")
'If (msgValue = vbYes) Then
'	'action upon user choosing Yes
'ElseIf (msgValue = vbNo) Then
'	'action upon user choosing No
'End If


'Technique to use Function confirmQuit to present a Yes/No choice before quitting
'If confirmQuit Then WScript.quit

'system objects to file IO
Dim objFSO, WshShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
Const vbFileForAppending = 8

'files to read data
Dim objFileCSV

'For loop counter
Dim intCount
Dim numPCIDtoProcess

'global variable, used by sub WaitInternetExplorer
Dim isIELoadCompleted

'Technique to Open Other .VBS "library" files
'http://technet.microsoft.com/en-us/library/ee176984.aspx
Set objFileCSV = objFSO.OpenTextFile(strFilenameCSVtoArray, 1)
Execute objFileCSV.ReadAll() 'reads file into memory, making function CSVtoArray() available for use
objFileCSV.Close

'Function CSVtoArray
'argument: strFileName - Path to the .csv file to read
'return value: an arry of dictionaries with the header row of the CSV file as the keys
'You can retrieve these values by: arrModels(1 to numModels - 1).Item("commonName") etc.
Dim arrModels, numModels
arrModels = CSVtoArray(strFilenameModelCSV)
numModels = UBound(arrModels)

Dim updateChoice
Dim arrPCIDData

'*****************************************************************************************
'************   Get PCID data from CSV file provided in command line argument   **********
'*****************************************************************************************
If useCSV Then
	'Dim arrPCID_CSV
	'arrPCID_CSV = CSVtoArray(strFilenameCSV)
	'numPCIDtoProcess = UBound(arrPCID_CSV)
	'ReDim arrPCIDData(numPCIDtoProcess)
	
	arrPCIDData = CSVtoArray(strFilenameCSV)
	''*****************************************************************************************shift this array to 1-indexed, or make manual input references 0-indexed!
	numPCIDtoProcess = UBound(arrPCIDData)
	
	updateChoice = 1 'always update if using CSV
Else 'not useCSV
	'*****************************************************************************************
	'************************          Get User Input        *********************************
	'*****************************************************************************************
	Dim strPrompt
	Dim strModelListPrompt, intCountModels
	strModelListPrompt = ""
	For intCountModels = 0 to (numModels)
		strModelListPrompt = strModelListPrompt & (intCountModels + 1) & ") " & arrModels(intCountModels).Item("commonName") & VBLF
	Next
	
	strPrompt = "Would you like to update the database?" & VBLF & VBLF & "1) Yes" & VBLF & "2) No" & VBLF & VBLF & _
		"Option 1 will prompt you for a Serial Number and it will also print after it is done updating" & VBLF & VBLF & _
		"Option 2 will open the PCID details page, and then update the text files that the script will to install apps on the PC"
	updateChoice = getUserInputInteger(strPrompt, 1, 2)

	strPrompt = "How many PCID's would you like to process?"
	numPCIDtoProcess = getUserInputInteger(strPrompt, 1, 50)

	'***************************************************************************************************************************
	'**************************   Get PCID, Serial, Model for each unit the user wants to process   ****************************
	'***************************************************************************************************************************
	ReDim arrPCIDData(numPCIDtoProcess - 1)
	Dim intCurrentPCID, strCurrentNewSerialNumber, intCurrentNewModel
	For intCount = 0 to (numPCIDtoProcess - 1)
		Set arrPCIDData(intCount) = CreateObject("Scripting.Dictionary")

		strPrompt = "PCID #" & (intCount + 1) & VBLF & VBLF & "What PCID do you want to search for?"
		intCurrentPCID = getUserInputString(strPrompt, 5)
		arrPCIDData(intCount).Add "PCID", intCurrentPCID
		
		If updateChoice=1 Then
			strPrompt = "PCID #" & (intCount + 1) & VBLF & VBLF & "PCID: " & intCurrentPCID & VBLF & VBLF & "New Serial Number?"
			strCurrentNewSerialNumber = getUserInputString(strPrompt, 10)
			arrPCIDData(intCount).Add "NewSerialNumber", strCurrentNewSerialNumber

			strPrompt = "PCID #" & (intCount + 1) & VBLF & VBLF & "PCID: " & intCurrentPCID & VBLF & "New Serial Number:" & strCurrentNewSerialNumber & VBLF & VBLF & "Which model are you currently working on?" & VBLF & VBLF
			strPrompt = strPrompt & strModelListPrompt
			intCurrentNewModel = getUserInputInteger(strPrompt, 1, numModels + 1)
			intCurrentNewModel = intCurrentNewModel - 1 'match up to the zero-indexed array
			arrPCIDData(intCount).Add "NewModel", intCurrentNewModel

		End IF 'End If: If updateChoice = 1
	Next 'End For: Loop through # of PCID's to process
End If 'useCSV or get use input

'DEBUG: Print arrPCIDData()
Dim strOutput
strOutput = "DEBUG: arrPCIDData() = " & vblf
For intCount = 0 to (numPCIDtoProcess - 1)
	strOutput = strOutput & intCount & ".) " _
				& arrPCIDData(intCount).Item("PCID") & "-"  _
				& arrPCIDData(intCount).Item("NewSerialNumber") & "-" _
				& arrPCIDData(intCount).Item("NewModel") & vblf
Next
MsgBox strOutput
'wscript.quit

'************************************************************************************
'****************     Modify Database Values in Internet Explorer      ******************
'************************************************************************************

'********************     Open IE and Navigate to the Web Page    ********************
Dim IE
Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
IE.Visible = True
IE.Navigate strURL
WaitInternetExplorer

With IE.Document

'************************     Login      *********************************
.getElementByID("txtUserName").value = strUsername
.getElementByID("txtPassword").value = strPassword
.getElementsByName("btnSubmit")(0).click
WaitInternetExplorer

'************************************************************************************
'****************     Loop Through the PCIDs User Wants to Process    ***************
'************************************************************************************

'variables for update
Dim isUpdateSuccessful
Dim confirmChanges
Dim strLogFileName, strLogAppsFileName
Dim strFileNameCurrentDate, strFileNameCurrentTime, strFileNameSerial
Dim strCurrentStatus

For intCount = 0 to (numPCIDtoProcess - 1)
	isUpdateSuccessful = True
	
	Set IE = findIEWindow(IE)
	'***********************     Search for PCID      *********************************
	.getElementByID("lvPCList_txtEnduserinformationID").value = arrPCIDData(intCount).Item("PCID")
	Wscript.sleep 300 '******************************************************* forced wait b/c IE seems to be hanging on this spot often
	.getElementByID("lvPCList_btnSearch").click
	'.getElementByID("lvPCList_txtEnduserinformationID").select
	'WshShell.SendKeys "{Enter}"
	WaitInternetExplorer

	If .getElementByID("lblTotalPCFound").innerhtml = "Total PC's Found : 1" Then
		Set IE = findIEWindow(IE)
		'************************************* Open the PCID we just found *************************************
		.getElementsByName("lvPCList_ctrl0_lnkSRID")(0).click
		WaitInternetExplorer
		Set IE = findIEWindow(IE)
		'*************** Create Log Files *******************
		strFileNameCurrentDate = Replace(.getElementById("lblCurrentDate").innerhtml, "/", "-")
		strFileNameCurrentTime = Replace(.getElementById("lblCurrentTime").innerhtml, ":", "")
		Set IE = findIEWindow(IE)
		If updateChoice=1 Then
			strFileNameSerial = arrPCIDData(intCount).Item("NewSerialNumber")
		Else
			strFileNameSerial = .getElementsByName("txtSerialNo")(0).Value
		End If
		strLogFileName = ".\Output\" & arrPCIDData(intCount).Item("PCID") & "-" & strFileNameSerial & "-" & strFileNameCurrentDate & "(" & strFileNameCurrentTime & ")" & ".txt"
		strLogAppsFileName = ".\Output\" & arrPCIDData(intCount).Item("PCID") & "-" & strFileNameSerial & "_APPS.txt"
		'*************** Write Log File with Preexisting Values (for both updateChoice = 1 and 2) *******************
		LogData IE, strLogFileName, "Original Settings"
		
		If updateChoice=1 Then
			'*************** Test to make sure it is a "New" status ***************
			strCurrentStatus = .getElementById("drpStatus").options(.getElementById("drpStatus").selectedIndex).text
			If Not (strCurrentStatus = "New" or strCurrentStatus = "Select....") Then
				msgValue = MsgBox("Warning, """ & strCurrentStatus & """ is an incorrect status for a new build" &VBLF& _
					"Proceed only if you're sure this PCID's current settings have been" &VBLF& "approved to be overridden. Do you want to proceed?", _
					vbYesNo + vbExclamation, "Error: Wrong Status")
				If (msgValue = vbNo) Then
					isUpdateSuccessful = False
				End If
			End If 'ENDIF: strCurrentStatus == "New"/"Select...."

			'*************** update Current Status, S/N, PC Category ***************
			If (isUpdateSuccessful) Then
				'*************** Set Current Status ***************
				Set IE = findIEWindow(IE)
				updateSelectControlTo IE, "drpStatus", "Built"
				WaitInternetExplorer
				Set IE = findIEWindow(IE)
				 
				'*************** Set Serial Number ***************
				.getElementById("txtSerialNo").Value = arrPCIDData(intCount).Item("NewSerialNumber")
				
				'*************** Set Model/OS ***************
				UpdateModelTo IE, arrPCIDData(intCount).Item("NewModel"), arrModels
				Set IE = findIEWindow(IE)
				
				'*************** Save Changes ***************
				confirmChanges = MsgBox("Do you want to save these changes?" & VBLF & "Click Yes to continue.", vbYesNo, "Confirm changes?")
				If (confirmChanges = vbNo) Then 'No
					.getElementsByName("btnCancel")(0).click
					isUpdateSuccessful = False
				Else
					.getElementsByName("btnSave")(0).click
				End If
				WaitInternetExplorer
				Set IE = findIEWindow(IE)
				
				If (isUpdateSuccessful) Then
					'********************* Re-Open PCID to Log and Print Update Values *********************
					.getElementByID("lvPCList_txtEnduserinformationID").value = arrPCIDData(intCount).Item("PCID")
					Wscript.sleep 500 '******************************************************* forced wait b/c IE seems to be hanging on this spot often
					.getElementByID("lvPCList_btnSearch").click
					'.getElementByID("lvPCList_txtEnduserinformationID").select
					'WshShell.SendKeys "{Enter}"
					WaitInternetExplorer
					Set IE = findIEWindow(IE)

					If .getElementByID("lblTotalPCFound").innerhtml = "Total PC's Found : 1" Then
						.getElementsByName("lvPCList_ctrl0_lnkSRID")(0).click
						WaitInternetExplorer
						Set IE = findIEWindow(IE)
						
						'*************** Append to Log File with New Values *******************
						LogData IE, strLogFileName, "Updated Settings"
						'*************** Write App Log File *******************
						LogApps IE, strLogAppsFileName

						'*******************   Print PCID *******************
						If PRINT_MODE Then
							'MsgBox "DEBUG: Trying to print!"
							'Prints using the system dialog hidden. Default printer and default options will be used
							'Only works with IE8 or below?
							'https://richardspowershellblog.wordpress.com/2007/07/11/printing-from-internet-explorer/
							'http://msdn.microsoft.com/en-us/library/ms691264.aspx
							IE.execwb 6,2
							Set IE = findIEWindow(IE)
						End If 'End If: PRINT_MODE on

						'********************* Close PCID and go back to Search form *************************
						.getElementsByName("btnCancel")(0).click
						WaitInternetExplorer
						Set IE = findIEWindow(IE)
					End If 'End If: PCID opened again
				Else 'isUpdateSuccessful = False
					'MsgBox "User cancelled saving the changes to this PCID." & VBLF & "Moving on to the next PCID."
				End If
		    Else 'isUpdateSuccessful = False
				MsgBox "Updating the PCID was not successful." & VBLF & "Updating the rest of this PCID will now be skipped!"
				'********************* Close PCID and go back to Search form *************************
				.getElementsByName("btnCancel")(0).click
				WaitInternetExplorer
				Set IE = findIEWindow(IE)
				'what else do we do here? log?
			End If
		Else 'if UpdateChoice = 0
			'*************** Write App Log File *******************
			LogApps IE, strLogAppsFileName
			
			MsgBox "This PCID has been logged." & VBLF & "Please press OK to proceed when you are finished viewing and this PCID will close."
			
			'********************* Close PCID and go back to Search form *************************
			.getElementsByName("btnCancel")(0).click
			WaitInternetExplorer
			Set IE = findIEWindow(IE)
		End If 'End if updateChoice = 1
		
		Set IE = findIEWindow(IE)
	Else 'Could not find PCID user entered.
		msgValue = MsgBox("The PCID you entered could not be located." & VBLF & "Do you want to keep running the script at the next PCID?" & VBLF & "Choose No to quit.", _
			vbYesNo + vbExclamation, "Quit?")
		If (msgValue = vbNo) Then 'No
			WScript.quit
		End If
	End If

Next 'End Loop: Loop through all the PCIDs the user wants to process

End With


'                                      -> msgValue = MsgBox("Script done!" & vblf & "Do you want to close the Internet Explorer window?", vbYesNo + vbQuestion, "Close Internet Explorer?")
'                                      -> Set IE = findIEWindow(IE)
'DEBUG: throws error "object required" -> If (msgValue = vbYes) Then IE.QUIT

MsgBox "Script done!"
WScript.quit



'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************





'http://stackoverflow.com/questions/23299134/failproof-wait-for-ie-to-load-vbscript
'http://msdn.microsoft.com/en-us/library/aa768329(v=vs.85).aspx
'Use:
'IE.Navigate strURL 'or, any other action that causes an IE.Document to reload

'Sub IE_DocumentComplete(byval pdisp, byval URL) is a built-in function for Internet Explorer that fires automatically when the page is loaded
'So, the variable "isIELoadCompleted" is being used as a global to watchdog when the IE_DocumentComplete event fires
Sub IE_DocumentComplete(byval pdisp, byval URL)
    'msgbox "IE completed has fired"
	'if URL = strURL then isIELoadCompleted = true
	isIELoadCompleted = true
End Sub

Sub WaitInternetExplorer
	isIELoadCompleted = false
	Do Until isIELoadCompleted
		WScript.sleep 100
	Loop
End Sub


'Sub Wait(IE)
'	Do While (IE is Nothing)
'		WScript.Sleep 200
'	Loop
'	Do While IE.ReadyState < 4 or IE.Busy
'		WScript.Sleep 200
'	Loop
'End Sub
'
'
'Sub LongWait(IE)
'  Do
'    WScript.Sleep 750
'  Loop While IE.ReadyState < 4 And IE.Busy
'End Sub

Function findIEWindow(oldIE)	
	Dim foundIE
	Dim tryAgain
	tryAgain = True
	Do While (tryAgain)
		If countIEWindows() <> 1 Then
			MsgBox("Error: There is more than one Internet Explorer window open!")
			Set foundIE = Nothing
		Else
			Set foundIE = getOpenIE("PC Details")
		End If
		
		If Not IsObject(foundIE) Then
			MsgBox "Did not find the Internet Explorer window after refresh!" & VBLF & VBLF & "Do you want to quit?. Choose NO to try again!"
			If confirmQuit Then WScript.quit
		Else
			tryAgain = False
		End If
	Loop

	'msgbox "findIEWindow: before LongWait foundIE"
	'LongWait foundIE
	'msgbox "findIEWindow: after LongWait foundIE"
	Set findIEWindow = foundIE
End Function

Function countIEWindows()
	Dim objShell, objWindows, objWindow
	Dim count
	count = 0
	Set objShell = CreateObject("Shell.Application")
    Set objWindows = objShell.Windows
    For Each objWindow In objWindows
        If LCase(Right(objWindow.FullName, 12)) = "iexplore.exe" Then
			count = count + 1
        End If
    Next
	countIEWindows = count
End Function

Function getOpenIE(strWindowTitle)
'see: http://www.experts-exchange.com/Programming/Languages/Visual_Basic/VB_Script/Q_25223158.html
    Dim objShell, objWindows, objWindow, objToReturn
	Set objShell = CreateObject("Shell.Application")
    Set objWindows = objShell.Windows
    For Each objWindow In objWindows
        If LCase(Right(objWindow.FullName, 12)) = "iexplore.exe" Then
			If (objWindow.document.title = strWindowTitle) Then
				Set objToReturn = objWindow
				Exit For
			End If
        End If
    Next
	If IsObject(objToReturn) Then
		'msgbox objToReturn.document.title
		Set getOpenIE = objToReturn
	Else
		'msgbox "nothing found"
		Set getOpenIE = Nothing
	End If
End Function

'DEBUG: NO ERROR CHECKING IF THE STR IS NOT FOUND; FUNCTION JUST QUITS THE SCRIPT!
Function findIndexOf(objIE, strSelectID, strItemText)
	Dim objOption
	Dim intFoundIndex
	intFoundIndex = -1
	For Each objOption In objIE.Document.getElementById(strSelectID).options
		If objOption.Text = strItemText Then
			intFoundIndex = objOption.Index
			Exit For
		End If
	Next
	If intFoundIndex = -1 Then
		MsgBox "The item '" & strItemText & "' was not found in Select box '" & strSelectID & "'!" & VBLF & VBLF & "Quit the script if you can't find it either!"
		If confirmQuit Then WScript.quit
	Else
		'MsgBox "Found index of " & strItemText & ": " & intFoundIndex
		findIndexOf = intFoundIndex
	End If
End Function

Function updateSelectControlTo(objIE, strSelectID, strItemToSet)
	Dim indexOfItem
	indexOfItem = findIndexOf(objIE, strSelectID, strItemToSet)
	objIE.Document.getElementById(strSelectID).selectedIndex = indexOfItem
	'LongWait objIE
	poke objIE, strSelectID
End Function

'causes the Internet Explorer web page to refresh (by registering an onChange event)
Function poke(objIE, strSelectID)
	objIE.Document.getElementById(strSelectID).fireEvent "onChange"
End Function

Function UpdateModelTo(objIE, intModelChoice, arrModels)
	'arrModels(1-max).Item("WMIC")
	'commonName,WMIC,PCCategory,make,model,operatingSystem
	'changed when the array bounds of arrModels changed to 0-x  If (CInt(intModelChoice) >= 1 AND CInt(intModelChoice) < UBound(arrModels)) Then
	If (CInt(intModelChoice) >= 0 AND CInt(intModelChoice) < UBound(arrModels)) Then
		updateSelectControlTo objIE, "drpPCCategory", arrModels(intModelChoice).Item("PCCategory")
		WaitInternetExplorer
		Set objIE = findIEWindow(objIE)
		
		updateSelectControlTo objIE, "drpMake", arrModels(intModelChoice).Item("make")
		WaitInternetExplorer
		Set objIE = findIEWindow(objIE)
		
		updateSelectControlTo objIE, "drpModel", arrModels(intModelChoice).Item("model")
		Set objIE = findIEWindow(objIE)
		
		updateSelectControlTo objIE, "drpOperatingSystem", arrModels(intModelChoice).Item("operatingSystem")
	Else
		MsgBox "Invalid model choice number provided to UpdateModelTo function.", 48, "Error"
	End If
End Function

Sub LogData(objIE, strFileName, strTitle)
	Dim currentPCID,currentSN,modifiedDateTime,currentModel,currentOs,currentStatus,Title
	'*******************************   Collect info from IE   ****************************
	Title = strTitle
	currentPCID=objIE.Document.getElementById("lblPCID").innerhtml
	currentSN=objIE.Document.getElementsByName("txtSerialNo")(0).Value
	modifiedDateTime=objIE.Document.getElementsByName("ddtStatusChangedOn$ddtStatusChangedOn")(0).Value
	currentModel=objIE.Document.getElementById("drpModel").options(objIE.Document.getElementById("drpModel").selectedIndex).text
	currentOs=objIE.Document.getElementById("drpOperatingSystem").options(objIE.Document.getElementById("drpOperatingSystem").selectedIndex).text
	currentStatus=objIE.Document.getElementById("drpStatus").options(objIE.Document.getElementById("drpStatus").selectedIndex).text
	
	'*******************************     write log file    *******************************
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFileLog
	Set objFileLog = objFSO.OpenTextFile(strFileName, vbFileForAppending, True)
	
	objFileLog.WriteLine "==================================="  
	objFileLog.WriteLine Title
	objFileLog.WriteLine "==================================="
	objFileLog.WriteLine "Last Modified:" & modifiedDateTime
	objFileLog.WriteLine "PCID:" & currentPCID
	objFileLog.WriteLine "Serial Number:" & currentSN
	objFileLog.WriteLine "Model:" & currentModel
	objFileLog.WriteLine "Operating System:" & currentOs
	objFileLog.WriteLine "Status:" & currentStatus
End Sub

Sub LogApps(objIE, strFileName)
	'*************** Write App Log File *******************
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFileLogApps
	Set objFileLogApps = objFSO.CreateTextFile(strFileName, True)
	
	Dim currentModel,currentOs,currentStatus
	currentModel=objIE.Document.getElementById("drpModel").options(objIE.Document.getElementById("drpModel").selectedIndex).text
	currentOs=objIE.Document.getElementById("drpOperatingSystem").options(objIE.Document.getElementById("drpOperatingSystem").selectedIndex).text
	currentStatus=objIE.Document.getElementById("drpStatus").options(objIE.Document.getElementById("drpStatus").selectedIndex).text
	objFileLogApps.WriteLine currentModel
	objFileLogApps.WriteLine currentOs
	objFileLogApps.WriteLine currentStatus
	
	Dim objOption
	For Each objOption in objIE.Document.getElementById("lstPartNew").Options
		objFileLogApps.WriteLine objOption.text
	Next	
End Sub


'Function WaitForElementAndReadValue(objIE, strElementID)
'	Do While (isEmpty(objIE.Document.getElementByID(strElementID).value))
'		Sleep 250
'	Loop
'	WaitForElementAndReadValue = objIE.Document.getElementByID(strElementID).value
'End Function
'
'Function WaitForElement(objIE, strElementID)
'	Do While (isEmpty(objIE.Document.getElementByID(strElementID).value))
'		Sleep 250
'	Loop
'End Function

		
Function confirmQuit()
	Dim vbYesNo, vbExclamation, msgValue
	vbYesNo = 4
	vbExclamation = 48
	msgValue = MsgBox("Are you sure you want to quit?", _
		vbYesNo + vbExclamation, "Quit?")
	If (msgValue = vbNo) Then 'No
		confirmQuit = False
	Else
		confirmQuit = True
	End If
End Function

Function getUserInputInteger(strPrompt, intMin, intMax)
	Dim intUserInput
	intUserInput = 0
	'msgbox "DEBUG: " & intUserInput & ", " & intMin & "-" & intMax
	Dim isInputSuccesful
	isInputSuccesful = false
	Do While (NOT isInputSuccesful)
		intUserInput = InputBox(strPrompt)
		'MsgBox "DEBUG - intUserInput: " & intUserInput
		'MsgBox "DEBUG - CInt(" & intUserInput & ") >= " & intMin & ": " & (CInt(intUserInput) >= intMin)
		'MsgBox "DEBUG - CInt(" & intUserInput & ") <= " & intMax & ": " & (CInt(intUserInput) <= intMax)
		'MsgBox "DEBUG - ((CInt(" & intUserInput & ") >= " & intMin & ") And (CInt(" & intUserInput & ") <= " & intMax & ")): " & ((CInt(intUserInput) >= intMin) And (CInt(intUserInput) <= intMax))
		If intUserInput = "" Then
		'	msgbox """"""
			If confirmQuit Then WScript.quit
		'ElseIf intUserInput = Null Then
		'	msgbox "Null"
		'	If confirmQuit Then WScript.quit
		ElseIf (NOT IsNumeric(intUserInput)) Then
			MsgBox """" & intUserInput & """ is not a numeric input.", 48, "Error"
		ElseIf ((CInt(intUserInput) >= intMin) And (CInt(intUserInput) <= intMax)) Then
			isInputSuccesful = true
		Else
			MsgBox intUserInput & " is not a number between " & intMin & " and " & intMax & ". Please try again.", 48, "Error"
		End If
	Loop
	getUserInputInteger = intUserInput
End Function

Function getUserInputString(strPrompt, intLength)
	Dim strUserInput
	strUserInput = ""
	Dim isInputSuccesful
	isInputSuccesful = false
	Do While (NOT isInputSuccesful)
		strUserInput = InputBox(strPrompt)
		'MsgBox "DEBUG - strCurrentNewSerialNumber: " & strCurrentNewSerialNumber
		If strUserInput = "" Then
			If confirmQuit Then WScript.quit
		ElseIf (Len(strUserInput) = intLength) Then
			isInputSuccesful = true
		Else
			MsgBox "That input is not " & intLength & " characters long. Please try again.", 48, "Error"
		End If
	Loop
	getUserInputString = strUserInput
End Function