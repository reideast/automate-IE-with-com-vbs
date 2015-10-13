Option Explicit


'NOTE: THIS DESCRIPTION IS VALID FOR THE OLD VERSION, NOT THE DICTIONARY-BASED VERSION

'Function CSVtoArray
'argument: strFileName - Path to the .csv file to read
'return value: an array of arrays, containing the individual comma separated values

'Description:
'reads the CSV file located at strFileName
'returns an array formatted as such:
'array(num lines)
'each element of array is itself an array() of strings of each item on that lines
'Ex:
'items.CSV
' firstName,lastName,ID
' John,Doe,1234
' Jane,Doe,4321
'will return: array(3), which equals:
'array(0) = "firstName", "lastName", "ID"
'array(1) = "John", "Doe", "1234"
'array(2) = "Jane", "Doe", "4321"

'Dim arrCSVFile, arrCSVLine, strValue, dictionaryLine, arrLine, arrKeys, intCount
'arrCSVFile = CSVtoArray("CSVtoArray_test.csv")
'for each dictionaryLine in arrCSVFile
'  arrKeys = dictionaryLine.Keys
'  arrLine = dictionaryLine.Items
'  intCount = 0
'  for each strValue in arrLine
'    MsgBox arrKeys(intCount) & ": " & strValue
'    intCount = intCount + 1
'  next
'Next

Function CSVtoArray(strFileName)
  Const vbRead = 1', vbWrite = 2
  Dim intRow, intColumn
  Dim objFSO, fileCSV
  Dim arrWholeCSVFile() ', arrCSVLine, strValue
  Dim intNumLines
  Dim arrStrHeaderRow
  Dim intNumColumns
  Dim strLine, arrLine, strColumn
  Set objFSO = CreateObject("Scripting.FileSystemObject")

  'count the lines in the file
  Set fileCSV = objFSO.OpenTextFile(strFileName, vbRead)
  fileCSV.ReadAll
  intNumLines = fileCSV.Line
  intNumLines = intNumLines - 1  'account to array zero-index
  intNumLines = intNumLines - 1  'don't count header row
  fileCSV.Close
  Redim arrWholeCSVFile(intNumLines)
  
  're-open text file (back at the top)
  Set fileCSV = objFSO.OpenTextFile(strFileName, vbRead)
  
  'read first line to find out column headers
  strLine = fileCSV.ReadLine
  arrStrHeaderRow = Split(strLine, ",")
  intNumColumns = UBound(arrStrHeaderRow)

  'read each line
  intRow = 0
  Do While fileCSV.AtEndOfStream <> True
    strLine = fileCSV.ReadLine
	'msgbox "DEBUG: Row #" & intRow & VBLF & strLine
    arrLine = Split(strLine, ",")
    Set arrWholeCSVFile(intRow) = CreateObject("Scripting.Dictionary")
    intColumn = 0
    For Each strColumn in arrLine
	  If intColumn > intNumColumns Then
	    MsgBox "Error in CSVtoArray: Too many columns in row number " & intRow + 1 & ". Extra columns will be skipped."
      Else
	    arrWholeCSVFile(intRow).Add arrStrHeaderRow(intColumn), strColumn
        intColumn = intColumn + 1
	  End If
    Next
    intRow = intRow + 1
    'MsgBox "Line " & intRow & ": '" & strLine & "'"
  Loop
  fileCSV.Close

  CSVtoArray = arrWholeCSVFile
End Function
