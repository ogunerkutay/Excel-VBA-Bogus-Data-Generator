Option Explicit

' Function to get a random column number
Function GetRandomColumn(ws As Worksheet) As Integer
    On Error GoTo ErrorHandler
    Dim usedColumns As Long
    usedColumns = ws.UsedRange.Columns.Count
    GetRandomColumn = Application.WorksheetFunction.RandBetween(1, usedColumns)
Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Function

' Function to get the last row of a column
Function GetLastRow(ws As Worksheet, col As Integer) As Integer
    On Error GoTo ErrorHandler
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
Exit Function
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Function

' Function to get a random row
Function GetRandomRow(lastRow As Integer) As Integer
    On Error GoTo ErrorHandler
    GetRandomRow = Application.WorksheetFunction.RandBetween(1, lastRow)
Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Function

' Function to convert an input string into a mixed case randomly
Function RandomMixedCase(strInput As String) As String
    Dim i As Integer
    Dim char As String
    Dim result As String

    ' Iterating over each character in the input string
    For i = 1 To Len(strInput)
        char = Mid(strInput, i, 1)
        ' Randomly converting the character to lower or upper case
        If Rnd < 0.5 Then
            result = result & UCase(char)
        Else
            result = result & LCase(char)
        End If
    Next i

    RandomMixedCase = result
End Function

' Function to process the capitalization of an input string based on a random choice
Function ProcessCapitalization(strInput As String) As String
    Dim capitalization As Variant
    Dim strResult As String

    ' Define possible capitalization options
    capitalization = Array("Lowercase", "UPPERCASE", "MixedCase")

    strResult = strInput

    ' Selecting capitalization type randomly and processing the input string accordingly
    Select Case capitalization(Application.WorksheetFunction.RandBetween(LBound(capitalization), UBound(capitalization)))
        Case "Lowercase"
            strResult = LCase(strResult)
        Case "UPPERCASE"
            strResult = UCase(strResult)
        Case "MixedCase"
            strResult = RandomMixedCase(strResult)
    End Select

    ProcessCapitalization = strResult
End Function

' Function to generate an array string with randomized capitalization
Function GenerateArrayString(arrItems As Variant) As String
    ' Generate the array string
    GenerateArrayString = "[" & Chr(34) & Join(arrItems, Chr(34) & "," & Chr(34)) & Chr(34) & "]"
End Function

' Function to get a random date format from an array of date formats.
Function GetRandomDateFormat() As String
    
    Dim dateFormats(1 To 45) As String
    dateFormats(1) = "dd.MM.yyyy"
    dateFormats(2) = "dd.MM.yyyy HH:mm:ss"
    dateFormats(3) = "d.MM.yyyy HH:mm:ss"
    dateFormats(4) = "MM/dd/yyyy"
    dateFormats(5) = "dd/MM/yyyy"
    dateFormats(6) = "yyyy-MM-dd"
    dateFormats(7) = "yyyy/MM/dd"
    dateFormats(8) = "dd-MM-yyyy"
    dateFormats(9) = "MM-dd-yyyy"
    dateFormats(10) = "M/d/yyyy"
    dateFormats(11) = "d/M/yyyy"
    dateFormats(12) = "yyyy-M-d"
    dateFormats(13) = "yyyy/M/d"
    dateFormats(14) = "d-M-yyyy"
    dateFormats(15) = "M-d-yyyy"
    dateFormats(16) = "yyyyMd"
    dateFormats(17) = "dMyyyy"
    dateFormats(18) = "Mdy"
    dateFormats(19) = "yyyyMd"
    dateFormats(20) = "MMddyy"
    dateFormats(21) = "ddMMyy"
    dateFormats(22) = "yyMMdd"
    dateFormats(23) = "dd/MM/yy"
    dateFormats(24) = "MM/dd/yy"
    dateFormats(25) = "yy-MM-dd"
    dateFormats(26) = "dd-MM-yy"
    dateFormats(27) = "MM-dd-yy"
    dateFormats(28) = "yyyy/MM/dd HH:mm:ss"
    dateFormats(29) = "MM/dd/yyyy HH:mm:ss"
    dateFormats(30) = "dd/MM/yyyy HH:mm:ss"
    dateFormats(31) = "yyyy-MM-dd HH:mm:ss"
    dateFormats(32) = "dd-MM-yyyy HH:mm:ss"
    dateFormats(33) = "MM-dd-yyyy HH:mm:ss"
    dateFormats(34) = "yyyy/MM/dd HH:mm"
    dateFormats(35) = "MM/dd/yyyy HH:mm"
    dateFormats(36) = "dd/MM/yyyy HH:mm"
    dateFormats(37) = "yyyy-MM-dd HH:mm"
    dateFormats(38) = "dd-MM-yyyy HH:mm"
    dateFormats(39) = "MM-dd-yyyy HH:mm"
    dateFormats(40) = "M/d/yy"
    dateFormats(41) = "d/M/yy"
    dateFormats(42) = "yy-M-d"
    dateFormats(43) = "yy/M/d"
    dateFormats(44) = "d-M-yy"
    dateFormats(45) = "M-d-yy"
   
  ' We use the RandBetween function to select a random index from the array of date formats.
    GetRandomDateFormat = dateFormats(Int((UBound(dateFormats) - LBound(dateFormats) + 1) * Rnd + LBound(dateFormats)))
  
End Function

Function GetRandomWords(wsWords As Worksheet, lastRow As Integer, randomLanguage As Integer) As String
    GetRandomWords = ProcessCapitalization(wsWords.Cells(GetRandomRow(lastRow), randomLanguage).Value)
End Function

Function GetRandomArray() As String
    Dim wsArrayValues As Worksheet
    Dim lastRow As Integer
    Dim randomNumberofArrayItems As Integer
    Dim arrValues() As Variant
    Dim i As Integer
    
    Set wsArrayValues = ThisWorkbook.Worksheets("ArrayValues")
    
    ' Choose a random number of values between 1 and the number of used rows in column A
    lastRow = GetLastRow(wsArrayValues, 1)
    randomNumberofArrayItems = GetRandomRow(lastRow)
    
    ' Populate the arrValues array with the random values
    ReDim arrValues(randomNumberofArrayItems - 1)
    For i = 0 To randomNumberofArrayItems - 1
        'Assuming ProcessCapitalization and GetRandomRow are valid functions, uncomment the next line
        arrValues(i) = ProcessCapitalization(wsArrayValues.Cells(GetRandomRow(lastRow), 1).Value)
    Next i
    
    ' Generate the array string and return it
    GetRandomArray = GenerateArrayString(arrValues)
End Function


' Function to generate Boolean data in different languages
Function GetRandomBoolean() As Variant

    Dim ws As Worksheet
    Dim options As Variant
    Dim language As Integer

    ' Assuming TrueFalse is the worksheet with the true/false values
    Set ws = ThisWorkbook.Sheets("TrueFalse")

    ' Choosing a language randomly
    language = Application.WorksheetFunction.RandBetween(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)

    ' Defining the true/false values in the chosen language
    options = Array(ws.Cells(1, language).Value, ws.Cells(2, language).Value)

    ' Returning a random Boolean value
    GetRandomBoolean = ProcessCapitalization(CStr(options(Application.WorksheetFunction.RandBetween(LBound(options), UBound(options)))))
    
End Function

' Function to return a random decimal number.
Function GetRandomDecimal(min As Integer, max As Integer) As Double
  ' We use the Rnd function multiplied by the RandBetween function to generate a random decimal number.
  ' The Round function is used to limit the number of decimal places to 2.
  GetRandomDecimal = Round(Rnd * Application.WorksheetFunction.RandBetween(min, max), 2)
End Function


' Function to get a random date within a year from the current date.
Function GetRandomDate() As Date
  ' We use the RandBetween function to get a random date within a year from the current date.
  ' We add the Rnd function to get a random time of the day.
  ' GetRandomDate = Format(Application.WorksheetFunction.RandBetween(CLng(Date), CLng(DateAdd("y", 1, Date))) + Rnd(), GetRandomDateFormat())
    Dim TempDate As Date
    Dim RandDateFormat As String
    Dim FormattedDate As String
  
    RandDateFormat = GetRandomDateFormat()
    TempDate = Application.WorksheetFunction.RandBetween(CLng(Date), CLng(DateAdd("y", 1, Date))) + Rnd()

    FormattedDate = Format(TempDate, RandDateFormat)

    If FormattedDate = "" Then
        MsgBox "Error with date format: " & RandDateFormat
        GetRandomDate = CVErr(xlErrValue)
    Else
        GetRandomDate = FormattedDate
    End If
  
End Function

' Main Subroutine to generate random data
Sub GenerateData()
    On Error GoTo ErrorHandler
    ' Declare variables
    Dim ws As Worksheet
    Dim wsWords As Worksheet
    Dim rowCount As Integer
    Dim i As Integer
    Dim randomLanguage As Integer
    Dim lastRow As Integer
    
    ' Set references to worksheets
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set wsWords = ThisWorkbook.Worksheets("Words")

    ' Turn off screen updating for performance
    Application.ScreenUpdating = False

    ' Define number of rows for which data is to be generated
    rowCount = 1000

    ' Generate data
    ' Loop for each row
    For i = 1 To rowCount

        ' Generating String Data
        randomLanguage = GetRandomColumn(wsWords)
        
        lastRow = GetLastRow(wsWords, randomLanguage)
        
        ws.Cells(i, 1).Value = GetRandomWords(wsWords, lastRow, randomLanguage)
        ' Generating Array Data
        ws.Cells(i, 2).Value = GetRandomArray()

        ' Generating Boolean Data
        ws.Cells(i, 3).Value = GetRandomBoolean()

        ' Generating Decimal Data
        ws.Cells(i, 4).Value = GetRandomDecimal(-1000, 1000)

        ' Generating Date Data
        ws.Cells(i, 5).Value = GetRandomDate()

    Next i

    ' Turn on screen updating
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "Data generation complete.", vbInformation
    Exit Sub
        
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description
    Resume CleanExit
    
End Sub



