Private Sub CommandButton3_Click()


    ' Declare variables
    Dim fileName As String
    Dim filePath As String
    Dim directoryPath As String
    Dim textLine As String
    Dim keyword1 As String, keyword2 As String
    Dim charCount1 As Integer
    Dim position1 As Integer, position2 As Integer
    Dim ws As Worksheet
    Dim nextRowA As Long, nextRowB As Long, nextRowC As Long
    Dim wordsArray() As String
    Dim firstNameLastName As String
    Dim remainingText As String
    Dim lastShares As String
    Dim sharesPosition As Integer
    Dim tempString As String

    ' Prompt for the file name
    fileName = InputBox("Enter the name of the text file:", "File Name")
    If fileName = "" Then Exit Sub ' If no file name is entered, exit

    ' Set the directory path (modify this to the directory you want)
    directoryPath = "C:\Users\jbennett\Desktop\Test\" ' Replace with your directory path

    ' Construct the full file path
    filePath = directoryPath & fileName & ".txt"

    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found. Please check the file name and try again.", vbExclamation
        Exit Sub
    End If

    ' Set the keywords and character count to search for
    keyword1 = "Acct/Pt."  ' Replace with your first keyword
    charCount1 = 16        ' Set the number of characters to extract after the first keyword

    keyword2 = "UPAL -"  ' Replace with your second keyword

    ' Set the worksheet to place data
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the first empty rows in columns A, B, and C
    nextRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    nextRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1
    nextRowC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row + 1

    ' Open the text file
    Open filePath For Input As #1

    ' Initialize lastShares as an empty string
    lastShares = ""

    ' Read through the text file line by line
    Do Until EOF(1)
        Line Input #1, textLine
        Debug.Print "Reading Line: "; textLine ' Debugging: Print each line read

        ' *** Handle Keyword1 Processing ***
        position1 = 1
        Do While position1 > 0
            position1 = InStr(position1, textLine, keyword1)
            If position1 > 0 Then
                ws.Cells(nextRowA, 1).Value = Mid(textLine, position1 + Len(keyword1), charCount1)
                nextRowA = nextRowA + 1

                ' Check for Keyword 2 in the same line after finding Keyword 1
                position2 = InStr(position1 + Len(keyword1), textLine, keyword2)
                If position2 > 0 Then
                    ' If Keyword 2 is found, skip further processing for this line
                Else
                    ' If Keyword 2 is not found, copy the rest of the line (after the first keyword1 and the 16 characters) to Column B
                    remainingText = Mid(textLine, position1 + Len(keyword1) + charCount1)
                    ws.Cells(nextRowB, 2).Value = Trim(remainingText)
                    nextRowB = nextRowB + 1
                End If
                ' Update position1 to search for the next occurrence of Keyword 1
                position1 = position1 + Len(keyword1) + charCount1
            End If
        Loop

        ' *** Handle Keyword2 Processing ***
        position2 = 1
        Do While position2 > 0
            position2 = InStr(position2, textLine, keyword2)
            If position2 > 0 Then
                ' Extract the text after the keyword
                remainingText = Mid(textLine, position2 + Len(keyword2))
                remainingText = Trim(remainingText)

                ' Split the remaining text into words and concatenate all words
                wordsArray = Split(remainingText, " ")

                firstNameLastName = wordsArray(0)
                Dim i As Integer
                For i = 1 To UBound(wordsArray)
                    firstNameLastName = firstNameLastName & " " & wordsArray(i)
                Next i

                ws.Cells(nextRowB, 2).Value = firstNameLastName

                nextRowB = nextRowB + 1
                position2 = position2 + Len(keyword2) + Len(firstNameLastName) ' Move position to continue search
            End If
        Loop

        ' *** Handle Shares Processing ***
        If InStr(1, textLine, "Shares", vbTextCompare) > 0 Then
            ' Find the position of "Shares"
            sharesPosition = InStr(1, textLine, "Shares", vbTextCompare)
            
            ' Extract up to 7 characters before "Shares"
            tempString = Trim(Mid(textLine, sharesPosition - 7, 7))
        
            ' Check if extracted string is numeric
            If IsNumeric(tempString) Then
                lastShares = tempString ' Update lastShares to the most recent valid number
                Debug.Print "Found Shares value: "; lastShares ' Debugging: Print the last shares value found
            End If
        End If

        ' *** Handle Confidential Processing ***
        If InStr(1, Trim(textLine), "Confidential", vbTextCompare) > 0 Then
            Debug.Print "Found Confidential: "; textLine ' Debugging: Print when "Confidential" is found
            ' If "Confidential" is found, write the last "Shares" value to Excel in Column C
            If lastShares <> "" Then
                ws.Cells(nextRowC, 3).Value = CDbl(lastShares)
                Debug.Print "Writing to Excel: "; lastShares ' Debugging: Print what is being written to Excel
                nextRowC = nextRowC + 1
            End If

            ' Reset lastShares after writing to Excel to ensure the next "Confidential" gets a new set
            lastShares = "" ' Make sure we reset it here after writing
        End If

    Loop

    ' Close the text file
    Close #1

    MsgBox "Parsing complete. All relevant data has been recorded in Columns A, B, and C."
End Sub
