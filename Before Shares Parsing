Private Sub CommandButton2_Click()
    Dim fileName As String
    Dim filePath As String
    Dim directoryPath As String
    Dim textLine As String
    Dim lastShares As String
    Dim ws As Worksheet
    Dim nextRowC As Long
    Dim sharesPosition As Integer
    Dim tempString As String

    ' Prompt for the file name
    fileName = InputBox("Enter the name of the text file:", "File Name")
    If fileName = "" Then Exit Sub ' If no file name is entered, exit

    ' Set the directory path (you can modify this to the directory you want)
    directoryPath = "C:\Users\jbennett\Desktop\Test\" ' Replace with your directory path

    ' Construct the full file path
    filePath = directoryPath & fileName & ".txt"

    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found. Please check the file name and try again.", vbExclamation
        Exit Sub
    End If

    ' Set the worksheet to place data
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the first empty row in column C
    nextRowC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row + 1

    ' Open the text file
    Open filePath For Input As #1

    ' Initialize lastShares as an empty string
    lastShares = ""

    ' Read through the text file line by line
    Do Until EOF(1)
        Line Input #1, textLine
        Debug.Print "Reading Line: "; textLine ' Debugging: Print each line read

        ' Check if the line contains "Shares"
        If InStr(1, textLine, "Shares", vbTextCompare) > 0 Then
            ' Find the position of "Shares"
            sharesPosition = InStr(1, textLine, "Shares", vbTextCompare)
            
            ' Extract up to 9 characters before "Shares"
            tempString = Trim(Mid(textLine, sharesPosition - 7, 7))
        
            ' Check if extracted string is numeric
            If IsNumeric(tempString) Then
                lastShares = tempString ' Update lastShares to the most recent valid number
                Debug.Print "Found Shares value: "; lastShares ' Debugging: Print the last shares value found
            End If
        End If

        ' Check if the line contains "Confidential"
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

    MsgBox "Processing complete. The last 'Shares' values before each 'Confidential' entry have been recorded in Column C."
End Sub
