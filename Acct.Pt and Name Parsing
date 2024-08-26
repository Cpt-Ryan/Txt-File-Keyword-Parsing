Private Sub CommandButton1_Click()
    Dim fileName As String
    Dim filePath As String
    Dim directoryPath As String
    Dim textLine As String
    Dim keyword1 As String, keyword2 As String
    Dim charCount1 As Integer
    Dim position1 As Integer, position2 As Integer
    Dim ws As Worksheet
    Dim nextRowA As Long, nextRowB As Long
    Dim wordsArray() As String
    Dim firstNameLastName As String
    Dim remainingText As String

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

    ' Set the keywords and character count to search for
    keyword1 = "Acct/Pt."  ' Replace with your first keyword
    charCount1 = 16        ' Set the number of characters to extract after the first keyword

    keyword2 = "UPAL -"  ' Replace with your second keyword

    ' Set the worksheet to place data
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the first empty rows in columns A and B
    nextRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    nextRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1

    ' Open the text file
    Open filePath For Input As #1

    ' Read through the text file line by line
    Do Until EOF(1)
        Line Input #1, textLine

        ' Search for all instances of the first keyword in the line
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

        ' Search for all instances of the second keyword in the line
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

    Loop

    ' Close the text file
    Close #1

    MsgBox "Parsing complete."
End Sub
