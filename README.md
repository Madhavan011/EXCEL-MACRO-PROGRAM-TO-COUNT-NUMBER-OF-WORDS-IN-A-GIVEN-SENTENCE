# EXCEL-MACRO-PROGRAM-TO-COUNT-NUMBER-OF-WORDS-IN-A-GIVEN-SENTENCE
Sub CountWordsInSelectedRange()

Dim rng As Range, cell As Range

Dim cellWords As Integer, totalWords As Integer

Dim content As String

' Set the selected range

Set rng = Selection

totalWords = 0

' Loop through each cell in the range

For Each cell In rng

If Not cell.HasFormula Then

content = Trim(cell.Value)

If content = "" Then

cellWords = 0

Else

cellWords = 1

' Count spaces to determine number of words

Do While InStr(content, " ") > 0

content = Mid(content, InStr(content, " ") + 1)

content = Trim(content)

cellWords = cellWords + 1

Loop
End If

totalWords = totalWords + cellWords

End If

Next cell

' Display the total word count

MsgBox totalWords & " words found in the selected range.", vbInformation, "Word Count"

End Sub
