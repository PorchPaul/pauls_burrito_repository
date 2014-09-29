Attribute VB_Name = "Module1"
'here we declare our variable we can access across multiple functions
Dim howmanyburritos As Long

Sub giveEveryoneBurritos(ByRef howmanyburritos)
'this function distributes burritos among the data ops team
'it can't be run independently - it requires the input of "howmanyburritos."
'the function "getBurritos" populates the variable "howmanyburritos" and then calls this function.

'declare some local variables
Dim burritoRange As Range
Dim startingBurritos As Long
Dim endingBurritos As Long
Dim burritomessage As String
Dim i As Integer

'Store in memory the range of cells in the spreadsheet we're working with
Set burritoRange = Worksheets("how many burritos").Range("a2:b16")

burritoRange.Select

'reset the count of burritos distributed
totalburritos = 0

'give each person more burritos

For i = 1 To burritoRange.Rows.Count ' <-- beginning of the loop. For ease of navigation, i've indented everything inside the loop
    
    'figure out how many burritos people have to begin with
    'this imports the data in the specified cell in the spreadsheet into memory.
    'The variable 'i' determines the row within the range, starting from the top.
    'the static number 2 specifies that we're looking in the 2nd column from the left.
    
    startingBurritos = burritoRange.Cells(i, 2)  '<--the variable on the left of the = sign (startingBurritos) is populated with the value on the right from the specified cell.
    
    'calculate how many burritos people will get.
    'note: Per business rules, Paul gets double burritos because you don't want him to get hangry
    
    If burritoRange.Cells(i, 1) = "Paul" Then
        endingBurritos = startingBurritos + howmanyburritos * 2
        Else
        endingBurritos = startingBurritos + howmanyburritos
    End If

    'there's no such thing as negative burritos.
    If endingBurritos < 0 Then endingBurritos = 0

    'update spreadsheet with the number of burritos
    burritoRange.Cells(i, 2) = endingBurritos '< -- notice that this time the cell reference is on the left of the = sign. The cell is populated by the value of the variable endingBurritos.

    'add to the total number of burritos distributed
    totalburritos = totalburritos + endingBurritos


Next i '<-- increment the variable "i" and start the loop over

'create a string expresing how many burritos were distributed

If howmanyburritos > 0 Then
    burritomessage = "Good job! the Data Ops team now has " & CStr(totalburritos) & " Burritos."
    Else: burritomessage = "hey where are my burritos? not cool bro"
End If

'display number of burritos in message box
MsgBox (burritomessage)

End Sub

Sub getBurritos()
'this function gathers the number of burritos to distribute per person from cell D2 of the spreadsheet.

    burritosperperson = Worksheets("how many burritos").Range("d2")
    
    Call giveEveryoneBurritos(burritosperperson)

End Sub
