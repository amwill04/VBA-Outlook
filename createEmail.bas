Attribute VB_Name = "createEmail"
Public Function stringHTML(Optional ByRef rng As Variant, Optional ByRef expression As Variant, _
                           Optional ByVal bold As Variant) As String
' --------------------------------------------------------------------------------------------------
'   Puclic function that takes the arguement of either a range or string expression and converts
'   to single line html next with <br> splitting. Bold parameter takes True or False and determines
'   if the range or string is to be converted into bold text.
' --------------------------------------------------------------------------------------------------
'   As either range or string can be used the if statements test to make sure one has been provided.

    If IsMissing(rng) And IsMissing(expression) Then
        MsgBox "A Range or single text string must be included"
    Exit Function
    End If
    If Not IsMissing(rng) And Not IsMissing(expression) Then
        MsgBox "Only use a Range or a an Expression. Not both."
    Exit Function
    End If
' --------------------------------------------------------------------------------------------------
    If IsMissing(bold) Then
        bold = False
    End If

    If Not IsMissing(rng) Then
        For Each element In rng.Cells
            If Not IsEmpty(element) Then
                stringHTML = stringHTML & element & "<br>"
            End If
        Next
        If bold = True Then
        stringHTML = "<b>" & stringHTML & "</b>"
        End If
    End If

    If Not IsMissing(expression) Then
        stringHTML = stringHTML & expresion & "<br>"
        If bold = True Then
        stringHTML = "<b>" & stringHTML & "</b>"
        End If
    End If

End Function

Public Function tableHTML(ByRef rng As Range) As String
' ---------------------------------------------------------------------------------------------------
'   Function that generates HTML table for embedding into email
' ---------------------------------------------------------------------------------------------------
'   Function Dimension

    Dim eachRow     As Range
' ---------------------------------------------------------------------------------------------------
' Loop over each row assigning <tr> and <td> to cells

    For Each eachRow In rng.Rows
        tableHTML = tableHTML & "<tr>"
            For Each element In eachRow.Cells
                If Not IsEmpty(element) Then
                tableHTML = tableHTML & "<td>" & element & "</td>"
                End If
            Next
        tableHTML = tableHTML & "</tr>"
    Next

    tableHTML = "<table border=1 style=""width:50%"">" & tableHTML & "</table>"

End Function


Sub test()

MsgBox Environ$("temp")

End Sub
