Attribute VB_Name = "Module6"
'-----------------------------------------------------------------------------
' LtrCol and ColLtr functions used for VBA hacking the Cells() function to allow letters:
'-----------------------------------------------------------------------------
' Original functions found on the web, somewhere in space and time...
'
'
' Documentation:
' If you're in management and use to be a programmer, then there a good chance you have run across Visual Basic code in your Excel workbooks.
' It's a dirty little secret of every large corporation. This language won't die, in fact, it's probably essential to running most Fortune 500
' companies.  It may be old and forgotten, but here are few tricks to making it a bit easier when your jamming out data sheets,
' munching through numbers, crunching your way through another day...
'
' Feel free to write me at om@orion.consulting with questions, comments concerns
'
' Orion Tip #1: Making the Cells() function a bit more friendly with ColLtr() and LtrCol()
' Do you ever use the Cells(row, col) function?  Wouldn't it be nice if you could specify a letter instead of a number for the column?
' So instead of this: Cells(1, 10) you could do this: Cells(1, "G")
' Well now you can, here's how:
'
' Pattern 1: Simple column names
' Cells(1, LtrCol("H")) = "this is row 1, column H"
'
' Pattern 2: Iterating across Columns
' Do you ever need to work from Column Z to AE and wonder: "okay, what is the number for column AE, so that I can get there."  Insetead use this:
'
' For ltrCol("Z") to ltrCol("AE")
'      Cells(row, col) = "This is row: " & row & " col:  " & ColLtr(col)
' Next
'
' Pattern 3: Picking out random columns
' Now you can get creative... Say you wanted to work with Row 1, Column A, F, G, Y, and Z. Now we can do that in a more readable fashion with the LtrCol trick like so:
'
' For Each col In Split("A,F,G,Y,Z", ",")
'      Cells(row, LtrCol(col)) = "This is row: " & row & " col:  " & col
' Next
'
'-----------------------------------------------------------------------------

Function LtrCol(ByVal InputLetter) As Integer
    OutputNumber = 0
    For i = 1 To Len(InputLetter)
        OutputNumber = (Asc(UCase(Mid(InputLetter, i, 1))) - 64) + OutputNumber * 26
    Next
    LtrCol = OutputNumber
End Function

Function ColLtr(ByVal iCol As Long, Optional sCol As String = "") As String
    If iCol = 0 Then
        ColLtr = sCol
    Else
        sCol = Chr(65 + (iCol - 1) Mod 26) & sCol
        iCol = (iCol - 1) \ 26
        ColLtr = ColLtr(iCol, sCol)
    End If
End Function
