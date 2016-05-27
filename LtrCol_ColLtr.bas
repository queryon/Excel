Attribute VB_Name = "LtrCol_ColLtr"
'-----------------------------------------------------------------------------
' LtrCol and ColLtr functions used for VBA hacking the Cells() function to allow letters:
'
' Note: Original functions below found on the web, somewhere in space and time long forgotten.  
'
'
' If you're in management, accounting or just playing around in the business world then there a good chance you have run across 
' Visual Basic code in your Excel workbooks.  This language won't die, in fact, it's probably essential to running 'most Fortune 500 companies.  
' It may be old and forgotten, but here are few tricks to making it a bit easier when your jamming out data sheets, munching through numbers, 
' crunching your way through another day...
'
' Orion Tip #1: Making the Cells() function a bit more friendly with ColLtr() and LtrCol()
'
' Do you ever use the Cells(row, col) function?  Wouldn't it be nice if you could specify a letter instead of a number for the column?  
'
' So instead of this: Cells(1, 10) you could do this: Cells(1, "J")
'
' Well now you can, here's how:
'
' Pattern 1: Simple column names
'	Cells(1, ltrCol("H")) = "this is row 1, column H"
'	Note: Yes, you can use Range("H1") to do this, but we are talking about Cells() function, and Range can't be used for everything take pattern 2 for example:
'
' Pattern 2: Iterating across Columns
'   Do you ever need to work from Column Z to AE and wonder: "okay, what is the number for column AE?  And you end up with something like this for col = 26 to 34?  Who's going to remember that column 34 was AE? Nobody. So try this instead:
'
'	For col = ltrCol("Z") to ltrCol("AE")
'	      Cells(row, col) = "This is row: " & row & " col:  "  & ColLtr(col)
'	Next
'
' Pattern 3: Picking out random columns
'	Now you can get creative. Say you wanted to work with Row 1, Column A, F, G, Y, and Z. Now we can do that in a more readable fashion with the LtrCol trick like so:
'
'	For Each col in split("A,F,G,Y,Z", ",")
'		Cells(row, LtrCol(col)) = "This is row: " & row & " col:  "  & col
'	Next
'
' Yes, Range() is powerful too, so don't forget that you can use that for some things. 
' But sometimes you just want Cells() for one reason or another.  
' And now you have a trick up your sleeve when the time comes.  Happy Excel hacking! 
'
' Here's the code: https://github.com/orionconsulting/Excel 
'
' - Orion 
'
' @omatthews om@orion.consulting 
'
' About
' Orion Consulting is a lean-by-design American software development company. 
' Formed in 2015, our globally distributed team of awesome programmers work 
' with established companies, foundations, and non-profits. http://orion.consulting 
'



'-----------------------------------------------------------------------------
' LtrCol
' Converts a Letter to an Excel Column number
'-----------------------------------------------------------------------------
Function LtrCol(ByVal InputLetter) As Integer
    OutputNumber = 0
    For i = 1 To Len(InputLetter)
        OutputNumber = (Asc(UCase(Mid(InputLetter, i, 1))) - 64) + OutputNumber * 26
    Next
    LtrCol = OutputNumber
End Function

'-----------------------------------------------------------------------------
' ColLtr
' Converts a column number to an Excel Column Letter 
' Note: This function uses recursion, so don't worry about the 
' second parameter which is used during that process...
'-----------------------------------------------------------------------------
Function ColLtr(ByVal iCol As Long, Optional sCol As String = "") As String
    If iCol = 0 Then
        ColLtr = sCol
    Else
        sCol = Chr(65 + (iCol - 1) Mod 26) & sCol
        iCol = (iCol - 1) \ 26
        ColLtr = ColLtr(iCol, sCol)
    End If
End Function
