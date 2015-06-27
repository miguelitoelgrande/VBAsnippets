


   Sub FindTables()
      Dim iResponse As Integer
      Dim tTable As Table
      'If any tables exist, loop through each table in collection.
      For Each tTable In ActiveDocument.Tables
         tTable.Select
         iResponse = MsgBox("Table found. Find next?", 68)
         If response = vbNo Then Exit For 'User chose to leave search.
      Next
      MsgBox prompt:="Search Complete.", buttons:=vbInformation
   End Sub



   Sub ForceAllTableWidth()
    
   
      Dim iResponse As Integer
      Dim tTable As Table
      'If any tables exist, loop through each table in collection.
      For Each tTable In ActiveDocument.Tables
         tTable.Select
         
         Selection.Rows.HeightRule = wdRowHeightExactly
    Selection.Rows.Height = CentimetersToPoints(2)
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = CentimetersToPoints(5)
    Selection.Cells.PreferredWidthType = wdPreferredWidthPoints
    Selection.Cells.PreferredWidth = CentimetersToPoints(5)
   ' zusaetzlich:
    Selection.Cells.SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    Selection.Cells.SetHeight RowHeight:=CentimetersToPoints(2), HeightRule:=wdAdjustNone
        Selection.Rows.SetHeight RowHeight:=CentimetersToPoints(2), HeightRule:=wdAdjustNone
            Selection.Rows.HeightRule = wdRowHeightExactly
    Selection.Rows.Height = CentimetersToPoints(2)
         'iResponse = MsgBox("Table found. Find next?", 68)
         ' If response = vbNo Then Exit For 'User chose to leave search.
      Next
      MsgBox prompt:="Search Complete.", buttons:=vbInformation
   End Sub



'remove CrLF in Excel -> substitute ^p with @@@@
Sub ReinsertLF()
Dim TheCell As Range
For Each TheCell In ActiveSheet.UsedRange
With TheCell
.Value = Replace(.Value, "@@@@", vbLf)
End With
Next TheCell
End Sub


Public Function SearchVDM(Content)
   SearchVDM = "-"
   Pos = InStr(Content, "vdm:")
   If Pos > 0 Then
       temp = Mid(Content, Pos + 4, 30)
       Pos = InStr(temp, vbLf)
       If Pos > 0 Then
           Pos = Pos
       Else
          Pos = Len(temp) + 1
       End If
       temp = Trim(Left(temp, Pos - 1))

       SearchVDM = temp
   End If
End Function



Public Function FindToken(Content, Token)
   FindToken = "-"
   Pos = InStr(Content, Token)
   If Pos > 0 Then
       temp = Mid(Content, Pos + Len(Token), 256)
       Pos = InStr(temp, vbLf)
       If Pos > 0 Then
           Pos = Pos
       Else
          Pos = Len(temp) + 1
       End If
       temp = Trim(Left(temp, Pos - 1))

       FindToken = temp
   End If
End Function

' Testet eine Liste (Range) von Begriffen auf das Vorkommen in "content".
' der erste Treffer wird zurückgeliefert...
Public Function InContextOf(Content, Wordlist As Range)
   InContextOf = ""
   
   Content = UCase(Content)
              
   For Each Word In Wordlist
      If Word = "" Then Resume
    
      If Content Like "*" + UCase(Word) + "*" Then
          InContextOf = Word
          Exit For
      End If
      
   Next Word
   
    InContextOf = Word

End Function
-----------------------------------------------

Function TokenIndex(token As String, searchRange As Range, Optional nth As Integer = 1)
' Returns the index of the [nth] occurence of a token within a Range of cells (rows!)
' similar to MATCH, however finds a token within a cell (not entire value, but separated
' with whitespaces, ;, etc. but NOT underscore (_) or further characters)...
' Begin (^) or End ($) of the content are extra cases
'
' see also:
' http://www.regular-expressions.info/vbscript.html
' =RegExpFind(Q3;"([^a-zA-Z0-9_]|^)("&A3&")([,?;]|[^a-zA-Z0-9_]|$)")

   ' Regular Expressions in Excel...
   'Dim reg As New VBScript_RegExp_55.RegExp
   Set reg = CreateObject("vbscript.RegExp")
   ' reg.Pattern = <PatternString>
   ' myBoolean = reg.Test(<MainString>)

   Dim cnt As Integer    ' count the cells...
   Dim cnt2 As Integer  ' count the occurrence
   Dim curCell As Range
   Dim curString As String
   Dim result As Object
        
   cnt = 0
   
   For Each curCell In searchRange.Cells
       cnt = cnt + 1
   '    MsgBox ("([^a-zA-Z0-9_])(" & token & ")([,?;]|[^a-zA-Z0-9_])")
       curString = curCell.Value
       
      ' MsgBox (cnt & "." & curString)
       reg.Pattern = "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"
       
     '  MsgBox (RegExpFind(curString, "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"))
'
       If reg.Test(curString) Then
           TokenIndex = cnt
          ' MsgBox ("exiting..." & cnt)
           Exit Function
       End If              
   Next curCell
    TokenIndex = -1
End Function


------------------------------------------------------------

Option Explicit
#Const LateBind = True
# http://www.tmehta.com/regexp/add_code.htm

Function RegExpSubstitute(ReplaceIn, _
        ReplaceWhat As String, ReplaceWith As String)
    #If Not LateBind Then
    Dim RE As RegExp
    Set RE = New RegExp
    #Else
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
        #End If
    RE.Pattern = ReplaceWhat
    RE.Global = True
    RegExpSubstitute = RE.Replace(ReplaceIn, ReplaceWith)
    End Function
Function RegExpFind(FindIn, FindWhat As String, _
        Optional IgnoreCase As Boolean = False)
    Dim i As Long
    #If Not LateBind Then
    Dim RE As RegExp, allMatches As MatchCollection, aMatch As Match
    Set RE = New RegExp
    #Else
    Dim RE As Object, allMatches As Object, aMatch As Object
    Set RE = CreateObject("vbscript.regexp")
        #End If
    RE.Pattern = FindWhat
    RE.IgnoreCase = IgnoreCase
    RE.Global = True
    Set allMatches = RE.Execute(FindIn)
    ReDim rslt(0 To allMatches.Count - 1)
    For i = 0 To allMatches.Count - 1
        rslt(i) = allMatches(i).Value
        Next i
    RegExpFind = rslt
    End Function




Public Function Tokenize(Content, tokenlist As Range)
      
   Set reg = CreateObject("vbscript.RegExp")
   result = ""
              
   For Each token In tokenlist
       reg.Pattern = "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"
     '  MsgBox (RegExpFind(curString, "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"))
     
       If reg.Test(Content) Then
           result = result & token & ";"
       End If
      
   Next token
      
   Tokenize = result

End Function


Function TokenIndexMulti(token As String, searchRange As Range, Optional nth As Integer = 1)

' Multi Matches...

' Returns the index of the [nth] occurence of a token within a Range of cells (rows!)

' similar to MATCH, however finds a token within a cell (not entire value, but separated

' with whitespaces, ;, etc. but NOT underscore (_) or further characters)...

' Begin (^) or End ($) of the content are extra cases

' see also:

' http://www.regular-expressions.info/vbscript.html

'

' =RegExpFind(Q3;"([^a-zA-Z0-9_]|^)("&A3&")([,?;]|[^a-zA-Z0-9_]|$)")

 

   ' Regular Expressions in Excel...

   'Dim reg As New VBScript_RegExp_55.RegExp

   Set reg = CreateObject("vbscript.RegExp")

   ' reg.Pattern = <PatternString>

   ' myBoolean = reg.Test(<MainString>)

 

   Dim cnt As Integer    ' count the cells...

   Dim cnt2 As Integer  ' count the occurrence

   Dim curCell As Range

   Dim curString As String

   Dim result As Object

        

   cnt = 0

   

   TokenIndexMulti = ""

   For Each curCell In searchRange.Cells

       cnt = cnt + 1

   '    MsgBox ("([^a-zA-Z0-9_])(" & token & ")([,?;]|[^a-zA-Z0-9_])")

       curString = curCell.Value

       

      ' MsgBox (cnt & "." & curString)

       reg.Pattern = "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"

       

     '  MsgBox (RegExpFind(curString, "([^a-zA-Z0-9_]|^)(" & token & ")([,?;]|[^a-zA-Z0-9_]|$)"))

'

       If reg.Test(curString) Then

           TokenIndexMulti = TokenIndexMulti & cnt & ";"

          ' MsgBox ("exiting..." & cnt)

          ' Exit Function

       End If

   Next curCell

   ' Backward compatible return value for nothing...

   If TokenIndexMulti = "" Then TokenIndexMulti = -1

End Function

 

Function IndexMulti(LookupColumn As Range, indexlist As String)

' like Excels Index function, but its supposed to show multiple lines regarding to the indexlist

' (e.g. "2;4;6;") as one string value with multiple lines...

 

   IndexMulti = ""

   For Each myIdx In Split(indexlist, ";")

     If myIdx = "" Or myIdx = -1 Then Exit Function

     'MsgBox (myIdx)

     IndexMulti = IndexMulti & "[" & myIdx & "] " & Application.Index(LookupColumn, myIdx) & vbCrLf

   Next myIdx

   If IndexMulti = "" Then IndexMulti = "-"

End Function





