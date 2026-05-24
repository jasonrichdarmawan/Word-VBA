Public Sub ZoteroLinkCitation()
    ' 1. SPEED & SAFETY OPTIMIZATIONS
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    ActiveWindow.View.ShowFieldCodes = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim fld As field
    Dim bibRange As Range
    
    ' 2. LOCATE THE BIBLIOGRAPHY
    For Each fld In doc.Fields
        If InStr(1, fld.code.Text, "ADDIN ZOTERO_BIBL", vbTextCompare) > 0 Then
            Set bibRange = fld.result.Duplicate
            doc.Bookmarks.Add Name:="Zotero_Bibliography", Range:=bibRange
            Exit For
        End If
    Next fld
    
    If bibRange Is Nothing Then
        MsgBox "Could not find a Zotero Bibliography.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 3. SETUP VARIABLES
    Dim fieldCode As String, title As String, searchTitle As String, titleAnchor As String
    Dim n1 As Long, n2 As Long
    Dim citeSearchRange As Range, bibSearchRange As Range, linkRange As Range
    Dim fieldCount As Long, citationCount As Long
    Dim newLink As Hyperlink
    
    ' Array variables for sorting Zotero's mismatched JSON
    Dim titles() As String
    Dim titleCount As Long
    Dim i As Long
    
    fieldCount = 0
    citationCount = 0
    
    ' 4. MAIN LOOP
    For Each fld In doc.Fields
        fieldCount = fieldCount + 1
        
        If fieldCount Mod 5 = 0 Then
            Application.StatusBar = "Processing field " & fieldCount & " of " & doc.Fields.Count & " | Links: " & citationCount
            DoEvents
        End If
        
        If InStr(1, fld.code.Text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0 Then
            fieldCode = fld.code.Text
            Set citeSearchRange = fld.result.Duplicate
            
            ' =======================================================
            ' FIX 4: PRE-LOAD ALL TITLES INTO AN ARRAY
            ' We must extract them all first so we can sort them
            ' =======================================================
            titleCount = 0
            ReDim titles(0 To 100) ' Max 100 citations per bracket
            
            Do While InStr(1, fieldCode, """title"":""") > 0
                n1 = InStr(1, fieldCode, """title"":""") + Len("""title"":""")
                n2 = InStr(n1, fieldCode, """,""")
                If n2 = 0 Then n2 = InStr(n1, fieldCode, """}")
                If n2 = 0 Then n2 = InStr(n1, fieldCode, """")
                
                If n2 > n1 Then
                    titles(titleCount) = Mid(fieldCode, n1, n2 - n1)
                End If
                
                titleCount = titleCount + 1
                fieldCode = Mid(fieldCode, n2 + 1)
            Loop
            
            ' Re-size array to actual count
            If titleCount > 0 Then ReDim Preserve titles(0 To titleCount - 1)
            
            ' =======================================================
            ' THE CORE FIX: MAP VISUAL NUMBERS TO BIBLIOGRAPHY FIRST
            ' =======================================================
            ' Instead of trusting Zotero's JSON order, we read the visual numbers
            ' left-to-right on the screen, search the bibliography for each number,
            ' grab the title next to that number, and create the link.
            
            With citeSearchRange.Find
                .ClearFormatting
                .Text = "^#" ' Find first digit on screen (e.g. "1")
                .Forward = True
                .Wrap = wdFindStop
                
                Do While .Execute
                    citeSearchRange.MoveEndWhile Cset:="0123456789"
                    Set linkRange = citeSearchRange.Duplicate
                    
                    ' We now know the exact number we are looking at (e.g. "10")
                    Dim visualNumber As String
                    visualNumber = linkRange.Text
                    
                    ' Find this exact number in the bibliography to get the CORRECT title
                    Set bibSearchRange = bibRange.Duplicate
                    With bibSearchRange.Find
                        .ClearFormatting
                        .Text = "[" & visualNumber & "]" ' Looks for [10] in the bib
                        .Forward = True
                        .Wrap = wdFindStop
                        .MatchCase = False
                        
                        If .Execute Then
                            ' We found [10] in the bibliography!
                            ' Now expand the range to grab the title next to it.
                            bibSearchRange.Expand wdParagraph
                            searchTitle = bibSearchRange.Text
                            
                            ' Clean the bibliography text to generate a valid bookmark
                            If InStr(searchTitle, ":") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, ":") - 1)
                            If InStr(searchTitle, "?") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, "?") - 1)
                            If InStr(searchTitle, "—") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, "—") - 1)
                            
                            ' Strip out the [10] part from the string
                            searchTitle = Replace(searchTitle, "[" & visualNumber & "]", "")
                            searchTitle = Trim(searchTitle)
                            
                            ' Strip LaTeX and special characters
                            searchTitle = Replace(searchTitle, "\", "")
                            searchTitle = Replace(searchTitle, "$", "")
                            searchTitle = Replace(searchTitle, "{", "")
                            searchTitle = Replace(searchTitle, "}", "")
                            
                            titleAnchor = MakeValidBMName(searchTitle) & "_" & citationCount
                            citationCount = citationCount + 1
                            
                            ' Bookmark the bibliography entry
                            doc.Bookmarks.Add Name:=titleAnchor, Range:=bibSearchRange
                            
                            ' Create the hyperlink
                            Set newLink = doc.Hyperlinks.Add( _
                                Anchor:=linkRange, _
                                Address:="", _
                                SubAddress:=titleAnchor, _
                                ScreenTip:="", _
                                TextToDisplay:=linkRange.Text)
                            
                            ' Vault completely over the new hyperlink's hidden codes
                            citeSearchRange.Start = newLink.Range.End
                        Else
                            ' If bib search fails, still step past the number
                            citeSearchRange.Start = linkRange.End
                        End If
                    End With
                    citeSearchRange.End = fld.result.End
                Loop
            End With
        End If
    Next fld

    ' 5. CLEANUP
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Application.StatusBar = "SUCCESS! " & citationCount & " citations perfectly ordered and linked."

End Sub

Function MakeValidBMName(strIn As String) As String
    Dim i As Long, tempStr As String, pFirstChr As String
    strIn = Trim(strIn)
    
    For i = 1 To Len(strIn)
        Select Case Asc(Mid$(strIn, i, 1))
            Case 49 To 57, 65 To 90, 97 To 122
                tempStr = tempStr & Mid$(strIn, i, 1)
        End Select
    Next i
    
    If Len(tempStr) = 0 Then tempStr = "Cite"
    pFirstChr = Left(tempStr, 1)
    If Not pFirstChr Like "[A-Za-z]" Then tempStr = "A_" & tempStr
    
    MakeValidBMName = Left(tempStr, 30)
End Function
