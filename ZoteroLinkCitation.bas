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
            
            Do While InStr(1, fieldCode, """title"":""") > 0
                citationCount = citationCount + 1
                
                ' TITLE EXTRACTION
                n1 = InStr(1, fieldCode, """title"":""") + Len("""title"":""")
                n2 = InStr(n1, fieldCode, """,""")
                If n2 = 0 Then n2 = InStr(n1, fieldCode, """}")
                If n2 = 0 Then n2 = InStr(n1, fieldCode, """")
                
                If n2 > n1 Then
                    title = Mid(fieldCode, n1, n2 - n1)
                Else
                    title = "Unknown"
                End If
                
                title = CleanZoteroTitle(title)
                
                ' =======================================================
                ' FIX 2: BULLETPROOF SEARCH TITLE FOR BIBLIOGRAPHY
                ' Chop off at the first colon, question mark, or dash
                ' This prevents the R-ROME non-breaking space bug!
                ' =======================================================
                searchTitle = title
                If InStr(searchTitle, ":") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, ":") - 1)
                If InStr(searchTitle, "?") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, "?") - 1)
                If InStr(searchTitle, "—") > 0 Then searchTitle = Left(searchTitle, InStr(searchTitle, "—") - 1)
                searchTitle = Trim(searchTitle)
                If Len(searchTitle) > 20 Then searchTitle = Left(searchTitle, 20)
                
                titleAnchor = MakeValidBMName(title) & "_" & citationCount
                
                ' NUMBER EXTRACTION
                With citeSearchRange.Find
                    .ClearFormatting
                    .Text = "^#" ' Find first digit
                    .Forward = True
                    .Wrap = wdFindStop
                    
                    If .Execute Then
                        citeSearchRange.MoveEndWhile Cset:="0123456789"
                        Set linkRange = citeSearchRange.Duplicate
                        
                        Set bibSearchRange = bibRange.Duplicate
                        With bibSearchRange.Find
                            .ClearFormatting
                            .Text = searchTitle
                            .Forward = True
                            .Wrap = wdFindStop
                            .MatchCase = False
                            
                            If .Execute Then
                                doc.Bookmarks.Add Name:=titleAnchor, Range:=bibSearchRange
                                
                                ' =======================================================
                                ' FIX 1: REMOVED SCREENTIP (TOOLTIP)
                                ' Prevents Mac Word PDF exporter from corrupting URLs
                                ' =======================================================
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
                            citeSearchRange.End = fld.result.End
                        End With
                    End If
                End With
                
                ' Move JSON parser forward
                fieldCode = Mid(fieldCode, n2 + 1)
            Loop
        End If
    Next fld

    ' 5. CLEANUP
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Application.StatusBar = "SUCCESS! " & citationCount & " citations linked perfectly for Mac PDF export."

End Sub

Function CleanZoteroTitle(str As String) As String
    Dim res As String
    res = str
    res = Replace(res, "<i>", "", 1, -1, vbTextCompare)
    res = Replace(res, "</i>", "", 1, -1, vbTextCompare)
    res = Replace(res, "<span class=""nocase"">", "", 1, -1, vbTextCompare)
    res = Replace(res, "<span class=\""nocase\"">", "", 1, -1, vbTextCompare)
    res = Replace(res, "</span>", "", 1, -1, vbTextCompare)
    CleanZoteroTitle = res
End Function

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

