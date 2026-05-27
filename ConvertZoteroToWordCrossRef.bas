Public Sub ConvertZoteroToWordCrossRef()
    
    ' 1. SPEED & SAFETY OPTIMIZATIONS
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' 2. PULL ALL NUMBERED LISTS
    Dim refItems As Variant
    refItems = doc.GetCrossReferenceItems(wdRefTypeNumberedItem)
    
    If IsEmpty(refItems) Then
        MsgBox "No numbered list found! Please make sure your bibliography is a Word Numbered List.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim fld As field
    Dim i As Long, fieldCount As Long, citationCount As Long
    fieldCount = 0
    citationCount = 0
    
    ' 3. LOOP THROUGH FIELDS BACKWARDS
    For i = doc.Fields.Count To 1 Step -1
        Set fld = doc.Fields(i)
        
        If InStr(1, fld.code.Text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0 Then
            fieldCount = fieldCount + 1
            
            ' =======================================================
            ' PROGRESS BAR FIX
            ' Yields to the OS to refresh the status bar without crashing
            ' =======================================================
            If fieldCount Mod 2 = 0 Then
                Application.StatusBar = "Converting citation block " & fieldCount & " | Cross-references built: " & citationCount
                DoEvents
            End If
            
            Dim citeRange As Range
            Set citeRange = fld.result.Duplicate
            
            ' UNLINK ZOTERO (Turns into plain text like [1][2][3])
            fld.Unlink
            
            ' Find all numbers inside this citation block
            Dim matches As Collection
            Set matches = New Collection
            
            Dim searchRng As Range
            Set searchRng = citeRange.Duplicate
            
            Do
                With searchRng.Find
                    .ClearFormatting
                    .Text = "^#" ' Find any digit
                    .Forward = True
                    .Wrap = wdFindStop
                    
                    If .Execute And searchRng.End <= citeRange.End Then
                        searchRng.MoveEndWhile Cset:="0123456789"
                        matches.Add searchRng.Duplicate
                        
                        searchRng.Collapse wdCollapseEnd
                        searchRng.End = citeRange.End
                    Else
                        Exit Do
                    End If
                End With
            Loop
            
            ' LOOP THROUGH MATCHES BACKWARDS (so text shifting doesn't break positions)
            Dim j As Long
            For j = matches.Count To 1 Step -1
                Dim targetRng As Range
                Set targetRng = matches(j)
                
                Dim numStr As String
                numStr = targetRng.Text
                
                Dim refIdx As Long
                refIdx = GetRefIndex(refItems, numStr)
                
                If refIdx > 0 Then
                    Dim bibText As String
                    bibText = Trim(refItems(refIdx))
                    
                    ' DOUBLE BRACKET FIX
                    If Left(bibText, 1) = "[" Then
                        If targetRng.Characters.First.Previous.Text = "[" And targetRng.Characters.Last.Next.Text = "]" Then
                            targetRng.Start = targetRng.Start - 1
                            targetRng.End = targetRng.End + 1
                        End If
                    End If
                    
                    targetRng.Text = "" ' Delete plain text
                    
                    ' Insert native Word Cross-Reference
                    targetRng.InsertCrossReference _
                        ReferenceType:=wdRefTypeNumberedItem, _
                        ReferenceKind:=wdNumberNoContext, _
                        ReferenceItem:=refIdx, _
                        InsertAsHyperlink:=True
                        
                    citationCount = citationCount + 1
                End If
            Next j
        End If
    Next i
    
    ' 4. CLEANUP
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Application.StatusBar = "SUCCESS! Converted " & citationCount & " citations to native Word Cross-References."
    MsgBox "Successfully converted " & citationCount & " Zotero citations to Word Cross-References!", vbInformation

End Sub

' =======================================================
' HELPER FUNCTION: MATCH NUMBER EXACTLY
' =======================================================
Function GetRefIndex(refItems As Variant, citeNum As String) As Long
    Dim i As Long
    Dim itemText As String
    
    ' Loop BACKWARDS so bibliography items are prioritized
    For i = UBound(refItems) To LBound(refItems) Step -1
        itemText = Trim(refItems(i))
        
        If Left(itemText, Len("[" & citeNum & "]")) = "[" & citeNum & "]" Or _
           Left(itemText, Len(citeNum & ".")) = citeNum & "." Or _
           Left(itemText, Len(citeNum & ")")) = citeNum & ")" Or _
           Left(itemText, Len(citeNum & " ")) = citeNum & " " Or _
           Left(itemText, Len(citeNum & vbTab)) = citeNum & vbTab Then
            GetRefIndex = i
            Exit Function
        End If
    Next i
    
    GetRefIndex = 0
End Function

