Sub AutoOpen()
    '--- Runs automatically when document is opened
    ActiveWindow.ActivePane.View.Zoom.Percentage = 64
    ActiveWindow.ActivePane.View.Zoom.Percentage = 64
    ActiveWindow.DocumentMap = True
    '--- Because Selection Line below gives error when in ReadingMode
    On Error Resume Next
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.1)
    '--- ReadingMode
    ActiveWindow.View.ReadingLayout = True
    ActiveWindow.View.ReadingLayoutActualView = True
End Sub
Sub fitpagezoom()
    '--- Shortcut ctrl+0
    ActiveWindow.ActivePane.DisplayRulers = Not ActiveWindow.ActivePane. _
        DisplayRulers
    ActiveWindow.DocumentMap = True
    ActiveWindow.ActivePane.View.Zoom.Percentage = 64
End Sub
Sub DeleteLine()
    '--- Shortcut ctrl+x
    If Selection.Type = wdSelectionIP Then
    Selection.HomeKey Unit:=wdLine
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    End If
    Selection.Cut
End Sub
Sub FormatDocument()
    Dim myDoc As Document
    Set myDoc = ActiveDocument

    Dim para As Paragraph
    Dim firstIndent As Double   'value in "points"
    For Each para In myDoc.Paragraphs
        If para.Style Like "Heading*" Then
            ' Getting heading indentation to use on text below
            firstIndent = myDoc.Styles(para.Style).ParagraphFormat.LeftIndent
            ' Setting Heading Case
            Dim i As Integer
            i = 0
            For Each wrd In para.Range.Words
                Dim tempWrd As String
                Dim tempWrdProper As String
                
                ' If first letter of word is lowercase, make the word TitleCase
                ' If the word is already title cased or fully uppercased, nothing will change
                If StrComp(Left(wrd, 1), UCase(Left(wrd, 1)), vbBinaryCompare) = 1 Then
                    wrd.Case = wdTitleWord
                End If
                
                ' Setting Case of unwanted words in Heading
                ' tempWrd Var used in code below to lowercase unwanted words even if they are uppercased (like AND)
                tempWrd = wrd
                tempWrdProper = StrConv(tempWrd, vbProperCase)
                Select Case Trim(tempWrdProper)
                   Case "And", "If", "Then", "At", "The", "Vs", "Vs.", "Of", "For", "Is", "In", "To", "Both", "Up", "With", "Are", "A", "Into", "From", "On", "Off"
                      ' Lower Case unwanted word if it's not the first word
                      If i > 0 Then
                        wrd.Case = wdLowerCase
                      End If
                      ' Title Case unwanted word if it's the first word and Heading is numberred
                        If para.Style = "Heading 2" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                ' If the current unwanted word is the first word after the numbers in this para
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                             End If
                        End If
                        If para.Style = "Heading 3" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                            ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                                If i = 4 Then
                                    para.Range.Words(5).Case = wdTitleWord
                                End If
                            End If
                        End If
                        If para.Style = "Heading 4" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                            ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                                If i = 6 Then
                                    para.Range.Words(7).Case = wdTitleWord
                                End If
                            End If
                        End If
                        If para.Style = "Heading 5" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                            End If
                        End If
                        If para.Style = "Heading 6" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                            End If
                        End If
                        If para.Style = "Heading 7" Then
                            If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                                If i = 2 Then
                                    para.Range.Words(3).Case = wdTitleWord
                                End If
                            End If
                        End If
                        ' End of Setting Case of unwanted word in Heading
                    Case Else
                      ' do nothing
                End Select
                i = i + 1
            Next wrd
        Else
            '--- So that the whole table is indented, not just its contents
            If para.Range.Tables.Count > 0 Then
              para.Range.Tables(1).Rows.LeftIndent = firstIndent
              para.Range.Tables(1).AutoFitBehavior (wdAutoFitContent)
            Else
                para.LeftIndent = firstIndent
            End If
            
            '--- Capitalize first letter in first word of every sentence
            If StrComp(Left(para.Range.Words(1), 1), UCase(Left(para.Range.Words(1), 1)), vbBinaryCompare) = 1 Then
                para.Range.Words(1).Case = wdTitleWord
            End If
        End If
    Next para
    '--- Refresh needed to show the changes just made
    Application.ScreenRefresh
End Sub
Sub NumberHeadings()
    Dim myDoc As Document
    Set myDoc = ActiveDocument
    Dim para As Paragraph
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim x As Integer
    Dim z As Integer
    Dim strInt As String
    Dim forCounter As Integer
    i = 0
    j = 0
    k = 0
    l = 0
    x = 0
    z = 0
    forCounter = 1
    strInt = ""
    
    For Each para In myDoc.Paragraphs
        If para.Style Like "Heading*" Then
            If para.Style = "Heading 1" Then
                i = 0
                j = 0
                k = 0
                l = 0
                x = 0
                z = 0
            End If
            If para.Style = "Heading 2" Then
                i = i + 1
                j = 0
                k = 0
                l = 0
                x = 0
                z = 0
                strInt = "" & i
                If i < 10 Then
                    strInt = "0" & i
                End If
                If StrComp(para.Range.Words(1), "00", vbBinaryCompare) = 0 Then
                    i = i - 1
                ElseIf StrComp(para.Range.Words(1), "0", vbBinaryCompare) = 0 Then
                    i = i - 1
                Else
                    If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                        para.Range.Words(1) = strInt
                    Else
                        para.Range.Words(1) = strInt & "- " & para.Range.Words(1)
                    End If
                End If
            End If
            If para.Style = "Heading 3" Then
                j = j + 1
                k = 0
                l = 0
                x = 0
                z = 0
                strInt = "" & j
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = i & "." & strInt
                ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                    forCounter = 0
                    For Each wrd In para.Range.Words
                        If StrComp(wrd, "- ", vbBinaryCompare) = 1 Then
                            forCounter = forCounter + 1
                        Else
                            Exit For
                        End If
                    Next wrd
                    For y = 1 To forCounter - 1
                        para.Range.Words(1) = ""
                    Next y
                    para.Range.Words(1) = i & "." & strInt
                Else
                    para.Range.Words(1) = i & "." & strInt & "- " & para.Range.Words(1)
                End If
            End If
            If para.Style = "Heading 4" Then
                k = k + 1
                l = 0
                x = 0
                z = 0
                strInt = "" & k
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = i & "." & j & "." & strInt
                ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                    forCounter = 0
                    For Each wrd In para.Range.Words
                        If StrComp(wrd, "- ", vbBinaryCompare) = 1 Then
                            forCounter = forCounter + 1
                        Else
                            Exit For
                        End If
                    Next wrd
                    For y = 1 To forCounter - 1
                        para.Range.Words(1) = ""
                    Next y
                    para.Range.Words(1) = i & "." & j & "." & strInt
                Else
                    para.Range.Words(1) = i & "." & j & "." & strInt & "- " & para.Range.Words(1)
                End If
            End If
            If para.Style = "Heading 5" Then
                l = l + 1
                x = 0
                z = 0
                strInt = "" & l
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = strInt
                Else
                    para.Range.Words(1) = strInt & "- " & para.Range.Words(1)
                End If
            End If
            If para.Style = "Heading 6" Then
                x = x + 1
                z = 0
                strInt = "" & x
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = strInt
                Else
                    para.Range.Words(1) = strInt & "- " & para.Range.Words(1)
                End If
            End If
            If para.Style = "Heading 7" Then
                z = z + 1
                strInt = "" & z
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = strInt
                Else
                    para.Range.Words(1) = strInt & "- " & para.Range.Words(1)
                End If
            End If
        End If
    Next para
    '--- Refresh needed to show the changes just made
    Application.ScreenRefresh
End Sub
Sub Normalize()
'
' Normalize Macro
' Imports Normal Styles and Sets Margins
'
    ActiveDocument.CopyStylesFromTemplate ("C:\Users\Jaad Chacra\AppData\Local\Packages\Microsoft.Office.Desktop_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Templates\Normal.dotm")
    ActiveDocument.PageSetup.LeftMargin = InchesToPoints(0.5)
    ActiveDocument.PageSetup.RightMargin = InchesToPoints(0.5)
    ActiveDocument.PageSetup.TopMargin = InchesToPoints(0.5)
    ActiveDocument.PageSetup.BottomMargin = InchesToPoints(0.5)
End Sub
Sub SelectionIncrementHeadings()
    Dim para As Paragraph
    Call SelectionClearHeaderNumbers
    For Each para In Selection.Range.Paragraphs
        If para.Style Like "Heading*" Then
            If para.Style = "Heading 2" Then
                para.Style = "Heading 1"
            End If
            If para.Style = "Heading 3" Then
                para.Style = "Heading 2"
            End If
            If para.Style = "Heading 4" Then
                para.Style = "Heading 3"
            End If
            If para.Style = "Heading 5" Then
                para.Style = "Heading 4"
            End If
            If para.Style = "Heading 6" Then
                para.Style = "Heading 5"
            End If
            If para.Style = "Heading 7" Then
                para.Style = "Heading 6"
            End If
        End If
    Next para
    '--- Refresh needed to show the changes just made
    Application.ScreenRefresh
    Call FullFormat
End Sub
Sub SelectionDecrementHeadings()
    Dim para As Paragraph
    Call SelectionClearHeaderNumbers
    For Each para In Selection.Range.Paragraphs
        If para.Style Like "Heading*" Then
            If para.Style = "Heading 1" Then
                para.Style = "Heading 2"
            ElseIf para.Style = "Heading 2" Then
                para.Style = "Heading 3"
            ElseIf para.Style = "Heading 3" Then
                para.Style = "Heading 4"
            ElseIf para.Style = "Heading 4" Then
                para.Style = "Heading 5"
            ElseIf para.Style = "Heading 5" Then
                para.Style = "Heading 6"
            ElseIf para.Style = "Heading 6" Then
                para.Style = "Heading 7"
            End If
        End If
    Next para
    '--- Refresh needed to show the changes just made
    Application.ScreenRefresh
    Call FullFormat
End Sub
Sub SelectionClearHeaderNumbers()
    Dim para As Paragraph
    Dim forCounter As Integer
    forCounter = 1
    For Each para In Selection.Range.Paragraphs
        If para.Style Like "Heading*" Then
            If para.Style = "Heading 2" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
            If para.Style = "Heading 3" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                    forCounter = 0
                    For Each wrd In para.Range.Words
                        If StrComp(wrd, "- ", vbBinaryCompare) = 1 Then
                            forCounter = forCounter + 1
                        Else
                            Exit For
                        End If
                    Next wrd
                    For y = 1 To forCounter - 1
                        para.Range.Words(1) = ""
                    Next y
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
            If para.Style = "Heading 4" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                ElseIf StrComp(para.Range.Words(2), ".", vbBinaryCompare) = 0 Then
                    forCounter = 0
                    For Each wrd In para.Range.Words
                        If StrComp(wrd, "- ", vbBinaryCompare) = 1 Then
                            forCounter = forCounter + 1
                        Else
                            Exit For
                        End If
                    Next wrd
                    For y = 1 To forCounter - 1
                        para.Range.Words(1) = ""
                    Next y
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
            If para.Style = "Heading 5" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
            If para.Style = "Heading 6" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
            If para.Style = "Heading 7" Then
                If StrComp(para.Range.Words(2), "- ", vbBinaryCompare) = 0 Then
                    para.Range.Words(1) = ""
                    para.Range.Words(1) = ""
                End If
            End If
        End If
    Next para
    '--- Refresh needed to show the changes just made
    Application.ScreenRefresh
End Sub
Sub RemoveBlankParas()
    Dim oDoc        As Word.Document
    Dim para        As Word.Paragraph
    Dim paraCount   As Integer

    Set oDoc = ActiveDocument
    paraCount = 1

    For Each para In oDoc.Paragraphs
        If Len(para.Range.Text) = 1 Then
            If para.Style Like "Heading*" Then
                para.Range.Delete
                paraCount = paraCount - 1
                GoTo NextIteration
            End If
        Else
            '-- Do Nothing
            GoTo NextIteration
        End If
        ' Everything below here means it's empty line and not heading
        If paraCount = 1 Then
            para.Range.Delete
            paraCount = paraCount - 1
            GoTo NextIteration
        End If
        If paraCount = oDoc.Paragraphs.Count Then
            para.Range.Delete
            paraCount = paraCount - 1
            GoTo NextIteration
        End If
        If paraCount > 1 And paraCount < oDoc.Paragraphs.Count Then
            If oDoc.Paragraphs(paraCount - 1).Style Like "Heading*" Then
                If oDoc.Paragraphs(paraCount + 1).Style Like "Heading*" Then
                    If oDoc.Paragraphs(paraCount - 1).Style = oDoc.Paragraphs(paraCount + 1).Style Then
                        '-- Do nothing
                        GoTo NextIteration
                    Else
                        If oDoc.Paragraphs(paraCount - 1).Style = "Heading 1" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 2" Then
                            '-- Do nothing, we will enter content there since heading 2 after and it's currently a heading
                            GoTo NextIteration
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 3" Then
                            If oDoc.Paragraphs(paraCount - 1).Style = "Heading 2" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                            Else
                                '-- Do nothing
                                GoTo NextIteration
                            End If
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 4" Then
                            '' We didnt add Heading 2 case since we never should have Heading 2 then 4
                            If oDoc.Paragraphs(paraCount - 1).Style = "Heading 3" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                            Else
                                '-- Do nothing
                                GoTo NextIteration
                            End If
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 5" Then
                            If oDoc.Paragraphs(paraCount - 1).Style = "Heading 4" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                            Else
                                '-- Do nothing
                                GoTo NextIteration
                            End If
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 6" Then
                            If oDoc.Paragraphs(paraCount - 1).Style = "Heading 5" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                            Else
                                '-- Do nothing
                                GoTo NextIteration
                            End If
                        ElseIf oDoc.Paragraphs(paraCount + 1).Style = "Heading 7" Then
                            If oDoc.Paragraphs(paraCount - 1).Style = "Heading 6" Then
                                para.Range.Delete
                                paraCount = paraCount - 1
                                GoTo NextIteration
                            Else
                                '-- Do nothing
                                GoTo NextIteration
                            End If
                        Else
                            para.Range.Delete
                            paraCount = paraCount - 1
                            GoTo NextIteration
                        End If
                    End If
                Else
                    ' If before heading and after text, delete this empty line
                    para.Range.Delete
                    paraCount = paraCount - 1
                    GoTo NextIteration
                End If
            End If
            If oDoc.Paragraphs(paraCount + 1).Style Like "Heading*" Then
                If oDoc.Paragraphs(paraCount - 1).Style Like "Heading*" Then
                    If oDoc.Paragraphs(paraCount - 1).Style = oDoc.Paragraphs(paraCount + 1).Style Then
                        '-- Do nothing
                        GoTo NextIteration
                    Else
                        para.Range.Delete
                        paraCount = paraCount - 1
                        GoTo NextIteration
                    End If
                Else
                    para.Range.Delete
                    paraCount = paraCount - 1
                    GoTo NextIteration
                End If
            End If
            If Len(oDoc.Paragraphs(paraCount + 1).Range.Text) = 1 Then
                para.Range.Delete
                paraCount = paraCount - 1
                GoTo NextIteration
            End If
        End If
NextIteration:
        paraCount = paraCount + 1
    Next para
End Sub
Sub AddBlankParas()
    Dim oDoc        As Word.Document
    Dim para        As Word.Paragraph
    Dim paraCount   As Integer
    Dim paraRange   As Word.Range
    Dim Rng         As Range
    Dim HOneCount   As Integer
    Dim HTwoCount   As Integer
    
    Set oDoc = ActiveDocument
    paraCount = 1
    HOneCount = 0
    HTwoCount = 0

    For Each para In oDoc.Paragraphs
        If paraCount < oDoc.Paragraphs.Count Then
            If para.Style Like "Heading*" Then
                ' Increment H1 and H2 Count (we dont want pagebreak before 1st H1 and H2)
                If para.Style = "Heading 1" Then
                    HOneCount = HOneCount + 1
                    HTwoCount = 0
                End If
                If para.Style = "Heading 2" Then
                    HTwoCount = HTwoCount + 1
                End If
                ' Add Blank para (not PageBreak) between Headings to be later removed by another Function
                If oDoc.Paragraphs(paraCount + 1).Style Like "Heading*" Then
                    oDoc.Paragraphs.Add _
                        Range:=ActiveDocument.Paragraphs(paraCount + 1).Range
                    oDoc.Paragraphs(paraCount + 1).Style = "Normal"
                End If
            Else
                '-- If it's a pagebreak, Do Nothing
                With ActiveDocument
                  Set Rng = .Paragraphs(paraCount).Range
                  Rng.End = Rng.End - 2
                  Rng.Collapse wdCollapseEnd
                  If Asc(Rng.Characters.Last) = 12 Then GoTo NextIteration
                End With
                
                If oDoc.Paragraphs(paraCount + 1).Style = "Heading 1" Then
                    If HOneCount = 0 Then
                        GoTo NextIteration
                    End If
                    With ActiveDocument
                        Set paraRange = .Paragraphs(paraCount + 1).Range
                        paraRange.Collapse Direction:=wdCollapseStart
                        paraRange.InsertBreak WdBreakType.wdPageBreak
                    End With
                    GoTo NextIteration
                End If
                                
                If oDoc.Paragraphs(paraCount + 1).Style = "Heading 2" Then
                    If HTwoCount = 0 Then
                        GoTo NextIteration
                    End If
                    With ActiveDocument
                        Set paraRange = .Paragraphs(paraCount + 1).Range
                        paraRange.Collapse Direction:=wdCollapseStart
                        paraRange.InsertBreak WdBreakType.wdPageBreak
                    End With
                    GoTo NextIteration
                End If
            End If
        End If
NextIteration:
        paraCount = paraCount + 1
    Next para
End Sub
Sub FullFormat()
'
' FullFormat Macro
'
'
'--- Shortcut ctrl + shift + j
    Call NumberHeadings
    Call FormatDocument
    
    MsgBox "Format Complete"
End Sub
Sub UltimateFormat()
'
' UltimateFormat Macro
'
' Shortcut = ctrl+j
    Call AddBlankParas
    Call RemoveBlankParas
    Call NumberHeadings
    Call FormatDocument
    
    MsgBox "Format Complete"
End Sub
Sub RemoveNumbering()
'
' RemoveNumbering Macro
'
'
    Dim para        As Word.Paragraph
    Dim i           As Integer
    
    i = 1
    Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    For Each para In Selection.Range.Paragraphs
        para.Range.Words(1) = i & "- " & para.Range.Words(1)
        i = i + 1
    Next para
    
    Call FormatDocument
End Sub