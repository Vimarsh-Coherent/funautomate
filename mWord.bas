Attribute VB_Name = "mWord"

'**********-Add Title

Function AddDocTitle()
    Dim myStr As String
    
    myStr = "Category - " & Sheet1.Range("D2").Value
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 18
        .Font.ColorIndex = wdDarkBlue
    End With
    
    wApp.Selection.TypeParagraph
    
End Function



'**********-Create Market Size and Trend

Function AddMarketSizeAndTrends()
    
    'Market Size and Trends
    wDoc.Activate
    myStr = "Market Size and Trends: "
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Select the third paragraph
    Set rngCopy = mDoc.Paragraphs(3).Range
    wApp.Selection.TypeText Text:=rngCopy.Text
    
    'Insert Driver image at the end of the target document
    ImagePath = FolderPath & "\Impact_Analysis.jpg"
    wDoc.Content.InsertAfter vbCrLf
    wDoc.InlineShapes.AddPicture FileName:=ImagePath, LinkToFile:=False, SaveWithDocument:=True
    
    'Move Cursor to end
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Select the Fourth paragraph
    Set rngCopy = mDoc.Paragraphs(4).Range
    wApp.Selection.TypeText Text:=rngCopy.Text
    'wApp.Selection.TypeParagraph
    
    
End Function



'**********-Create Market Driver

Function AddMarketDriver()
    
    'Market Driver 1
    wDoc.Activate
    myStr = "Market Driver - " & Sheet1.Range("C15").Value
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Drivers of the Market:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Drivers of the Market:") + 1
    End If
        
    
    'Get end postion of "Market Size and Trends"
    With rngCopy.Find
        .Text = "Key Takeaways of Analyst:" 'Sheet1.Range("C16").Value
        .Forward = True
        .Wrap = 1
    End With
    
    eMDocCursor = ""
    If rngCopy.Find.Execute Then
        eMDocCursor = rngCopy.Start
    End If
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    wApp.Selection.TypeParagraph
    
    
    
End Function



'**********-Create Key Takeaway

Function AddKeyTakeAway()
    
    'Market Driver 1
    wDoc.Activate
    myStr = "Key Takeaways of Analyst:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Key Takeaways of Analyst:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Key Takeaways of Analyst:")
    End If
        
    
    'Get end postion of "Market Size and Trends"
    With rngCopy.Find
        .Text = "Market Challenges and Opportunities:" 'Sheet1.Range("C16").Value
        .Forward = True
        .Wrap = 1
    End With
    
    eMDocCursor = ""
    If rngCopy.Find.Execute Then
        eMDocCursor = rngCopy.Start
    End If
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
End Function



'**********-Create Market Challenge

Function AddMarketChallenge()
    
    'Market Driver 1
    wDoc.Activate
    myStr = "Market Challenge - " & Sheet1.Range("C17").Value
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Market Challenges and Opportunities:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Market Challenges and Opportunities:") + 1
    End If
        
    
    'Get end postion of "Market Size and Trends"
    With rngCopy.Find
        .Text = "Segmental Analysis:" 'Sheet1.Range("C16").Value
        .Forward = True
        .Wrap = 1
    End With
    
    eMDocCursor = ""
    If rngCopy.Find.Execute Then
        eMDocCursor = rngCopy.Start
    End If
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
End Function


'**********-Create Segment Analysis

Function AddSegmentAnalysis()
    
    wDoc.Activate
    myStr = "Segmental Analysis:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Insert Driver image at the end of the target document
    ImagePath = FolderPath & "\Segmental_Insights.jpg"
    wDoc.Content.InsertAfter vbCrLf
    wDoc.InlineShapes.AddPicture FileName:=ImagePath, LinkToFile:=False, SaveWithDocument:=True
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Segmental Analysis:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Segmental Analysis:")
    End If
        
    
    'Get end postion of "Market Size and Trends"
    With rngCopy.Find
        .Text = "Regional Analysis:" 'Sheet1.Range("C16").Value
        .Forward = True
        .Wrap = 1
    End With
    
    eMDocCursor = ""
    If rngCopy.Find.Execute Then
        eMDocCursor = rngCopy.Start
    End If
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    

End Function



'**********-Create Regional Analysis

Function AddRegionalAnalysis()
    
    wDoc.Activate
    myStr = "Regional Insights:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Insert Driver image at the end of the target document
    ImagePath = FolderPath & "\Regional_Insights.jpg"
    wDoc.Content.InsertAfter vbCrLf
    wDoc.InlineShapes.AddPicture FileName:=ImagePath, LinkToFile:=False, SaveWithDocument:=True
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Regional Analysis:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Regional Analysis:")
    End If
        
    
    'Get end postion of "Market Size and Trends"
    With rngCopy.Find
        .Text = "Competitive Landscape:" 'Sheet1.Range("C16").Value
        .Forward = True
        .Wrap = 1
    End With
    
    eMDocCursor = ""
    If rngCopy.Find.Execute Then
        eMDocCursor = rngCopy.Start
    End If
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Competitive Landscape
    myStr = "Competitive Landscape:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.TypeParagraph
    
    'Insert Driver image at the end of the target document
    ImagePath = FolderPath & "\Market_KeyPlayer.jpg"
    wDoc.Content.InsertAfter vbCrLf
    wDoc.InlineShapes.AddPicture FileName:=ImagePath, LinkToFile:=False, SaveWithDocument:=True
    
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
End Function



'**********-Create Key Development

Function AddKeyDevelopment()
    Dim sCursor As Long, eCursor As Long
    wDoc.Activate
    
    'Header
    myStr = "Key Developments:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    'Content
    myStr = Sheet1.Range("D40").Value & vbNewLine & Sheet1.Range("D41").Value & vbNewLine & Sheet1.Range("D42").Value & vbNewLine & Sheet1.Range("D43").Value & vbNewLine & Sheet1.Range("D44").Value
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    'wApp.Selection.TypeParagraph
    
End Function



'**********-Create Defination

Function AddDefination()
    
    wDoc.Activate
    myStr = "*Definition:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    'wApp.Selection.TypeParagraph
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Definition:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Definition:")
    End If
        
    
    'Get end postion of "Market Size and Trends"
    wApp.Selection.EndKey unit:=wdStory
    eMDocCursor = wApp.Selection.End - Len("Definition:")
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Italic = True
    
    wDoc.Range(Start:=1, End:=eCursor).ParagraphFormat.Alignment = wdAlignParagraphJustify
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    

End Function



'**********-Create Key takeaway

Function AddKeyTakeAways()
    
    wDoc.Activate
    myStr = "Key Takeaways from Analyst"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    
    For kt = 40 To 44
        If Sheet1.Range("D" & kt).Value <> "" Then
            wApp.Selection.TypeParagraph
            myStr = Sheet1.Range("D" & kt).Value
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=myStr
            eCursor = wApp.Selection.End
            wDoc.Range(Start:=sCursor, End:=eCursor).Font.Italic = wdToggle
        End If
    Next kt
    
    wDoc.Range(Start:=1, End:=eCursor).ParagraphFormat.Alignment = wdAlignParagraphJustify
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    

End Function


'**************** Table

Function CreateDocTable()
    
    'Fill data in Word Table sheet
    'Title
    Sheet5.Range("A1").Value = Sheet1.Range("D2").Value
    
    'Base Year
    Sheet5.Range("B3").Value = Sheet1.Range("D4").Value
    
    'Market Size
    Sheet5.Range("C3").Value = "Market Size in " & Sheet1.Range("D4").Value + 1 & ":"
    Sheet5.Range("D3").Value = Sheet1.Range("D7").Value
    
    'Historial Data for
    Sheet5.Range("B4").Value = Sheet1.Range("D5").Value & " To " & Sheet1.Range("D4").Value - 1
    
    'Forecast Period
    Sheet5.Range("D4").Value = Sheet1.Range("D4").Value + 1 & " To " & Sheet1.Range("D6").Value
    
    'Forecast Period CGR
    Sheet5.Range("A5").Value = "Forecast Period " & Sheet1.Range("D4").Value + 1 & " To " & Sheet1.Range("D6").Value & " CAGR:"
    Sheet5.Range("B5").Value = Sheet1.Range("D9").Text
    
    'Value Project
    Sheet5.Range("C5").Value = Sheet1.Range("D6").Value & " Value Projection:"
    Sheet5.Range("D5").Value = Sheet1.Range("D8").Text
    
    
    'Geopgraphics covered
    TxtContent = ""
    
    For K = 1 To Sheet3.Range("A1").End(xlToRight).Column
        TxtContent = TxtContent & Sheet3.Cells(1, K).Value & ": "
        For r = 2 To Sheet3.Cells(Rows.Count, K).End(xlUp).Row
            TxtContent = TxtContent & Sheet3.Cells(r, K).Value & ", "
        Next r
        TxtContent = Mid(TxtContent, 1, Len(TxtContent) - 2) & Chr(10)
        
    Next K
    
    Sheet5.Range("B6").Value = TxtContent
    
    'Segments Covered
    TxtContent = ""
    
    For K = 8 To Sheet1.Range("H2").End(xlToRight).Column
        If Not Sheet1.Cells(2, K).Value Like "% Market share*" And Sheet1.Cells(3, K).Value <> "" Then
            TxtContent = TxtContent & Sheet1.Cells(2, K).Value & ": "
            
            For r = 3 To Sheet1.Cells(Rows.Count, K).End(xlUp).Row
                il = Len(Sheet1.Cells(r, K).Value) - Len(Replace(Sheet1.Cells(r, K).Value, ">", ""))
                If il <= 1 Then
                    SubContent = ""
                    For Z = r + 1 To 10
                        il = Len(Sheet1.Cells(Z, K).Value) - Len(Replace(Sheet1.Cells(Z, K).Value, ">", ""))
                        If il = 2 Then
                            If SubContent = "" Then
                                SubContent = "(" & Replace(Sheet1.Cells(Z, K).Value, ">", "")
                            Else
                                SubContent = SubContent & ", " & Replace(Sheet1.Cells(Z, K).Value, ">", "")
                            End If
                        Else
                            If SubContent <> "" Then
                                SubContent = SubContent & ")"
                            End If
                            Exit For
                        End If
                        'r = r + 1
                    Next Z
                    TxtContent = TxtContent & Replace(Sheet1.Cells(r, K).Value, ">", "") & " " & SubContent & ", "
                End If
            Next r
            
            TxtContent = Mid(TxtContent, 1, Len(TxtContent) - 2) & Chr(10)
            
        End If
        
    Next K

    Sheet5.Range("B7").Value = TxtContent
    
    'Companies Covered
    
    TxtContent = ""
        
    For r = 3 To Sheet1.Cells(Rows.Count, "G").End(xlUp).Row
        TxtContent = TxtContent & Sheet1.Cells(r, "G").Value & ", "
    Next r
    
    TxtContent = Mid(TxtContent, 1, Len(TxtContent) - 2)
    
    Sheet5.Range("B8").Value = TxtContent
    
    
    'Growth Drivers
    
    TxtContent = Sheet1.Cells(15, "C").Value & Chr(10)
    TxtContent = TxtContent & Sheet1.Cells(16, "C").Value
    
    Sheet5.Range("B9").Value = TxtContent
    
    'Restraint and Challenge
    TxtContent = ""
    TxtContent = Sheet1.Cells(17, "C").Value & Chr(10)
    TxtContent = TxtContent & Sheet1.Cells(18, "C").Value
    
    Sheet5.Range("B10").Value = TxtContent
    
    
    'Move the Table to word doc
    'Page setup
    With wDoc.PageSetup
        .TopMargin = Application.InchesToPoints(1)
        .LeftMargin = Application.InchesToPoints(0.8)
        .RightMargin = Application.InchesToPoints(0.5)
    End With
    
    
    'Paste the range into Word
    Sheet5.Visible = xlSheetVisible
    Sheet5.Range("A1:D10").Copy
    eCursor = wApp.Selection.End
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    'Header
    myStr = "Market Report Scope - After Analyst View"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    'wDoc.Content.InsertAfter vbCrLf
    wDoc.Range(Start:=eCursor).PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    
    Sheet5.Visible = xlSheetVeryHidden
    
    'Adjust the size of the pasted content to fit within the page width
    wDoc.Tables(1).PreferredWidth = wApp.Selection.PageSetup.PageWidth - 90
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    ParaJustifyStart = wApp.Selection.End
    
End Function


'*************** Quetionaries
Function CreateDocQuestionaries()
    Dim Q1 As String, Q2 As String, Q3 As String, Q4 As String, Q5 As String
    
    wDoc.Activate
    wApp.Selection.EndKey unit:=wdStory
    'wApp.Selection.TypeParagraph
    
    'Header
    myStr = "Frequently Asked Questions:"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    
    Q1 = "What will be the CAGR of " & Sheet1.Range("D2").Value & "?" & vbNewLine & "The CAGR of " & Sheet1.Range("D2").Value & " is projected to be " & Sheet1.Range("D9").Text & " from " & Sheet1.Range("D4").Value + 1 & " to " & Sheet1.Range("D6").Value & "."
    wApp.Selection.TypeText Text:=Q1
    wApp.Selection.TypeParagraph
    
    'Question 2
    'What are the major factors driving the XYZ Market growth?
    'Driver 1 and Driver 2 are the major factor driving the growth of XYZ market.
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    Q2 = "What are the major factors driving the " & Sheet1.Range("D2").Value & " growth?" & vbNewLine & Sheet1.Range("C15").Value & " and " & Sheet1.Range("C16").Value & " are the major factor driving the growth of " & Sheet1.Range("D2").Value & "."
    wApp.Selection.TypeText Text:=Q2
    wApp.Selection.TypeParagraph
    
    'Question 3
    'What are the key factors hampering growth of the XYZ Market?
    'Restrain  1 and Driver 2 are the major factor hampering the growth of XYZ market.
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    Q3 = "What are the key factors hampering growth of the " & Sheet1.Range("D2").Value & "?" & vbNewLine & Sheet1.Range("C17").Value & " and " & Sheet1.Range("C18").Value & " are the major factor hampering the growth of " & Sheet1.Range("D2").Value & "."
    wApp.Selection.TypeText Text:=Q3
    wApp.Selection.TypeParagraph
    
    
    'Question 4
    'Which is the leading segment 1 name here segment in the XYZ Market?
    'In terms of segment 1 name here, Dominating segment name here, estimated to dominate the market revenue share 2024.
    SubTitle = ""
    pnt = 0
    For i = 3 To Sheet1.Cells(Rows.Count, "I").End(xlUp).Row
       If Sheet1.Range("I" & i).Value > pnt Then
            SubTitle = Replace(Sheet1.Range("H" & i).Value, ">", "")
            pnt = Sheet1.Range("I" & i).Value
       End If
    Next i
    
    Segment = StrConv(Replace(LCase(Sheet1.Range("H2").Value), "by", ""), vbProperCase)
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    Q4 = "Which is the leading " & Trim(Segment) & " in the " & Sheet1.Range("D2").Value & "?" & vbNewLine & "In terms of " & Trim(Segment) & ", " & SubTitle & ", estimated to dominate the market revenue share " & Sheet1.Range("D4").Value + 1 & "."
    wApp.Selection.TypeText Text:=Q4
    wApp.Selection.TypeParagraph
    
    
    'Question 5
    'Which are the major players operating in the XYZ Market?
    'Nexans, Prysmian Group, General Cable, Sumitomo Electric Industries, Encore Wire, Finolex Cables, KEI Industries, Polycab Wires, APAR Industries, Sterlite Technologies are the major players.
    SubTitle = ""
    For i = 3 To Sheet1.Cells(Rows.Count, "G").End(xlUp).Row
        SubTitle = SubTitle & ", " & Sheet1.Range("G" & i).Value
    Next i
    
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    Q5 = "Which are the major players operating in the " & Sheet1.Range("D2").Value & "?" & vbNewLine & Mid(SubTitle, 3, Len(SubTitle) - 2) & " are the major players."
    wApp.Selection.TypeText Text:=Q5
    wApp.Selection.TypeParagraph
    
    'Question 6
    'Which region will lead the XYZ Market?
    'Asia Pacific is expected to lead the XYZ Market.
    SubTitle = ""
    For i = 23 To 28
        If Sheet1.Range("D" & i).Value = "Dominating" Then
            SubTitle = Sheet1.Range("C" & i).Value
        End If
    Next i
    
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    Q6 = "Which region will lead the " & Sheet1.Range("D2").Value & "?" & vbNewLine & SubTitle & " is expected to lead the " & Sheet1.Range("D2").Value & "."
    wApp.Selection.TypeText Text:=Q6
    wApp.Selection.TypeParagraph
      
End Function


'**********-Create Table of Content
Function AddSegmentation()
    Dim sCursor As Long, eCursor As Long, sCursor1 As Long, eCursor1 As Long
    
    'Header
    myStr = "Market Segmentation"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Color = wdColorBlue
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Size = 14
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    'Segmentation
    'Segments Covered
    sCursor1 = wApp.Selection.Start
    For K = 8 To Sheet1.Range("H2").End(xlToRight).Column
        If Not Sheet1.Cells(2, K).Value Like "% Market share*" And Sheet1.Cells(3, K).Value <> "" Then
            'Operating System Insights (Revenue, USD Billion, 2019 - 2031)
            
            TxtContent = Replace(StrConv(Sheet1.Cells(2, K).Value, vbProperCase), "By", "") & " Insights (Revenue, " & Sheet1.Range("D10").Value & ", " & Sheet1.Range("D5").Value & " - " & Sheet1.Range("D6").Value & ")"
            wApp.Selection.EndKey unit:=wdStory
            wApp.Selection.TypeText Text:=TxtContent
            wApp.Selection.TypeParagraph
            
            TxtContent = ""
            For r = 3 To Sheet1.Cells(Rows.Count, K).End(xlUp).Row
                il = Len(Sheet1.Cells(r, K).Value) - Len(Replace(Sheet1.Cells(r, K).Value, ">", ""))
                If il <= 4 Then
                    If r = 3 Then
                        TxtContent = Replace(Sheet1.Cells(r, K).Value, ">", "")
                    Else
                        TxtContent = TxtContent & vbNewLine & Replace(Sheet1.Cells(r, K).Value, ">", "")
                    End If
                    
                End If
            Next r
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=TxtContent
            eCursor = wApp.Selection.End
            Call AddFirstLabelBullet(sCursor, eCursor)
                        
        End If
    Next K
    
    eCursor1 = wApp.Selection.End
    'wDoc.Range(sCursor1, eCursor1).Paragraphs.SpaceAfter = 0
    
    'Regional Segment
    'Regional Insights (Revenue, USD Billion, 2019 - 2031)
        
    TxtContent = "Regional Insights (Revenue, " & Sheet1.Range("D10").Value & ", " & Sheet1.Range("D5").Value & " - " & Sheet1.Range("D6").Value & ")"
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph

    wApp.Selection.TypeText Text:=TxtContent
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
        
    For K = 1 To Sheet3.Range("A1").End(xlToRight).Column
        'Regional Insights (Revenue, USD Billion, 2019 - 2031)
        TxtContent = Sheet3.Cells(1, K).Value
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
        Call AddFirstLabelBullet(sCursor, eCursor)
        
        TxtContent = ""
        For r = 2 To Sheet3.Cells(Rows.Count, K).End(xlUp).Row
            If r = 2 Then
                TxtContent = Sheet3.Cells(r, K).Value
            Else
                TxtContent = TxtContent & vbNewLine & Sheet3.Cells(r, K).Value
            End If
            
        Next r
        
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
        Call AddSecondLabelBullet(sCursor, eCursor)
        
    Next K
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph

    'Key Player
    TxtContent = "Key Players Insights"
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph

    TxtContent = ""
    For r = 3 To Sheet1.Cells(Rows.Count, "G").End(xlUp).Row
        If r = 3 Then
            TxtContent = Sheet1.Cells(r, "G").Value
        Else
            TxtContent = TxtContent & vbNewLine & Sheet1.Cells(r, "G").Value
        End If
    Next r
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    Call AddFirstLabelBullet(sCursor, eCursor)
      
    ParaJustifyEnd = eCursor
    
End Function


'**********-Create Table of Content
Function CreateTableOfContent()

    Dim myStr As String, SectionNumber As Integer, sCursor As Long, eCursor As Long
    Dim bYear As Long, sYear As Long, eYear As Long
    Dim cType As String, MarketName As String, Country As String
    
    MarketName = Sheet1.Range("D2").Value
    bYear = Sheet1.Range("D4").Value
    sYear = Sheet1.Range("D5").Value
    eYear = Sheet1.Range("D6").Value
    cType = Sheet1.Range("D10").Value
    
    '&H25CB = Circle
    '&H25A0 = Filled Square
    '&H25A1 =
    '&H2022 =
    
    wDoc.Activate
    'wDoc.DefaultTabStop = 10
    wDoc.Paragraphs.Format.LineUnitAfter = 0
    'TOC Header
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    myStr = Sheet1.Range("D2").Value & " Report - Table of Contents"
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.TypeParagraph
    
    'Section 1 'RESEARCH OBJECTIVES AND ASSUMPTIONS
    sCursor = wApp.Selection.Start
    myStr = "1. RESEARCH OBJECTIVES AND ASSUMPTIONS"
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.TypeParagraph
    
    sCursor = wApp.Selection.Start
    myStr = "Research Objectives" & vbNewLine & "Assumptions" & vbNewLine & "Abbreviations"
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    'Section 2 'MARKET PURVIEW
    myStr = "2. MARKET PURVIEW"
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.TypeParagraph
    
    sCursor = wApp.Selection.Start
    myStr = "Report Description"
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    sCursor = wApp.Selection.Start
    myStr = "Market Definition and Scope"
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddSecondLabelBullet(sCursor, eCursor)
    
    sCursor = wApp.Selection.Start
    myStr = "Executive Summary"
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    'Add Dynamic Segmentation
    myStr = ""
    For K = 8 To Sheet1.Range("H2").End(xlToRight).Column
        If Not Sheet1.Cells(2, K).Value Like "% Market share*" Then
            If myStr = "" Then
                myStr = Sheet1.Range("D2").Value & ", " & Sheet1.Cells(2, K).Value
            Else
                myStr = myStr & vbNewLine & Sheet1.Range("D2").Value & ", " & Sheet1.Cells(2, K).Value
            End If
        End If
    Next K
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    
    Call AddSecondLabelBullet(sCursor, eCursor)
    
    '''''
    SectionNumber = 3
    'Section 3 Customized Section
    For K = 2 To Sheet1.Range("G52").End(xlToLeft).Column
        If Sheet1.Cells(53, K).Value <> "" Then
            myStr = SectionNumber & ". " & Sheet1.Cells(52, K).Value
            wApp.Selection.TypeParagraph
            
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=myStr
            eCursor = wApp.Selection.End
            wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
            wApp.Selection.TypeParagraph
                    
            For r = 53 To Sheet1.Cells(Rows.Count, K).End(xlUp).Row
                il = Len(Sheet1.Cells(r, K).Value) - Len(Replace(Sheet1.Cells(r, K).Value, ">", ""))
                If il <= 3 Then
                    TxtContent = Replace(Sheet1.Cells(r, K).Value, ">", "")
                    wApp.Selection.EndKey unit:=wdStory
                    sCursor = wApp.Selection.Start
                    If il = 1 Then
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        Call AddFirstLabelBullet(sCursor, eCursor)
                    ElseIf il = 2 Then
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        Call AddSecondLabelBullet(sCursor, eCursor)
                        
                    ElseIf il = 3 Then
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        Call AddThirdLabelBullet(sCursor, eCursor)
                        
                    End If
                    
                End If
            Next r
            SectionNumber = SectionNumber + 1
        End If
    Next K
    
    
    '''''
    'Country = Sheet1.Range("D22").Value
    'Segmentation Section
    'Add Dynamic Segmentation content
    For K = 8 To Sheet1.Range("H2").End(xlToRight).Column
        If Not Sheet1.Cells(2, K).Value Like "% Market share*" And Sheet1.Cells(3, K).Value <> "" Then
            
            'Segment Header
            TxtContent = SectionNumber & ". " & Country & " " & MarketName & ", " & Sheet1.Cells(2, K).Value & ", " & bYear + 1 & "-" & eYear & ", " & "(" & cType & ")"
            
            wApp.Selection.EndKey unit:=wdStory
            wApp.Selection.TypeParagraph
            
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=TxtContent
            eCursor = wApp.Selection.End
            wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
            'Introduction
            TxtContent = "Introduction"
            wApp.Selection.EndKey unit:=wdStory
            wApp.Selection.TypeParagraph
            
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=TxtContent
            eCursor = wApp.Selection.End
            
            Call AddFirstLabelBullet(sCursor, eCursor)
                        
            TxtContent = "Market Share Analysis, " & bYear + 1 & " and " & eYear & " (%)"
            TxtContent = TxtContent & vbNewLine & "Y-o-Y Growth Analysis, " & sYear & " - " & eYear
            TxtContent = TxtContent & vbNewLine & "Segment Trends"
            
            sCursor = wApp.Selection.Start
            wApp.Selection.TypeText Text:=TxtContent
            eCursor = wApp.Selection.End
            
            Call AddSecondLabelBullet(sCursor, eCursor)
            
            
            For r = 3 To Sheet1.Cells(Rows.Count, K).End(xlUp).Row
                il = Len(Sheet1.Cells(r, K).Value) - Len(Replace(Sheet1.Cells(r, K).Value, ">", ""))
                If il <= 3 Then
                    TxtContent = Replace(Sheet1.Cells(r, K).Value, ">", "")
                    wApp.Selection.EndKey unit:=wdStory
                    
                    If il = 1 Then
                        sCursor = wApp.Selection.Start
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        
                        Call AddFirstLabelBullet(sCursor, eCursor)
                        
                        'Introduction & Market Size
                        TxtContent = "Introduction"
                        TxtContent = TxtContent & vbNewLine & "Market Size and Forecast, and Y-o-Y Growth, " & sYear & "-" & eYear & ", " & "(" & cType & ")"
                        
                        sCursor = wApp.Selection.Start
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        
                        Call AddSecondLabelBullet(sCursor, eCursor)
                        
                        
                    ElseIf il = 2 Then
                        sCursor = wApp.Selection.Start
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        
                        Call AddThirdLabelBullet(sCursor, eCursor)
                        
                    ElseIf il = 3 Then
                        sCursor = wApp.Selection.Start
                        wApp.Selection.TypeText Text:=TxtContent
                        eCursor = wApp.Selection.End
                        
                        Call AddThirdLabelBullet(sCursor, eCursor)
                        
                    End If
                    
                End If
            Next r
            SectionNumber = SectionNumber + 1
        End If
    Next K
    
    'Start of Regional Section
    'If Country = "Global" Then
    'Add Regional Section
    TxtContent = SectionNumber & ".  Global " & MarketName & ", By Region, " & sYear & " - " & eYear & ", Value (" & cType & ")"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    SectionNumber = SectionNumber + 1
    
    'Introduction
    TxtContent = "Introduction"
    TxtContent = TxtContent & vbNewLine & "Market Share (%) Analysis, " & bYear + 1 & "," & bYear + 4 & " & " & eYear & ", Value (" & cType & ")"
    TxtContent = TxtContent & vbNewLine & "Market Y-o-Y Growth Analysis (%), " & sYear & " - " & eYear & ", Value (" & cType & ")"
    TxtContent = TxtContent & vbNewLine & "Regional Trends"
                    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
            
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    For K = 1 To Sheet3.Range("A1").End(xlToRight).Column
        TxtContent = Sheet3.Cells(1, K).Value
        
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
            
        Call AddSecondLabelBullet(sCursor, eCursor)
        
        TxtContent = "Introduction"
        For s = 8 To Sheet1.Range("G2").End(xlToRight).Column
            If Not Sheet1.Cells(2, s).Value Like "*Market share*" Then
                TxtContent = TxtContent & vbNewLine & "Market Size and Forecast, " & Sheet1.Cells(2, s).Value & " , " & sYear & " - " & eYear & ", Value (" & cType & ")"
            End If
        Next s
                    
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
            
        Call AddThirdLabelBullet(sCursor, eCursor)
        
        TxtContent = Sheet3.Cells(2, K).Value
        For c = 3 To Sheet3.Cells(1, K).End(xlDown).Row
            TxtContent = TxtContent & vbNewLine & Sheet3.Cells(c, K).Value
            
        Next c
        
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
            
        Call AddFourthLabelBullet(sCursor, eCursor)
        
        wApp.Selection.EndKey unit:=wdStory
        wApp.Selection.TypeParagraph
    
        
    Next K
            
    'End If
    'End of Regional Section
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
        
    'COMPETITIVE LANDSCAPE
    TxtContent = SectionNumber & ". COMPETITIVE LANDSCAPE"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    SectionNumber = SectionNumber + 1
    
    For r = 3 To Sheet1.Cells(Rows.Count, "G").End(xlUp).Row
        TxtContent = Sheet1.Range("G" & r).Value
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
                        
        Call AddFirstLabelBullet(sCursor, eCursor)
                            
        'Company Highlights, 'Product Portfolio, Key Developments, Financial Performance, Strategies
        TxtContent = "Company Highlights" & vbNewLine & "Product Portfolio" & vbNewLine & "Key Developments" & vbNewLine & "Financial Performance" & vbNewLine & "Strategies"
        
        sCursor = wApp.Selection.Start
        wApp.Selection.TypeText Text:=TxtContent
        eCursor = wApp.Selection.End
                        
        Call AddSecondLabelBullet(sCursor, eCursor)
        
    Next r
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    
    'Analyst Recomendation Section
    TxtContent = SectionNumber & ".  Analyst Recommendations"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    SectionNumber = SectionNumber + 1
    
    'Introduction
    TxtContent = "Wheel of Fortune"
    TxtContent = TxtContent & vbNewLine & "Analyst View"
    TxtContent = TxtContent & vbNewLine & "Coherent Opportunity Map"
                    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
            
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    'Research Methodology
    TxtContent = SectionNumber & ". References and Research Methodology"
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    
    
    TxtContent = "References" & vbNewLine & "Research Methodology" & vbNewLine & "About us"
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    
    Call AddFirstLabelBullet(sCursor, eCursor)
    
    
    TxtContent = "*Browse 32 market data tables and 28 figures on '" & MarketName & "' - Global forecast to " & eYear
        
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=TxtContent
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
        
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
        
    'For Each para In wDoc.Paragraphs
        'para.Format.SpaceAfter = 0
    'Next para
    
End Function



Function AddKeyPlayerPR()
    
    wDoc.Activate
    
    myStr = "Key Player:"
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
    With wDoc.Range(Start:=sCursor, End:=eCursor)
        .Font.Bold = wdToggle
        .Font.Size = 14
        .Font.ColorIndex = wdDarkBlue
    End With
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
     
    myStr = Sheet5.Range("B8").Value
    sCursor = wApp.Selection.Start
    wApp.Selection.TypeText Text:=myStr
    eCursor = wApp.Selection.End
    'wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = wdToggle
     
End Function


Function AddCompetitiveLandscape()
    
    'Get Start postion of "Market Size and Trends"
    mDoc.Activate
    Set rngCopy = mDoc.Content
    With rngCopy.Find
        .Text = "Competitive Landscape:" 'Sheet1.Range("C15").Value
        .Forward = True
        .Wrap = 1
    End With
    
    sMDocCursor = ""
    If rngCopy.Find.Execute Then
        sMDocCursor = rngCopy.Start + Len("Competitive Landscape:")
    End If
        
    
    'Get end postion of "Market Size and Trends"
    wApp.Selection.EndKey unit:=wdStory
    eMDocCursor = wApp.Selection.End - Len("Competitive Landscape:")
    
    Set rngCopy = mDoc.Range(Start:=sMDocCursor, End:=eMDocCursor)
    
    wDoc.Activate
    sCursor = wApp.Selection.Start
    If (sMDocCursor <> "" Or sMDocCursor = 0) Or (eMDocCursor <> "" Or eMDocCursor = 0) Then
        wApp.Selection.TypeText Text:=rngCopy.Text
    Else
        wApp.Selection.TypeText Text:="Unable to search the text. please do it manually."
    End If
    eCursor = wApp.Selection.End
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Bold = False
    wDoc.Range(Start:=sCursor, End:=eCursor).Font.Italic = False
    
    wDoc.Range(Start:=1, End:=eCursor).ParagraphFormat.Alignment = wdAlignParagraphJustify
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    

End Function
