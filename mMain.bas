Attribute VB_Name = "mMain"
Option Base 1

'Power Point Variable
Dim pptApp As PowerPoint.Application
Dim pptTem As PowerPoint.Presentation
Dim pptSlide As PowerPoint.slide

'Doc Variable
Public wApp As Word.Application
Public wDoc As Word.Document, mDoc As Word.Document
Public FolderPath As String

Public ParaJustifyStart As Long, ParaJustifyEnd As Long


Sub CreateImageAndTable()
    
    Dim imgPath As String
    Dim FolderPath As String, FileName As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
     
    'Set Output folder
    FolderPath = ThisWorkbook.Path & "\Output\" & Format(Date, "YYYY") & "\" & Format(Date, "MMM-YY") & "\" & Format(Date, "DD") & "\" & Sheet1.Range("D2").Value
     
    'Open Template
    Set pptApp = New PowerPoint.Application
    pptApp.Visible = msoTrue
    
    Set pptTem = pptApp.Presentations.Open(ThisWorkbook.Path & "\Template\Collateral Designs.pptx")
    pptTem.SaveAs ThisWorkbook.Path & "\Template\Output.pptx"
    
    'Minimise Power Point Window
    For i = 1 To pptApp.Windows.Count
        pptApp.Windows.Item(i).WindowState = ppWindowMinimized
    Next i
    
    '******************************************
    '-Regional Insights
    Call ChangeRegionalInsights
    
    '******************************************
    '-Impact Analysis of Key Factors
    Call ChangeImpactAnalysisOfKeyFactors
    
    '******************************************
    '-Key Takeaways of Analyst:
    Call ChangeKeyTakeAwaysOfAnalyst
    
    '******************************************
    '-Segmental Insights
    Call ChangeSegmentalInsights
    
    '******************************************
    '-Market Key Player Concentration
    Call ChangeMarketKeyPlayer
    pptTem.Save
    
    '******************************************
    '-Create Image of all side
    Call CreateImageFromSlide
    
    '******************************************
    'Closing Template
    pptTem.Save
    
    pptTem.SaveAs FolderPath & "\Output.pptx"
    
    pptTem.Close
    pptApp.Quit
    
    Set pptTem = Nothing
    Set pptApp = Nothing
    
    
    'Create Docemnt
    Set wApp = New Word.Application
    wApp.Visible = True
    
    'Minimise Word Window
    For i = 1 To wApp.Windows.Count
        wApp.Windows.Item(i).WindowState = wdWindowStateMinimize
    Next i
    
    '******************************************
    Set wDoc = wApp.Documents.Add
    
    'Add Title
    Call AddDocTitle
    
    'Market Size and Trend:
    filepath = FolderPath & "\" & "RS - " & Sheet1.Range("D2").Value & ".doc"
    Set mDoc = wApp.Documents.Open(filepath)
    
    Call AddMarketSizeAndTrends
    
    'Segmentation
    Call AddSegmentAnalysis
    
    'Regional Insight
    Call AddRegionalAnalysis
    
    'Regional Insight
    Call AddCompetitiveLandscape
    
    'Key Development
    Call AddKeyDevelopment
    
    'Key Takeaway:
    'Call AddKeyTakeAway
    
    'Defination
    Call AddKeyTakeAways
    
    'Market Opportunity:
    'Call AddMarketChallenge
    
    'Defination
    'Call AddDefination
    
    '-Create Table in word
    Call CreateDocTable
    
    'Market Driver: 1
    Call AddMarketDriver
    
    'Market Challenge:
    Call AddMarketChallenge
    
    mDoc.Close SaveChanges:=False
    
    '-Create Questionaries in word
    Call CreateDocQuestionaries
    
    'Segmentation
    Call AddSegmentation
    
    wDoc.Range(Start:=ParaJustifyStart, End:=ParaJustifyEnd).ParagraphFormat.Alignment = wdAlignParagraphJustify
    
    FileName = "Combined - " & Sheet1.Range("D2").Value & ".docx"
    wDoc.SaveAs2 FileName:=FolderPath & "\" & FileName
    wDoc.Close
    
    '******************************************
    '-Create TOC document
    Set wDoc = wApp.Documents.Add
    
    'Create TOC
    Call CreateTableOfContent
    
    FileName = "TOC - " & Sheet1.Range("D2").Value & ".docx"
    wDoc.SaveAs2 FileName:=FolderPath & "\" & FileName
    wDoc.Close
    
    '******************************************
    '-Create PR document
    If Sheet1.Range("D1").Value = "Updated Collateral" Then
        FileName = "PR - " & Sheet1.Range("D2").Value & ".doc"
        prDocPath = FolderPath & "\" & FileName
        Set wDoc = wApp.Documents.Open(prDocPath)
        
        wApp.Selection.EndKey unit:=wdStory
        wApp.Selection.TypeParagraph
        
        'Create TOC
        Call AddKeyDevelopment
    
        'Add Key Player
        Call AddKeyPlayerPR
        
        wApp.Selection.WholeStory
        wApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
        
        wDoc.Save
        wDoc.Close
    End If
    
    '******************************************
    '-Create RD document
    FileName = "RD - " & Sheet1.Range("D2").Value & ".doc"
    prDocPath = FolderPath & "\" & FileName
    Set wDoc = wApp.Documents.Open(prDocPath)
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph

    'Create TOC
    Call AddSegmentation
     
    wApp.Selection.WholeStory
    wApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
         
    wDoc.Save
    wDoc.Close
    
    '******************************************
    'Closing the Application
    wApp.Quit
    
    Set wApp = Nothing
    Set wDoc = Nothing
    
    
    On Error Resume Next
    Kill ThisWorkbook.Path & "\Template\Output.pptx"
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'MsgBox "Document Created Sucessfully !!"
    
End Sub





'**********-Regional Insights

Function ChangeRegionalInsights()
    Dim Title As String, pValue As String, DominatingRagion As String, FastestGrowingRegion As String
    
    
    'Find Dominating Region and Fast Growing Region
    
    For i = 23 To 28
        If Sheet1.Range("D" & i).Value = "Dominating" Then
            DominatingRagion = Sheet1.Range("C" & i).Value
        End If
        
        If Sheet1.Range("D" & i).Value = "Fastest Growing" Then
            FastestGrowingRegion = Sheet1.Range("C" & i).Value
        End If
    Next i
    
    'Color Dominating Region
    '1 East Africa,2 Latin America, 3 Africa, 4 North America, 5 Europe, 6 Asia
    'pptTem.Slides(1).Shapes(7).Select
    Select Case DominatingRagion
        Case Is = "North America"
            pptTem.Slides(1).Shapes(4).Fill.ForeColor.RGB = RGB(23, 52, 97)
            
        Case Is = "Europe"
            pptTem.Slides(1).Shapes(5).Fill.ForeColor.RGB = RGB(23, 52, 97)
            
        Case Is = "Asia Pacific"
            pptTem.Slides(1).Shapes(6).Fill.ForeColor.RGB = RGB(23, 52, 97)
            
        Case Is = "Latin America"
            pptTem.Slides(1).Shapes(2).Fill.ForeColor.RGB = RGB(23, 52, 97)
            
        Case Is = "Middle East"
            pptTem.Slides(1).Shapes(1).Fill.ForeColor.RGB = RGB(23, 52, 97)
            
        Case Is = "Africa"
            pptTem.Slides(1).Shapes(3).Fill.ForeColor.RGB = RGB(23, 52, 97)
        
    End Select
    
        
    'Color Fast Growning Region
    ''1 East Africa,2 Latin America, 3 Africa, 4 North America, 5 Europe, 6 Asia
    Select Case FastestGrowingRegion
        Case Is = "North America"
            pptTem.Slides(1).Shapes(4).Fill.ForeColor.RGB = RGB(117, 200, 146)
            
        Case Is = "Europe"
            pptTem.Slides(1).Shapes(5).Fill.ForeColor.RGB = RGB(117, 200, 146)
            
        Case Is = "Asia Pacific"
            pptTem.Slides(1).Shapes(6).Fill.ForeColor.RGB = RGB(117, 200, 146)
            
        Case Is = "Latin America"
            pptTem.Slides(1).Shapes(2).Fill.ForeColor.RGB = RGB(117, 200, 146)
            
        Case Is = "Middle East"
            pptTem.Slides(1).Shapes(1).Fill.ForeColor.RGB = RGB(117, 200, 146)
            
        Case Is = "Africa"
            pptTem.Slides(1).Shapes(3).Fill.ForeColor.RGB = RGB(117, 200, 146)
        
    End Select
        
    
    
    'Change Title
    '16 Total Market Size
    
    '9 Slide Title
    Title = "Regional Insights, " & Format(Date, "YYYY")
    With pptTem.Slides(1).Shapes(9)
        .TextFrame.TextRange.Text = Title
        
    End With
    
    '13 Market Name
    Title = Sheet1.Range("D2").Value
    With pptTem.Slides(1).Shapes(13)
        .TextFrame.TextRange.Text = Title
        
    End With
    
    
    'Change Market Share %
    '12 Percentage, 11 Percentage Comment
    pValue = Sheet1.Range("D22").Text
    With pptTem.Slides(1).Shapes(12)
        .TextFrame.TextRange.Text = pValue
        
    End With
    
    Title = DominatingRagion & " - Estimated Market Revenue Share, " & Sheet1.Range("D4").Value + 1
    With pptTem.Slides(1).Shapes(11)
        .TextFrame.TextRange.Text = Title
        '.TextFrame.TextRange.Characters(1, 4).Select
        With .TextFrame.TextRange.Characters(Start:=1, Length:=Len(DominatingRagion) + 1)
                '.Select
                .Font.Bold = msoCTrue
                
        End With
        With .TextFrame.TextRange.Characters(Start:=Len(DominatingRagion) + 2, Length:=Len(Title))
                '.Select
                .Font.Bold = msoFalse
                
        End With
    End With
    
    
    
    '16 market size,
    
    Title = "Total Market Size: " & Sheet1.Range("D7").Value
    With pptTem.Slides(1).Shapes(16)
        .TextFrame.TextRange.Text = Title
        With .TextFrame.TextRange.Characters(Start:=19, Length:=Len(Sheet1.Range("D7").Value) + 19)
                .Font.Bold = msoCTrue
                .Font.Size = 28
        End With
    End With
    
    
End Function




'**********-Impact Analysis Of Key Factors

Function ChangeImpactAnalysisOfKeyFactors()
    Dim DriverIndicator1 As Integer, DriverIndicator2 As Integer
    Dim RestraintIndicator1 As Integer, RestraintIndicator2 As Integer
    Dim OpprtunityIndicator1 As Integer, OpprtunityIndicator2 As Integer
    
    Dim Driver1 As String, Driver2 As String
    Dim Restraint1 As String, Restraint2 As String
    Dim Opprtunity1 As String, Opprtunity2 As String
    
    
    'Drivers
    '3 Driver1, 4 Driver2
    pptTem.Slides(2).Select
    
    Driver1 = Sheet1.Range("C15").Value
    With pptTem.Slides(2).Shapes(3)
        .TextFrame.TextRange.Text = Driver1
        
    End With
    
    Driver2 = Sheet1.Range("C16").Value
    With pptTem.Slides(2).Shapes(4)
        .TextFrame.TextRange.Text = Driver2
        
    End With
    
    'Driver Indicator
    '14 Driver1, 16 Driver2
    DriverIndicator1 = Sheet1.Range("D15").Value
    With pptTem.Slides(2).Shapes(14)
        '.Select
        .Left = 340 + (100 * DriverIndicator1)
        .Top = 109
    End With
    
    DriverIndicator2 = Sheet1.Range("D16").Value
    With pptTem.Slides(2).Shapes(16)
        '.Select
        .Left = 340 + (100 * DriverIndicator2)
        .Top = 167
    End With
    
    'Restraint
    '5 Restraint1, 6 Restraint2
    Restraint1 = Sheet1.Range("C17").Value
    With pptTem.Slides(2).Shapes(5)
        .TextFrame.TextRange.Text = Restraint1
        
    End With
    
    Restraint2 = Sheet1.Range("C18").Value
    With pptTem.Slides(2).Shapes(6)
        .TextFrame.TextRange.Text = Restraint2
        
    End With
    
    '13 ResPointer1,  15 ResPointer2
    RestraintIndicator1 = Sheet1.Range("D17").Value
    With pptTem.Slides(2).Shapes(13)
        '.Select
        .Left = 340 + (100 * RestraintIndicator1)
        .Top = 228
    
    End With
    
    RestraintIndicator2 = Sheet1.Range("D18").Value
    With pptTem.Slides(2).Shapes(15)
        '.Select
        .Left = 340 + (100 * RestraintIndicator2)
        .Top = 284
    
    End With
    
    'Opportunity
    '7 Oppertunity1, 8 Oppertunity1
    Opprtunity1 = Sheet1.Range("C19").Value
    With pptTem.Slides(2).Shapes(7)
        .TextFrame.TextRange.Text = Opprtunity1
        
    End With
    
    Opprtunity2 = Sheet1.Range("C20").Value
    With pptTem.Slides(2).Shapes(8)
        .TextFrame.TextRange.Text = Opprtunity2
        
    End With
    
    '17 OpporPointer1, 18 OpporPointer1
    OpprtunityIndicator1 = Sheet1.Range("D19").Value
    With pptTem.Slides(2).Shapes(17)
        '.Select
        .Left = 340 + (100 * OpprtunityIndicator1)
        .Top = 345
    
    End With
    
    OpprtunityIndicator2 = Sheet1.Range("D20").Value
    With pptTem.Slides(2).Shapes(18)
        '.Select
        .Left = 340 + (100 * OpprtunityIndicator2)
        .Top = 407
    
    End With
    
    
    'Title
    ''12 Slide Title,
    
    Title = "Impact Analysis of Key Factors" & vbNewLine & Sheet1.Range("D2").Value
    
    With pptTem.Slides(2).Shapes(12).TextFrame.TextRange
        .Font.Bold = msoFalse
        .Text = Title
        With .Characters(Start:=1, Length:=Len("Impact Analysis of Key Factors") + 1)
                '.Select
                .Font.Bold = msoTrue
                
        End With
    End With
    
    
End Function



'**********-Key Take Aways Of Analyst


Function ChangeKeyTakeAwaysOfAnalyst()

    Dim myStrArr() As String
    Dim TotalLine As Double
    
    pptTem.Slides(3).Select
    
    TotalLine = 0
    TotalLength = 0
    For i = 40 To 44
        If Sheet1.Range("D" & i).Value <> "" Then
            TotalLine = TotalLine + Round((Len(Sheet1.Range("D" & i).Value) / 111), 1) + 1
            TotalLength = TotalLength + Len(Sheet1.Range("D" & i).Value)
            ReDim Preserve myStrArr(i - 39)
            myStrArr(i - 39) = Sheet1.Range("D" & i).Value
        End If
    Next i
    
    '1 Textbox
    
    With pptTem.Slides(3).Shapes(1).TextFrame.TextRange
        .Text = ""
        For i = LBound(myStrArr) To UBound(myStrArr)
            StartPosition = VBA.Len(.Text) + 1
            '.Characters(Start:=StartPosition).Text = myStrArr(i) & vbNewLine
            With .Characters(Start:=StartPosition, Length:=Len(myStrArr(i)) + StartPosition)
                '.Select
                '.ParagraphFormat.Bullet.Character = 11162
                .Characters(Start:=StartPosition).Text = myStrArr(i) & vbNewLine & vbNewLine
                
            End With
         Next i
         
    End With
    
    TotalLine = Round(TotalLine - 11.45, 2)
    'pptTem.Slides(3).Shapes(1).Select
    pptTem.Slides(3).Shapes(1).TextFrame2.TextRange.Font.Size = 22 - (0.66 * TotalLine)
    
    
    Erase myStrArr
    
    
    
End Function




'**********-Segmental Insights


Function ChangeSegmentalInsights()
    
    Dim pValue As String, SubTitle As String, Title As String, SegmentName As String
    Dim mLabel As String
    
    'Prepare Table for Chart
    Sheet4.Range("H1:I" & Sheet4.Cells(Rows.Count, "H").End(xlUp).Row).ClearContents
    Sheet4.Range("H1").Value = Sheet1.Range("H2").Value
    Sheet4.Range("I1").Value = Sheet1.Range("I2").Value
    
    For r = 3 To Sheet1.Cells(Rows.Count, "H").End(xlUp).Row
        il = Len(Sheet1.Cells(r, "H").Value) - Len(Replace(Sheet1.Cells(r, "H").Value, ">", ""))
        If il <= 1 Then
            x = Sheet4.Cells(Rows.Count, "H").End(xlUp).Row + 1
            Sheet4.Range("H" & x).Value = Replace(Sheet1.Cells(r, "H").Value, ">", "")
            Sheet4.Range("I" & x).Value = Sheet1.Cells(r, "I").Text
            
        End If
    Next r
    
    
    '3 Chart
    With pptTem.Slides(4).Shapes.AddChart2(251, Type:=xlDoughnut, Left:=120, Top:=110, Width:=600, Height:=400).Chart
        Application.Wait (Now + TimeValue("0:00:2"))
        
        'Changing the source data
        '.ChartData.Workbook.Worksheets(1).Activate
        .ChartData.Workbook.Worksheets(1).UsedRange.ClearContents
        Application.Wait (Now + TimeValue("0:00:1"))
        
        Sheet4.Range("H1:I" & Sheet4.Cells(Rows.Count, "H").End(xlUp).Row).Copy
        Err.Clear
        On Error Resume Next
        .ChartData.Workbook.Worksheets(1).Cells(1, 1).PasteSpecial ppPasteText
        If Err.Description <> "" Then
            Sheet4.Range("H1:I" & Sheet4.Cells(Rows.Count, "H").End(xlUp).Row).Copy
            .ChartData.Workbook.Worksheets(1).Cells(1, 1).PasteSpecial ppPasteText
        End If
        On Error GoTo 0
        
        Application.Wait (Now + TimeValue("0:00:1"))
        .SetSourceData Source:="='Sheet1'!" & .ChartData.Workbook.Worksheets(1).Cells(1, 1).CurrentRegion.Address, PlotBy:=xlColumns
        Application.Wait (Now + TimeValue("0:00:1"))
        .FullSeriesCollection(1).XValues = "=Sheet1!$A$2:$A$" & .ChartData.Workbook.Worksheets(1).Cells(.ChartData.Workbook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
        
        .ChartData.Workbook.RefreshAll
        Application.Wait (Now + TimeValue("0:00:1"))
        
        .ChartData.Workbook.Close
        Application.Wait (Now + TimeValue("0:00:1"))
        
        'Chart Titel
        '.HasTitle = True
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 20
        '.SetElement msoElementChartTitleCenteredOverlay
        
        .FullSeriesCollection(1).Select
        .ChartGroups(1).DoughnutHoleSize = 65
        With .FullSeriesCollection(1)
            .ApplyDataLabels
            .Select
            '.DataLabels.Position = xlLabelPositionAbove
            .DataLabels.ShowValue = False
            .DataLabels.ShowPercentage = True
            .DataLabels.NumberFormat = "0.0%"
            .DataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .DataLabels.Format.TextFrame2.TextRange.Font.Bold = msoTrue
            .DataLabels.Format.TextFrame2.TextRange.Font.Size = 14
            .DataLabels.Format.TextFrame2.VerticalAnchor = msoAnchorTop
            .DataLabels.Format.TextFrame2.HorizontalAnchor = msoAnchorNone
            .HasLeaderLines = True
        End With
        
        '''''''''
        'Make Data point as XX for Sample
        DataPoint = 0
        For i = 3 To Sheet1.Cells(Rows.Count, "I").End(xlUp).Row
            If Sheet1.Range("I" & i).Value > DataPoint Then
                mLabel = .FullSeriesCollection(1).Points(i - 2).DataLabel.Text
                DataPoint = Sheet1.Range("I" & i).Value
            End If
        Next i
        
        'Make Data point as XX for Sample
        For i = 1 To .FullSeriesCollection(1).Points.Count
            If .FullSeriesCollection(1).Points(i).DataLabel.Text = mLabel Then
                'MsgBox "test"
            Else
                .FullSeriesCollection(1).Points(i).DataLabel.Formula = "xx.x%"
            End If
        Next i
        
        '''''''
        
        .HasLegend = True
        .Legend.Format.TextFrame2.TextRange.Font.Size = 14
        
        
    End With
    
    
    
    
    'Percentage :8 Percentage
    'pptTem.Slides(4).Shapes(7).Select
    With pptTem.Slides(4).Shapes(8)
        .TextFrame.TextRange.Text = mLabel
        
    End With
    
    '7 Percentage Comment,
    pnt = 0
    For i = 2 To Sheet4.Cells(Rows.Count, "H").End(xlUp).Row
       If Sheet4.Range("I" & i).Value > pnt Then
            SubTitle = Sheet4.Range("H" & i).Value
            pnt = Sheet4.Range("I" & i).Value
       End If
    Next i
    'pptTem.Slides(4).Shapes(6).Select
    Title = SubTitle & " " & Trim(Replace(Sheet1.Range("H2").Value, "By", "")) & " - Estimated Market Revenue Share, " & Sheet1.Range("D4").Value + 1
    With pptTem.Slides(4).Shapes(7)
        .TextFrame.TextRange.Text = Title
        .TextFrame.TextRange.Font.Bold = msoFalse
         With .TextFrame.TextRange.Characters(Start:=1, Length:=Len(SubTitle) + 1)
            .Font.Bold = msoCTrue
            '.ParagraphFormat.Bullet.Character = 11162
            
        End With

    End With
    
    '9 Market Name
    Title = Sheet1.Range("D2").Value
    'pptTem.Slides(4).Shapes(8).Select
    With pptTem.Slides(4).Shapes(9)
        .TextFrame.TextRange.Text = Title
        
    End With
    
    
    
    'Slide Title: 1 Slide Title
    SubTitle = Sheet1.Range("D2").Value & ", " & Sheet1.Range("H2").Value & ", " & Sheet1.Range("D4").Value + 1
    
    With pptTem.Slides(4).Shapes(1)
        .TextFrame.TextRange.Text = SubTitle
        
    End With
    
    '3 Market Size
    
    Title = "Total Market Size: " & Sheet1.Range("D7").Value
    With pptTem.Slides(4).Shapes(3)
        .TextFrame.TextRange.Text = Title
        .TextFrame.TextRange.Font.Bold = msoFalse
         With .TextFrame.TextRange.Characters(Start:=19, Length:=Len(Title) + 19)
            '.Select
            .Font.Bold = msoCTrue
            .Font.Size = 28
            '.ParagraphFormat.Bullet.Character = 11162
            
        End With
    End With
    
    
    
End Function



'**********-Market Key Player


Function ChangeMarketKeyPlayer()
    
    Dim MajorPlayer As String, mConPoint As Integer, Title As String, SubTitle As String
    
    pptTem.Slides(5).Select
    
    
    'Move Arrow: 8 Arrow, 9 Arrow Comment
    mConPoint = Sheet1.Range("D31").Value
    Title = Sheet1.Range("D2").Value
    With pptTem.Slides(5).Shapes(8)
        'Bottom = 380, Top = 170
        .Top = 380 - (42 * mConPoint)
        
    End With
    With pptTem.Slides(5).Shapes(9)
        'Bottom = 380, Top = 170
        .TextFrame.TextRange.Text = Title
        .Top = 380 - (42 * mConPoint)
        
    End With

    '12 Leading Player Right,13 Leading Player Left
    NoOfPlayer = Sheet1.Cells(Rows.Count, "G").End(xlUp).Row - 2
    NoOfPlayerLine = Round((NoOfPlayer / 2) + 0.1, 0)
    
    Constant = 100
    For i = 1 To NoOfPlayerLine
        With pptTem.Slides(5).Shapes.AddShape(msoShapeRectangle, Left:=470, Top:=Constant, Width:=230, Height:=35)
            
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            .Fill.Transparency = 1
            .TextFrame.TextRange.Font.Color = vbBlack
            .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
            .TextFrame.VerticalAnchor = msoAnchorTop
            .TextFrame.TextRange.Font.Size = 15
            .TextFrame.TextRange.Text = Sheet1.Range("G" & (i * 2) + 1).Value
            .TextFrame.TextRange.ParagraphFormat.Bullet.Character = 11162
            
        End With
        
        With pptTem.Slides(5).Shapes.AddShape(msoShapeRectangle, Left:=710, Top:=Constant, Width:=230, Height:=35)
            
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            .Fill.Transparency = 1
            .TextFrame.TextRange.Font.Color = vbBlack
            .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
            .TextFrame.VerticalAnchor = msoAnchorTop
            .TextFrame.TextRange.Font.Size = 15
            .TextFrame.TextRange.Text = Sheet1.Range("G" & (i * 2) + 2).Value
            .TextFrame.TextRange.ParagraphFormat.Bullet.Character = 11162
            
        End With
        
        
        'With pptTem.Slides(5).Shapes.AddShape(msoShapeRectangle, Left:=470, Top:=Constant + 35, Width:=460, Height:=0)
            
            '.Line.Visible = msoTrue
            '.Fill.Visible = msoFalse '.ForeColor.RGB = RGB(219, 219, 217)
            '.Line.DashStyle = msoLineRoundDot
            
        'End With
        
        Constant = Constant + 40
        
    Next i
    
    'Adjust the box left and right
    '1 Leading Player Right,19 Leading Player Left
    With pptTem.Slides(5).Shapes(12)
        .Height = 20 + (NoOfPlayerLine * 40)
    End With
    With pptTem.Slides(5).Shapes(13)
        .Height = 20 + (NoOfPlayerLine * 40)
    End With
    
End Function


'**********-Image From Slide


Function CreateImageFromSlide()
    
    'Create Folder if not exits
    'Output Folder
    'FolderPath = ThisWorkbook.Path & "/Output" & Format(Date, "YYYY")
    FolderPath = ThisWorkbook.Path & "/Output"
    
    ' Check if the folder exists
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    'Year Folder
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "YYYY")
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    'Month Folder
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "YYYY") & "/" & Format(Date, "MMM-YY")
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    'Day Folder
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "YYYY") & "/" & Format(Date, "MMM-YY") & "/" & Format(Date, "DD")
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    'Markte Folder
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "YYYY") & "/" & Format(Date, "MMM-YY") & "/" & Format(Date, "DD") & "/" & Sheet1.Range("D2").Value
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    '-Generage Image
    'Delete previous files
    FileName = Dir(FolderPath & "*")
    Do While FileName <> ""
        ' Check if the file is not a directory
        If Not (GetAttr(FolderPath & FileName) And vbDirectory) = vbDirectory Then
            ' Delete the file
            On Error Resume Next
            Kill FolderPath & FileName
            On Error GoTo 0
        End If
        ' Get the next file
        FileName = Dir
    Loop
    
    '-Create Image of all side
    imgPath = FolderPath & "/Regional_Insights.jpg"
    pptTem.Slides(1).Export imgPath, "JPG"
    
    imgPath = FolderPath & "/Impact_Analysis.jpg"
    pptTem.Slides(2).Export imgPath, "JPG"
    
    imgPath = FolderPath & "/Key_Takeaways.jpg"
    pptTem.Slides(3).Export imgPath, "JPG"
    
    imgPath = FolderPath & "/Segmental_Insights.jpg"
    pptTem.Slides(4).Export imgPath, "JPG"
    
    imgPath = FolderPath & "/Market_KeyPlayer.jpg"
    pptTem.Slides(5).Export imgPath, "JPG"
    
End Function

