Attribute VB_Name = "mFormatSegment"

Function FormatSegmentForBot()
    Dim List1 As String, List2 As String, List3 As String
    
    'Visible the Hide sheet
    Sheet4.Visible = xlSheetVisible
    
    'Clear old data
    Sheet4.Range("R1").ClearContents
    Sheet4.Range("Q4:S6").ClearContents
    
    'No of segment
    If Sheet1.Range("K2").Value <> "" Then
        Sheet4.Range("R1").Value = 3
        
    ElseIf Sheet1.Range("J2").Value <> "" Then
        Sheet4.Range("R1").Value = 2
    Else
        Sheet4.Range("R1").Value = 1
    End If
    
    'Segment Title
    Sheet4.Range("Q4").Value = Sheet1.Range("H2").Value
    Sheet4.Range("Q5").Value = Sheet1.Range("J2").Value
    Sheet4.Range("Q6").Value = Sheet1.Range("K2").Value
    
    'Segment List
    List3 = ""
    For i = 3 To Sheet1.Cells(Rows.Count, "K").End(xlUp).Row
        List3 = List3 & Replace(Sheet1.Range("K" & i).Value, ">", "") & ", "
        
    Next i
    
    List2 = ""
    For i = 3 To Sheet1.Cells(Rows.Count, "J").End(xlUp).Row
        List2 = List2 & Replace(Sheet1.Range("J" & i).Value, ">", "") & ", "
        
    Next i
    
    List1 = ""
    For i = 3 To Sheet1.Cells(Rows.Count, "H").End(xlUp).Row
        List1 = List1 & Replace(Sheet1.Range("H" & i).Value, ">", "") & ", "
        
    Next i
    
    If List1 <> "" Then
        List1 = Left(List1, Len(List1) - 2)
    End If
    If List2 <> "" Then
        List2 = Left(List2, Len(List2) - 2)
    End If
    If List3 <> "" Then
        List3 = Left(List3, Len(List3) - 2)
    End If
    
    Sheet4.Range("R4").Value = List1
    Sheet4.Range("R5").Value = List2
    Sheet4.Range("R6").Value = List3
    
    'Dominating Segment
    Sheet4.Range("S4").Value = Replace(Sheet1.Range("H3").Value, ">", "")
    Sheet4.Range("S5").Value = Replace(Sheet1.Range("J3").Value, ">", "")
    Sheet4.Range("S6").Value = Replace(Sheet1.Range("K3").Value, ">", "")
    
    'Driver
    Sheet4.Range("R8").Value = Sheet1.Range("C15").Value
    Sheet4.Range("R9").Value = Sheet1.Range("C16").Value
    
    'Restrain
    Sheet4.Range("R10").Value = Sheet1.Range("C17").Value
    
    'Opertunities
    Sheet4.Range("R11").Value = Sheet1.Range("C19").Value
    
    'check Output folder and create it
    FolderPath = ThisWorkbook.Path & "/Output"
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    'check Year folder and create it
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "yyyy")
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    'check Month folder and create it
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "yyyy") & "/" & Format(Date, "MMM-yy")
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    'check Day folder and create it
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "yyyy") & "/" & Format(Date, "MMM-yy") & "/" & Format(Date, "dd")
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    'check Market folder and create it
    FolderPath = ThisWorkbook.Path & "/Output/" & Format(Date, "yyyy") & "/" & Format(Date, "MMM-yy") & "/" & Format(Date, "dd") & "/" & Sheet1.Range("D2").Value
    If Dir(FolderPath, vbDirectory) = "" Then
        ' If the folder doesn't exist, create it
        MkDir FolderPath
    End If
    
    
End Function

Function test()
    For i = 4 To 3
        MsgBox i
    Next i
End Function

Function HideWorksheet()
    Sheet4.Visible = xlSheetVisible
    
End Function

