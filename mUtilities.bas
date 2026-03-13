Attribute VB_Name = "mUtilities"

'****First Label Bullet********
Function AddFirstLabelBullet(sCursor As Long, eCursor As Long)
    wDoc.Range(Start:=sCursor, End:=eCursor).Select
    
    With wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "o" 'ChrW(&H25A0)
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = wApp.CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = wApp.CentimetersToPoints(1.27)
        .Font.Name = "Courier New"
    End With
    
    wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    wApp.Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    
End Function



'****Second Label Bullet********
Function AddSecondLabelBullet(sCursor As Long, eCursor As Long)
    wDoc.Range(Start:=sCursor, End:=eCursor).Select
    
    With wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(61607)  '"&H25CF"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = wApp.CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = wApp.CentimetersToPoints(1.27)
        .Font.Name = "Wingdings"
    End With
    
    wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    
    wApp.Selection.Range.ListFormat.ListIndent
    'wDoc.Range.ListFormat.para
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    wApp.Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    
End Function


'****Third Label Bullet********
Function AddThirdLabelBullet(sCursor As Long, eCursor As Long)
    wDoc.Range(Start:=sCursor, End:=eCursor).Select
    
    With wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(8226) 'ChrW(9679) '"&H25A1"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = wApp.CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = wApp.CentimetersToPoints(1.27)
        .Font.Name = "Courier New"
        '.Font.Size = 5
    End With
    
    wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    wApp.Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    
End Function


'****Third Label Bullet********
Function AddFourthLabelBullet(sCursor As Long, eCursor As Long)
    wDoc.Range(Start:=sCursor, End:=eCursor).Select
    
    With wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(&H25AB) '"o"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = wApp.CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = wApp.CentimetersToPoints(1.27)
        .Font.Name = "Courier New"
    End With
    
    wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.Range.ListFormat.ListIndent
    wApp.Selection.EndKey unit:=wdStory
    wApp.Selection.TypeParagraph
    wApp.Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    
End Function


'****First Label Bullet in Table Cell********
Function AddFirstLabelBulletInTable()
    
    With wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "&H25A0" '"o"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = wApp.CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = wApp.CentimetersToPoints(1.27)
        .Font.Name = "Courier New"
    End With
    
    wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    
    
End Function




