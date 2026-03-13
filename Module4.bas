Attribute VB_Name = "Module4"

Sub test()
    Dim wApp As Word.Application
    Dim wDoc As Word.Document
    
    Set wApp = New Word.Application
    wApp.Visible = True
    Set wDoc = wApp.Documents.Add
    
    wDoc.Paragraphs.Add
    wDoc.Range.Text = "here is something"
    wDoc.Paragraphs(wDoc.Paragraphs.Count).Range.Select
    wApp.ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1).NumberFormat = ChrW(&H2022)
    wDoc.ListTemplates(1).ListLevels(1).NumberFormat = ChrW(&H2022)
    
    'wApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=wApp.ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    wDoc.Paragraphs(1).Range.ListFormat.ApplyListTemplate wApp.ListGalleries(wdBulletGallery).ListTemplates(1), False
    
    wDoc.Paragraphs.Add
    wDoc.Range.Text = "Here is my line"
    
End Sub
