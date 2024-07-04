Sub InsertExcelObjectInMail()
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    
    ' Configuration de la plage à copier
    Set ws = ThisWorkbook.Sheets("VotreFeuille")  ' Modifiez pour correspondre à votre nom de feuille
    ws.Range("A1:A115").Copy
    
    ' Ouvrir Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Créer un nouvel e-mail
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .Display
        .To = "exemple@exemple.com"
        .Subject = "Incorporation d'un objet Excel"
        .GetInspector.WordEditor.Windows(1).Selection.PasteSpecial Link:=True, DataType:=14 ' 14 pour un objet OLE, Link pour lier
        .Display
    End With
    
    ' Nettoyage
    Set olMail = Nothing
    Set olApp = Nothing
End Sub
