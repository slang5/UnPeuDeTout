Sub Le_Mail()
'Ne pas oublier d'activer dans l'onglet {Outils > References} la référence "Microsoft Outlook 16.0 Object Library"
    Dim OutlookApp As Object
    Dim Mail As Object
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Mail = OutlookApp.CreateItem(0)
    
    Dim Fichertemporaire As String
    Dim NameExcelSheet as String 
    NameExcelSheet = ""
    Dim NameOfFile as String
    NameOfFile = ""
    
    ThisWorkbook.Sheets(NameExcelSheet).Copy
    Application.DisplayAlerts = False
    Set Feuilletemporaire = ActiveWorkbook
    Feuilletemporaire.SaveAs Filename:=NameOfFile
    
    With Mail
        .To = "test@test.test"
        .Subject = "" 
        .BCC = ""
        .CC = ""
        .Body = ""
        .Attachments.Add Feuilletemporaire.FullName
    End With
    
    Mail.Display
    Workbooks(NameOfFile + ".xlsx").Close SaveChanges:=False
End Sub
