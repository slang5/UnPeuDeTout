Sub Reporting(Dep, Arr, DateCom, DateVoy, Pri, Red, Tva)

    Worksheets("Historique des billets").Rows(2).Insert (1)

    Dim MatriceDate(0 To 1) As Variant
    MatriceDate(0) = DateCom
    MatriceDate(1) = DateVoy

    With Worksheets("Historique des billets").Rows(2)
    .Columns(1).Value = Dep
    .Columns(2).Value = Arr
    .Columns(11).Value = Pri
    .Columns(12).Value = Red
    .Columns(13).Value = Tva

    For i = 0 To 1:
    .Columns(3 + 4 * i).Value = MatriceDate(i)
    .Columns(4 + 4 * i).Value = Day(MatriceDate(i))
    .Columns(5 + 4 * i).Value = Month(MatriceDate(i))
    .Columns(6 + 4 * i).Value = Year(MatriceDate(i))
    Next i
        
    End With
End Sub

Sub Ticket()

    Dim erreur(0 To 1) As String: erreur(0) = "": erreur(1) = ""
    Dim nombre As Integer: nombre = 0

    If Worksheets("Exercice_2").Cells(5, 2).Value = Worksheets("Exercice_2").Cells(6, 2).Value Then
        nombre = nombre + 1
        erreur(0) = "Les villes de départ et d'arrivée sont les mêmes"
    End If

    If IsEmpty(Worksheets("Exercice_2").Cells(14, 2).Value) Then
        nombre = nombre + 1
        erreur(1) = "La date du voyage n'est pas renseignée"
    End If

    If nombre > 0 Then
        Dim titre As String: titre = "Liste erreurs"
        MsgBox "Les erreurs sont les suivantes : " + vbCrLf + erreur(0) + vbCrLf + erreur(1), 0, ["Liste des erreurs"]
        Exit Sub
    End If
    With Worksheets("Exercice_2")
        .Cells(5, 9).Value = 0
        .Cells(14, 5).Value = 0
        .Cells(6, 9).Value = 0
        .Cells(9, 9).Value = 0
        .Cells(15, 8).Value = 0
    end With

    Dim villes(0 To 7) As String 'ici on a une matrice de taille 8
    Dim Index As Integer
    For Index = 0 To 7
        villes(Index) = Worksheets("Paramètres").Cells(4 + Index, 1).Value
    Next Index
    
    Dim voyage(0 To 1) As String
    For Index = 0 To 1
        voyage(Index) = Worksheets("Exercice_2").Cells(5 + Index, 2).Value
    Next Index
    

    Dim voyageLigneColonne(0 To 1) As Integer
    For Index = 0 To 1
        For i = 0 To 7
            If villes(i) = voyage(Index) Then
                voyageLigneColonne(Index) = i
            End If
        Next i
    Next Index

    Dim distance As Integer
    distance = Worksheets("Paramètres").Cells(4 + voyageLigneColonne(0), 2 + voyageLigneColonne(1)).Value
    Worksheets("Exercice_2").Cells(7, 2).Value = distance


    Dim prix As Double
    prix = distance * Worksheets("Paramètres").Cells(16 + voyageLigneColonne(0), 2 + voyageLigneColonne(1)).Value
    
    

    Dim MatriceDate(0 To 2, 0 To 3) As Date
    For i = 0 To 3
        For j = 0 To 2
            MatriceDate(j, i) = Worksheets("Paramètres").Cells(5 + j, 11 + i).Value
        Next j
    Next i
    

    Dim datevoyage As Date: datevoyage = Worksheets("Exercice_2").Cells(14, 2).Value
    


    Dim matricereduction(0 To 1) As Double: matricereduction(0) = 0
    Dim matriceTVA(0 To 1) As Double: matriceTVA(0) = 0
    For i = 0 To 2
        If ((MatriceDate(i, 0)) <= (datevoyage)) And ((datevoyage) <= (MatriceDate(i, 1))) Then
            matricereduction(0) = MatriceDate(i, 2)
            matriceTVA(0) = MatriceDate(i, 3)
        End If
    Next i
    
    matricereduction(1) = matricereduction(0) * prix
    matriceTVA(1) = matriceTVA(0) * prix
    Dim prixtotal As Double: prixtotal = prix - matricereduction(1) + matriceTVA(1)
    With Worksheets("Exercice_2")
        .Cells(5, 9).Value = prix
        .Cells(14, 5).Value = (matricereduction(0))
        .Cells(6, 9).Value = matricereduction(1)
        .Cells(9, 9).Value = matriceTVA(1)
        .Cells(15, 8).Value = prixtotal
        a = .Cells(5, 2).Value
        b = .Cells(6, 2).Value
        c = .Cells(17, 2).Value
    end with

    D = datevoyage
    e = prixtotal
    f = matricereduction(0)
    g = matriceTVA(0)
    
    Reporting a, b, c, D, e, f, g
    
End Sub

Sub ClearHistorique()

    Dim Ligne As Integer: Ligne = 1
    Dim Binaire As Boolean: Binaire = False

    While Binaire = False
        If Worksheets("Historique des billets").Cells(Ligne, 1).Value <> 0 Then
            Ligne = Ligne + 1
   
        Else
        Binaire = True
        End If
     
    Wend

    Ligne = Ligne - 1
    If Ligne <= 1 Then
        Exit Sub
    Else 
    Worksheets("Historique des billets").Range(Cells(2, 1), Cells(Ligne, 13)).Clear
    End If

End Sub


Sub VBA_Presentation()
    
    Const msoTrue = -1
    Const ppWindowMaximized = 2
    
    Dim PApplication As Object
    Dim PPT As Object
    Dim PPTSlide As Object
    Dim PPTCharts As Excel.ChartObject


    Set PApplication = CreateObject("PowerPoint.Application")
    PApplication.Visible = msoTrue
    PApplication.WindowState = ppWindowMaximized

    Set PPT = PApplication.Presentations.Add

    For Each PPTCharts In Worksheets("Slide2PPT").ChartObjects

        Set PPTSlide = PPT.Slides.Add(PPT.Slides.Count + 1, 12)

        PPTCharts.Chart.ChartArea.Copy
        PPTSlide.Shapes.PasteSpecial(DataType:=0).Select

    Next PPTCharts
    
    Set PPTSlide = Nothing
    Set PPT = Nothing
    Set PApplication = Nothing

End Sub
