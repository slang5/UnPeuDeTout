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

    Worksheets("Exercice_2").Cells(5, 9).Value = 0
    Worksheets("Exercice_2").Cells(14, 5).Value = 0
    Worksheets("Exercice_2").Cells(6, 9).Value = 0
    Worksheets("Exercice_2").Cells(9, 9).Value = 0
    Worksheets("Exercice_2").Cells(15, 8).Value = 0

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
    
    Dim matricedate(0 To 2, 0 To 3) As Date
    For i = 0 To 3
        For j = 0 To 2
            matricedate(j, i) = Worksheets("Paramètres").Cells(5 + j, 11 + i).Value
        Next j
    Next i
    
    Dim datevoyage As Date: datevoyage = Worksheets("Exercice_2").Cells(14, 2).Value
    
    Dim matricereduction(0 To 1) As Double: matricereduction(0) = 0
    Dim matriceTVA(0 To 1) As Double: matriceTVA(0) = 0
    For i = 0 To 2
        If ((matricedate(i, 0)) <= (datevoyage)) And ((datevoyage) <= (matricedate(i, 1))) Then
            matricereduction(0) = matricedate(i, 2)
            matriceTVA(0) = matricedate(i, 3)
        End If
    Next i
    
    matricereduction(1) = matricereduction(0) * prix
    matriceTVA(1) = matriceTVA(0) * prix
    Dim prixtotal As Double: prixtotal = prix - matricereduction(1) + matriceTVA(1)
    
    Worksheets("Exercice_2").Cells(5, 9).Value = prix
    Worksheets("Exercice_2").Cells(14, 5).Value = (matricereduction(0))
    Worksheets("Exercice_2").Cells(6, 9).Value = matricereduction(1)
    Worksheets("Exercice_2").Cells(9, 9).Value = matriceTVA(1)
    Worksheets("Exercice_2").Cells(15, 8).Value = prixtotal
End Sub
