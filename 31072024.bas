Option Explicit
Option Base 1

Sub RechercheEtAcquisition()
    Dim PATH As String
    Dim Cible1 As String
    Dim Cible2 As String
    Dim Cible3 As String
    
    PATH = ActiveSheet.Cells(1, 2).Value
    Cible1 = ActiveSheet.Cells(2, 2).Value
    Cible2 = ActiveSheet.Cells(3, 2).Value
    Cible3 = ActiveSheet.Cells(4, 2).Value
    
    Dim App As New Application
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim TargetedWS As Worksheet
    Dim C1 As String, C2 As String, C3 As String
    Dim indicatrice As Integer 'La sheet recherchee a ete trouvee
    indicatrice = 0
    
    App.Visible = True
    Set WB = App.Workbooks.Open(PATH)
    
    For Each WS In WB.Sheets
        With WS
        C1 = .Cells(2, 2).Value
        C2 = .Cells(7, 3).Value
        C3 = .Cells(9, 5).Value
        End With
        If Verif(C1, Cible1) = True Then
            If Verif(C2, Cible2) = True Then
                If Verif(C3, Cible3) = True Then
                    Set TargetedWS = WS
                    indicatrice = 1
                    Exit For
                End If
            End If
        End If
        
        If Not indicatrice = 1 Then
            C1 = "9999999"
            C2 = "9999999"
            C3 = "9999999"
        End If
    Next WS
    If indicatrice = 0 Then
        MsgBox "Le produit suivant " & C1 & " | " & C2 & " | " & C3 & " n'a pas ete trouve. Cela n'a aucune consequence grave, il faudra completer manuellement", vbExclamation, "Produit non trouvé."
    Else
        ActiveSheet.Cells(2, 3).Value = C1
        ActiveSheet.Cells(3, 3).Value = C2
        ActiveSheet.Cells(4, 3).Value = C3
        ActiveSheet.Cells(6, 3).Value = WS.Cells(100, 1).Value
    End If
    
    
    WB.Close False
    App.Quit
    Set WB = Nothing
    Set App = Nothing
End Sub

Function Verif(Input1, Input2)
    If Input1 = Input2 Then
        Verif = True
    Else
        Verif = False
    End If
End Function
