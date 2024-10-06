Sub TurnOffDuringProcess(App As Application, IsEnd As Boolean)
    With App
        If IsEnd <> True Then
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
        Else
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

Sub FaireUnExtract(ClientID As String, Position As String)
    Dim SQL As New SQL_InVBA
    SQL.ConnectionDB ("###Path/WorkbookSource.xlsx")
    SQL.DesignRequest (ClientID)
    SQL.ExecuteRequest
    SQL.PrintOutput (Position)
End Sub

Sub GetArray()
    Dim maxRow As Integer
    maxRow = DB_Mapping.Range("a1").End(xlDown).Row
    Dim Lib As New Scripting.Dictionary
    Dim rng() As Variant
    
    ReDim rng(1 To maxRow - 1, 1 To 2)
    rng() = DB_Mapping.Range("a2:b" & CStr(maxRow)).Value2
    Dim iter As Integer
    For iter = 1 To maxRow - 1
        Lib.Add (rng(iter, 1)), rng(iter, 2)
    Next iter
    Dim IsIn As String
    IsIn = Lib("fdsf")
    MsgBox IsIn
End Sub


Option Explicit
Option Base 1

Dim connection As Object
Dim Request As String
Dim OutputClass As Object

Sub Class_Initialize()
    Debug.Print "Object initialisé"
    Set connection = CreateObject("ADODB.connection")
    Set OutputClass = CreateObject("ADODB.Recordset")
End Sub

Sub ConnectionDB(Path As String)
    connection.Provider = "Microsoft.ACE.OLEDB.16.0"
    connection.ConnectionString = "Data Source=" & Path & ";" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    connection.Open
    Debug.Print "Connection à la base dont le chemin d'accès est : " & Path
End Sub

Sub DesignRequest(ClientID As String)
    Request = "SELECT * FROM [Query1$] Where [Client] = '" & CStr(ClientID) & "' ;"
    Debug.Print ("La request SQL est : " & Request)
    Debug.Print ("Client concerné : " & ClientID)
End Sub

Sub ExecuteRequest()
    OutputClass.Open Request, connection
    Debug.Print "Requete realisée"
End Sub

Sub PrintOutput(Position As String)
    Output.Cells.Clear
    Output.Range(Position).CopyFromRecordset OutputClass
End Sub


