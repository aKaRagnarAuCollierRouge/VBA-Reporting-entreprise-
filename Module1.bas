Attribute VB_Name = "Module1"
Sub Mise_a_jour_reporting(SeuilColor As Integer)

   
'Création d'une nouvelle colonne
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1").Select
'Nom de la colonne
ActiveCell.FormulaR1C1 = "Chiffre d'affaires"
Range("K2").Select
Application.CutCopyMode = False
'Formule pour calculer le chiffre d'affaires par produit
ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K45")
Range("K2:K45").Select
'changement du format
Selection.NumberFormat = "#,##0.00 $"
Columns("K:K").Select
Range("K2").Activate
'Changement de la couleur et gras sur la colonne

'ajout des lignes
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'création de la somme pour le calcul du CA total
Range("K1").Select
ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[46]C)"

'sous_programme pour le titre
titre_reporting
'sous programme pour la mise en forme du CA
mise_en_forme_CA SeuilColor
'sous programme pour la mise en forme de Glutenfree
glutenfree

'changement des données de la colonne Bio
Range("F4:F100").Select
Selection.Replace What:="Oui", Replacement:="Bio", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

Sheets.Add(After:=ActiveSheet).Name = "Calcul"
    
Range("A1").Value = "Calcul de la quantité moyenne de vente"
Range("A2").Value = "Calcul de la plus petite vente"
Range("A3").Value = "calcul de la plus grande vente"
Range("A4").Value = "nombre de vente à 0"
Range("A5").Value = "Somme du nombre de vente"
Range("B1").Value = "=AVERAGE(Main!R[1]C[8]:R[44]C[8])"
Range("B2").Value = "=MIN(Main!RC[8]:R[43]C[8])"
Range("B3").Value = "=MAX(Main!R[-1]C[8]:R[42]C[8])"
Range("B4").Value = "=COUNTIF(Main!RC[8]:R[43]C[8],0)"
Range("B5").Value = "=SUM(Main!R[-3]C[8]:R[40]C[8])"

End Sub

Sub glutenfree()

Dim modif_var As String

i = 4
Do While i < 100
    modif_var = Range("E" & i).Value
    If modif_var = "Oui" Then
    Range("E" & i).Value = "Gluten Free"
    Else
    End If
i = i + 1
Loop

End Sub

Sub mise_en_forme_CA(SeuilColor As Integer)

Dim test_CA As Double

'Mise en forme des cellules CA

For i = 4 To 100

test_CA = Range("K" & i).Value
Range("K" & i).Select
'En rouge et jaune
    If test_CA < SeuilColor Then
        'En rouge
        With Selection.Font
            .Color = -16776961
            .Bold = True
        End With
    Else
        'en vert
        With Selection.Font
            .Color = -11489280
            .Bold = True
        End With
    End If
Next i

End Sub

Sub titre_reporting()

Dim mois As Integer
Dim mois_titre As String

    'calcul de la date du jour
    Range("A1").Value = "=TODAY()"
    
    'calul du mois
    Range("B1").Value = "=MONTH(RC[-1])"
    
    'chargement de la variable mois
    mois = Range("B1").Value
    
    'selection du mois pour le titre du reporting
    Select Case mois
        Case 1
        mois_titre = "Janvier"
        Case 2
        mois_titre = "Fevrier"
        Case 3
        mois_titre = "Mars"
        Case 4
        mois_titre = "Avril"
        Case 5
        mois_titre = "Mai"
        Case 6
        mois_titre = "Juin"
        Case 7
        mois_titre = "Juillet"
        Case 8
        mois_titre = "Aout"
        Case 9
        mois_titre = "Septembre"
        Case 10
        mois_titre = "Octobre"
        Case 11
        mois_titre = "Novembre"
        Case 12
        mois_titre = "Decembre"
    End Select

'Titre du reporting
Range("E1").Value = "Reporting du mois de " & mois_titre

End Sub

Sub Supprimer_Lignes_Colonnes()
    ' Référence à la feuille "Main"
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Sheets("Main")
    
    ' Suppression des lignes A1 et A2 dans la feuille "Main"
    wsMain.Rows("1:2").Delete

    ' Recherche de la colonne "Chiffre d'affaires" dans la feuille "Main"
    Dim col As Integer
    Dim found As Range

    Set found = wsMain.Rows(1).Find(What:="Chiffre d'affaires", LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        col = found.Column
        ' Suppression de la colonne dans la feuille "Main"
        wsMain.Columns(col).Delete
    Else
        MsgBox "La colonne 'Chiffre d'affaires' n'a pas été trouvée dans la feuille 'Main'."
    End If
    
    ' Suppression de la feuille de calcul "Calcul" si elle existe
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Calcul")
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        MsgBox "La feuille 'Calcul' a été supprimée."
    Else
        MsgBox "La feuille 'Calcul' n'a pas été trouvée."
    End If
End Sub


