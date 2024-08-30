VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterForm 
   Caption         =   "EnterForm"
   ClientHeight    =   5530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8910.001
   OleObjectBlob   =   "EnterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ExtiBtn_Click()
 Unload Me
End Sub

Private Sub MiseAJourBtn_Click()
    Dim MainPage As Worksheet
    Dim SeuilColor As Integer
    
    If Me.SeuilCA.Value = "" Then
        SeuilColor = 1000
    Else
        SeuilColor = CInt(Me.SeuilCA.Value)
    
    End If
    Set MainPage = ThisWorkbook.Sheets("Main")
    
    If MainPage.Range("A1") <> "Date" Then
        Supprimer_Lignes_Colonnes
    End If
    Mise_a_jour_reporting SeuilColor
    
End Sub

Private Sub SauvegardeBtn_Click()
Dim cheminFichier As String
    
    ' V�rifie si le classeur a d�j� �t� enregistr�
    If ThisWorkbook.Path = "" Then
        ' Ouvre la bo�te de dialogue Enregistrer sous si le fichier n'a pas �t� enregistr�
        cheminFichier = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsm), *.xlsm")
        
        ' V�rifie si l'utilisateur a annul� la bo�te de dialogue
        If cheminFichier <> "False" Then
            ' Enregistre le fichier au chemin sp�cifi�
            ThisWorkbook.SaveAs Filename:=cheminFichier, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Else
            MsgBox "L'enregistrement a �t� annul�."
        End If
    Else
        ' Enregistre le fichier dans le dossier courant
        ThisWorkbook.Save
        MsgBox "Le fichier a �t� enregistr� avec succ�s dans le dossier courant."
    End If
End Sub

Private Sub UserForm_Click()

End Sub
