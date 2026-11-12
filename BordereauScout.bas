Attribute VB_Name = "BordereauScout"
Option Explicit

' =====================================================================
'  BORDEREAU DE MATÉRIEL - module complémentaire de l'inventaire scout
'  - Feuille Bordereau : le bordereau en cours de préparation (imprimable / PDF)
'  - Feuille Prets     : le journal de toutes les sorties et de leurs retours
'  L'inventaire n'est pas modifié par une sortie : la quantité disponible est
'  calculée (stock - sorties non rendues).
' =====================================================================

Private Const F_ACCUEIL As String = "Accueil"
Private Const F_INVENTAIRE As String = "Inventaire"
Private Const F_DATA As String = "Data"
Private Const F_BORD As String = "Bordereau"
Private Const F_PRETS As String = "Prets"

' Numérotation des bordereaux : préfixe + 3 chiffres (BOD001, BOD002...)
Private Const PREFIXE_NUM As String = "BOD"

Private Const L_INV_ENTETE As Long = 5

' --- Feuille Bordereau : en-tête
Private Const CELL_GROUPE As String = "C3"
Private Const CELL_DATE As String = "C4"
Private Const CELL_RESP As String = "C5"
Private Const CELL_DEST As String = "C6"
Private Const CELL_NUM As String = "C7"

' --- Feuille Bordereau : tableau (colonnes A à E imprimées, G et H hors impression)
Private Const L_B_ENTETE As Long = 10
Private Const L_B_DEBUT As Long = 11
Private Const L_B_FIN As Long = 40
Private Const L_B_TOTAL As Long = 42
Private Const B_NUMLIG As Long = 1
Private Const B_NOM As Long = 2
Private Const B_QTE As Long = 3
Private Const B_ETAT As Long = 4
Private Const B_OBS As Long = 5
Private Const B_ID As Long = 7
Private Const B_TYPE As Long = 8

' --- Colonnes de la feuille Prets
Private Const P_NUM As Long = 1
Private Const P_DATE As Long = 2
Private Const P_RESP As Long = 3
Private Const P_GROUPE As Long = 4
Private Const P_DEST As Long = 5
Private Const P_ID As Long = 6
Private Const P_TYPE As Long = 7
Private Const P_NOM As Long = 8
Private Const P_QTE As Long = 9
Private Const P_ETAT As Long = 10
Private Const P_OBS As Long = 11
Private Const P_RETOUR As Long = 12
Private Const P_ETATRET As Long = 13


' =====================================================================
'  INSTALLATION (à lancer une seule fois après l'import du module)
'  Relancer l'installation reconstruit la feuille Bordereau (le bordereau
'  en cours de préparation est effacé) mais conserve le journal Prets.
' =====================================================================
Public Sub InstallerBordereau()
    Dim wsB As Worksheet, wsP As Worksheet, ws As Worksheet, g As Double, h As Double

    Set wsB = FeuilleOuCree(F_BORD)
    Set wsP = FeuilleOuCree(F_PRETS)
    MiseEnFormeBordereau wsB
    MiseEnFormePrets wsP

    g = wsB.Range("J3").Left
    h = wsB.Range("J3").Top
    PoserBouton wsB, "btnBValider", "Valider et créer le PDF", "ValiderBordereau", g, h
    PoserBouton wsB, "btnBRetirer", "Retirer la ligne", "RetirerLigneBordereau", g, h + 34
    PoserBouton wsB, "btnBNouveau", "Nouveau bordereau", "NouveauBordereau", g, h + 68
    PoserBouton wsB, "btnBInv", "Voir l'inventaire", "AllerInventaireDepuisBordereau", g, h + 102
    PoserBouton wsB, "btnBPrets", "Voir les sorties", "AllerPrets", g, h + 136

    g = wsP.Range("O2").Left
    h = wsP.Range("O2").Top
    PoserBouton wsP, "btnPRetour", "Enregistrer le retour", "EnregistrerRetour", g, h
    PoserBouton wsP, "btnPBord", "Voir le bordereau", "AllerBordereau", g, h + 34
    PoserBouton wsP, "btnPInv", "Voir l'inventaire", "AllerInventaireDepuisBordereau", g, h + 68

    Set ws = ThisWorkbook.Worksheets(F_INVENTAIRE)
    PoserBouton ws, "btnInvAjoutBord", "Ajouter au bordereau", "AjouterAuBordereau", ws.Range("A3").Left + 5 * 148, ws.Range("A3").Top + 4
    PoserBouton ws, "btnInvVoirBord", "Voir le bordereau", "AllerBordereau", ws.Range("A3").Left + 6 * 148, ws.Range("A3").Top + 4

    Set ws = ThisWorkbook.Worksheets(F_ACCUEIL)
    PoserBouton ws, "btnAccBord", "Bordereau de matériel", "AllerBordereau", ws.Range("E4").Left, ws.Range("E4").Top

    InstallerIndicateurs
    wsB.Activate
    MsgBox "Bordereau de matériel installé." & vbCrLf & vbCrLf & _
           "Pensez à enregistrer le classeur (toujours au format .xlsm).", vbInformation, "Inventaire scout"
End Sub

Private Function FeuilleOuCree(ByVal nom As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If
    Set FeuilleOuCree = ws
End Function

Private Sub PoserBouton(ws As Worksheet, ByVal nom As String, ByVal texte As String, ByVal macro As String, _
                        ByVal gauche As Double, ByVal haut As Double)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nom).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, gauche, haut, 150, 26)
    With shp
        .Name = nom
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = texte
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
        .OnAction = macro
        .Placement = xlFreeFloating
    End With
End Sub

Private Sub MiseEnFormeBordereau(ws As Worksheet)
    Dim r As Long, titres As Variant, libelles As Variant, i As Long, groupe As String

    ' On conserve le nom du groupe s'il était déjà saisi
    On Error Resume Next
    groupe = CStr(ws.Range(CELL_GROUPE).Value)
    On Error GoTo 0

    With ws
        .Cells.UnMerge
        .Cells.Clear
        .Cells.Font.Name = "Times New Roman"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 34
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 34
        .Columns("F").ColumnWidth = 3
        .Columns("G").ColumnWidth = 9
        .Columns("H").ColumnWidth = 14
        .Columns("I").ColumnWidth = 3
        .Columns("J").ColumnWidth = 32

        ' --- Titre
        .Range("A1:E1").Merge
        .Range("A1").Value = "BORDEREAU DE MATÉRIEL"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(0, 112, 192)
        .Range("A1").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 34

        ' --- En-tête
        libelles = Array("Groupe Local :", "Date :", "Responsable :", "Destination :", "N° de bordereau :")
        For i = 0 To 4
            r = 3 + i
            .Range("A" & r & ":B" & r).Merge
            .Range("A" & r).Value = libelles(i)
            .Range("A" & r).Font.Bold = True
            .Range("A" & r).Font.Size = 14
            With .Range("C" & r & ":E" & r)
                .Merge
                .HorizontalAlignment = xlLeft
                .Font.Size = 14
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(166, 166, 166)
            End With
            .Rows(r).RowHeight = 24
        Next i
        .Range("C4:E4").NumberFormat = "dd/mm/yyyy"
        .Range(CELL_GROUPE).Value = groupe
        .Range(CELL_DATE).Value = Date

        ' --- Tableau
        .Range("A9").Value = "Détails du Matériel :"
        .Range("A9").Font.Size = 16
        .Range("A9").Font.Bold = True
        .Range("A9").Font.Color = RGB(31, 78, 121)
        .Rows(9).RowHeight = 26

        titres = Array("N°", "Désignation", "Nb", "État", "Observations")
        For i = 0 To 4
            .Cells(L_B_ENTETE, i + 1).Value = titres(i)
        Next i
        With .Range(.Cells(L_B_ENTETE, 1), .Cells(L_B_ENTETE, 5))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 32, 96)
            .HorizontalAlignment = xlCenter
        End With
        With .Range(.Cells(L_B_DEBUT, 1), .Cells(L_B_FIN, 5))
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        .Range(.Cells(L_B_DEBUT, B_NUMLIG), .Cells(L_B_FIN, B_NUMLIG)).HorizontalAlignment = xlCenter
        .Range(.Cells(L_B_DEBUT, B_QTE), .Cells(L_B_FIN, B_QTE)).HorizontalAlignment = xlCenter

        ' Colonnes techniques, hors impression
        .Cells(L_B_ENTETE, B_ID).Value = "ID objet"
        .Cells(L_B_ENTETE, B_TYPE).Value = "Type"
        With .Range(.Cells(L_B_ENTETE, B_ID), .Cells(L_B_FIN, B_TYPE))
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(128, 128, 128)
        End With

        ' --- Total et signatures
        .Range("A" & L_B_TOTAL & ":B" & L_B_TOTAL).Merge
        .Range("A" & L_B_TOTAL).Value = "Total des articles :"
        .Cells(L_B_TOTAL, B_QTE).Formula = "=SUM(C" & L_B_DEBUT & ":C" & L_B_FIN & ")"
        .Range("A" & L_B_TOTAL & ":C" & L_B_TOTAL).Font.Bold = True
        .Range("A" & L_B_TOTAL & ":C" & L_B_TOTAL).Font.Size = 14
        .Cells(L_B_TOTAL, B_QTE).HorizontalAlignment = xlCenter

        .Range("A44").Value = "Signatures :"
        .Range("A46").Value = "Responsable de la réception :"
        .Range("A50").Value = "Responsable de la remise :"
        .Range("A44,A46,A50").Font.Bold = True
        .Range("A44,A46,A50").Font.Size = 14
        .Range("B48:D48").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B52:D52").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Rows(48).RowHeight = 30
        .Rows(52).RowHeight = 30

        ' --- Mode d'emploi (hors impression)
        With .Range("J1,J13:J16")
            .Font.Name = "Arial"
            .Font.Size = 9
            .Font.Color = RGB(89, 89, 89)
        End With
        .Range("J1").Value = "Mode d'emploi"
        .Range("J1").Font.Bold = True
        .Range("J13").Value = "1. Feuille Inventaire : cliquez sur un objet puis sur « Ajouter au bordereau »."
        .Range("J14").Value = "2. Ici : remplissez groupe, responsable et destination, corrigez l'état et ajoutez vos observations."
        .Range("J15").Value = "3. Cliquez sur « Valider et créer le PDF »."
        .Range("J16").Value = "Seules les colonnes A à E sont imprimées."

        On Error Resume Next
        With .Range(.Cells(L_B_DEBUT, B_ETAT), .Cells(L_B_FIN, B_ETAT)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=Referentiels!$B$2:$B$40"
            .ShowError = False
        End With
        With .PageSetup
            .PrintArea = "$A$1:$E$53"
            .Orientation = xlPortrait
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
        End With
        On Error GoTo 0
    End With
End Sub

' Met (ou remet) les titres et le format du journal, sans toucher aux lignes déjà enregistrées
Private Sub MiseEnFormePrets(ws As Worksheet)
    Dim titres As Variant, largeurs As Variant, i As Long
    titres = Array("N° bordereau", "Date sortie", "Responsable", "Groupe local", "Destination", "ID objet", "Type", _
                   "Désignation", "Nb", "État au départ", "Observations", "Date retour", "État au retour")
    largeurs = Array(13, 12, 22, 20, 24, 9, 14, 34, 8, 15, 30, 12, 16)
    With ws
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 10
        For i = 0 To 12
            .Cells(1, i + 1).Value = titres(i)
            .Columns(i + 1).ColumnWidth = largeurs(i)
        Next i
        With .Range(.Cells(1, 1), .Cells(1, 13))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With
        .Columns(P_DATE).NumberFormat = "dd/mm/yyyy"
        .Columns(P_DEST).NumberFormat = "General"
        .Columns(P_RETOUR).NumberFormat = "dd/mm/yyyy"
        .Columns(14).ColumnWidth = 3
        .Columns(15).ColumnWidth = 30
        .Range("O1").Value = "Cliquez sur une ligne non rendue, puis sur « Enregistrer le retour »."
        .Range("O1").Font.Size = 9
        .Range("O1").Font.Color = RGB(89, 89, 89)
    End With
End Sub


' =====================================================================
'  INDICATEURS DE SUIVI SUR L'ACCUEIL (colonnes G et H)
'  Formules recalculées automatiquement à partir de la feuille Prets.
'  Peut être relancé sans risque : ne touche ni au bordereau ni au journal.
' =====================================================================
Public Sub InstallerIndicateurs()
    Dim ws As Worksheet, wsP As Worksheet, libelles As Variant, formules As Variant, i As Long

    Set wsP = FeuilleOuCree(F_PRETS)
    If Trim$(CStr(wsP.Range("A1").Value)) = "" Then MiseEnFormePrets wsP
    Set ws = ThisWorkbook.Worksheets(F_ACCUEIL)

    libelles = Array("Bordereaux en cours", "Objets sortis (lignes non rendues)", "Quantité sortie", _
                     "Quantité disponible", "Objets rendus (lignes)", "Retours avec changement d'état")
    ' Dans ces formules, l'apostrophe remplace le guillemet (converti plus bas)
    formules = Array( _
        "=SUMPRODUCT((Prets!$A$2:$A$2000<>'')*(Prets!$L$2:$L$2000='')/(COUNTIFS(Prets!$A$2:$A$2000,Prets!$A$2:$A$2000&'',Prets!$L$2:$L$2000,'')+1-(Prets!$A$2:$A$2000<>'')*(Prets!$L$2:$L$2000='')))", _
        "=COUNTIFS(Prets!$A$2:$A$2000,'<>',Prets!$L$2:$L$2000,'')", _
        "=SUMIFS(Prets!$I$2:$I$2000,Prets!$A$2:$A$2000,'<>',Prets!$L$2:$L$2000,'')", _
        "=C5-H6", _
        "=COUNTIF(Prets!$L$2:$L$2000,'<>')", _
        "=SUMPRODUCT((Prets!$L$2:$L$2000<>'')*(Prets!$M$2:$M$2000<>Prets!$J$2:$J$2000))")

    With ws
        .Columns("G").ColumnWidth = 34
        .Columns("H").ColumnWidth = 10
        .Range("G3").Value = "Suivi des prêts et des retours"
        .Range("G3").Font.Name = "Arial"
        .Range("G3").Font.Size = 12
        .Range("G3").Font.Bold = True
        .Range("G3").Font.Color = RGB(31, 78, 121)
        For i = 0 To 5
            .Cells(4 + i, 7).Value = libelles(i)
            .Cells(4 + i, 8).Formula = Replace(formules(i), "'", Chr$(34))
        Next i
        With .Range("G4:H9")
            .Font.Name = "Arial"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
        .Range("H4:H9").Font.Bold = True
        .Range("H4:H9").HorizontalAlignment = xlCenter
        .Range("H4:H9").Interior.Color = RGB(221, 235, 247)
        PoserBouton ws, "btnAccSorties", "Voir les sorties", "AllerPrets", .Range("G11").Left, .Range("G11").Top
    End With
End Sub

Public Sub LancerInstallationIndicateurs()
    InstallerIndicateurs
    ThisWorkbook.Worksheets(F_ACCUEIL).Activate
    MsgBox "Indicateurs de suivi ajoutés sur l'accueil (colonnes G et H)." & vbCrLf & _
           "Pensez à enregistrer le classeur.", vbInformation, "Inventaire scout"
End Sub


' =====================================================================
'  NAVIGATION
' =====================================================================
Public Sub AllerBordereau()
    ThisWorkbook.Worksheets(F_BORD).Activate
End Sub

Public Sub AllerPrets()
    ThisWorkbook.Worksheets(F_PRETS).Activate
End Sub

Public Sub AllerInventaireDepuisBordereau()
    ThisWorkbook.Worksheets(F_INVENTAIRE).Activate
End Sub


' =====================================================================
'  OUTILS
' =====================================================================
Private Function DerniereLignePrets() As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(F_PRETS)
    r = ws.Cells(ws.Rows.Count, P_NUM).End(xlUp).Row
    If r < 1 Then r = 1
    DerniereLignePrets = r
End Function

' Quantité en stock d'un objet (ligne active de la feuille Data), 0 s'il n'existe plus
Private Function QuantiteStock(ByVal idObjet As Long) As Long
    Dim ws As Worksheet, r As Long, der As Long
    Set ws = ThisWorkbook.Worksheets(F_DATA)
    der = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To der
        If ws.Cells(r, 11).Value = 1 Then
            If ws.Cells(r, 2).Value = idObjet Then
                QuantiteStock = CLng(Val(CStr(ws.Cells(r, 5).Value)))
                Exit Function
            End If
        End If
    Next r
End Function

' Quantité actuellement sortie (non rendue) pour un objet
Private Function QuantiteSortie(ByVal idObjet As Long) As Long
    Dim ws As Worksheet, r As Long, der As Long, total As Long
    Set ws = ThisWorkbook.Worksheets(F_PRETS)
    der = DerniereLignePrets()
    For r = 2 To der
        If Trim$(CStr(ws.Cells(r, P_RETOUR).Value)) = "" Then
            If ws.Cells(r, P_ID).Value = idObjet Then
                total = total + CLng(Val(CStr(ws.Cells(r, P_QTE).Value)))
            End If
        End If
    Next r
    QuantiteSortie = total
End Function

Private Function NomFichierPropre(ByVal s As String) As String
    Dim interdits As Variant, i As Long
    interdits = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    s = Trim$(s)
    For i = LBound(interdits) To UBound(interdits)
        s = Replace(s, interdits(i), "_")
    Next i
    NomFichierPropre = Replace(s, " ", "_")
End Function

' Partie numérique d'un numéro de bordereau : "BOD012" donne 12 (et 12 donne 12)
Private Function PartieNumerique(ByVal v As Variant) As Long
    Dim s As String, i As Long, ch As String, chiffres As String
    If IsError(v) Then Exit Function
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then chiffres = chiffres & ch
    Next i
    If chiffres <> "" Then PartieNumerique = CLng(Val(chiffres))
End Function

' Renumérote la colonne N° du bordereau (1, 2, 3...) et renvoie le nombre de lignes
Private Function Renumeroter() As Long
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = ThisWorkbook.Worksheets(F_BORD)
    For i = L_B_DEBUT To L_B_FIN
        If Trim$(CStr(ws.Cells(i, B_ID).Value)) <> "" Then
            n = n + 1
            ws.Cells(i, B_NUMLIG).Value = n
        Else
            ws.Cells(i, B_NUMLIG).ClearContents
        End If
    Next i
    Renumeroter = n
End Function


' =====================================================================
'  PRÉPARATION DU BORDEREAU
' =====================================================================
' Depuis la feuille Inventaire : ajoute l'objet de la ligne sélectionnée au bordereau en cours
Public Sub AjouterAuBordereau()
    Dim wsI As Worksheet, wsB As Worksheet
    Dim r As Long, i As Long, idObjet As Long, lig As Long, premVide As Long
    Dim deja As Long, dispo As Long, n As Long, nb As Long
    Dim rep As String, nom As String, d As Double

    Set wsI = ThisWorkbook.Worksheets(F_INVENTAIRE)
    Set wsB = ThisWorkbook.Worksheets(F_BORD)
    If Not ActiveSheet Is wsI Then
        MsgBox "Placez-vous d'abord sur la feuille Inventaire.", vbExclamation
        Exit Sub
    End If
    r = ActiveCell.Row
    If r <= L_INV_ENTETE Or Not IsNumeric(wsI.Cells(r, 1).Value) Or Trim$(CStr(wsI.Cells(r, 1).Value)) = "" Then
        MsgBox "Cliquez d'abord sur une ligne de l'inventaire.", vbExclamation
        Exit Sub
    End If
    If Trim$(CStr(wsB.Range(CELL_NUM).Value)) <> "" Then
        MsgBox "Le bordereau affiché (n° " & wsB.Range(CELL_NUM).Value & ") est déjà validé." & vbCrLf & _
               "Allez sur la feuille Bordereau et cliquez sur « Nouveau bordereau » pour en commencer un autre.", vbExclamation
        Exit Sub
    End If

    idObjet = CLng(wsI.Cells(r, 1).Value)
    nom = CStr(wsI.Cells(r, 3).Value)

    ' L'objet est-il déjà sur le bordereau ? Sinon, première ligne libre
    For i = L_B_DEBUT To L_B_FIN
        If Trim$(CStr(wsB.Cells(i, B_ID).Value)) = "" Then
            If premVide = 0 Then premVide = i
        ElseIf wsB.Cells(i, B_ID).Value = idObjet Then
            lig = i
        End If
    Next i
    If lig > 0 Then deja = CLng(Val(CStr(wsB.Cells(lig, B_QTE).Value)))

    dispo = QuantiteStock(idObjet) - QuantiteSortie(idObjet) - deja
    If dispo <= 0 Then
        MsgBox "Plus aucun exemplaire disponible pour « " & nom & " »" & vbCrLf & _
               "(tout est déjà sorti ou déjà sur ce bordereau).", vbExclamation
        Exit Sub
    End If

    rep = InputBox("Quantité à sortir pour « " & nom & " » ?" & vbCrLf & vbCrLf & _
                   "Disponible : " & dispo, "Ajouter au bordereau", "1")
    If Trim$(rep) = "" Then Exit Sub
    If Not IsNumeric(rep) Then
        MsgBox "Quantité invalide.", vbExclamation
        Exit Sub
    End If
    d = CDbl(rep)
    If d < 1 Or d <> Fix(d) Or d > dispo Then
        MsgBox "La quantité doit être un entier entre 1 et " & dispo & ".", vbExclamation
        Exit Sub
    End If
    n = CLng(d)

    If lig > 0 Then
        wsB.Cells(lig, B_QTE).Value = deja + n
    Else
        If premVide = 0 Then
            MsgBox "Le bordereau est plein (30 lignes). Validez-le puis commencez-en un nouveau.", vbExclamation
            Exit Sub
        End If
        wsB.Cells(premVide, B_ID).Value = idObjet
        wsB.Cells(premVide, B_TYPE).Value = wsI.Cells(r, 2).Value
        wsB.Cells(premVide, B_NOM).Value = nom
        wsB.Cells(premVide, B_QTE).Value = n
        wsB.Cells(premVide, B_ETAT).Value = wsI.Cells(r, 5).Value
        wsB.Cells(premVide, B_OBS).Value = wsI.Cells(r, 7).Value
    End If

    nb = Renumeroter()
    MsgBox "Ajouté : " & n & " x " & nom & "." & vbCrLf & "Le bordereau contient " & nb & " ligne(s).", vbInformation
End Sub

' Sur la feuille Bordereau : retire la ligne sélectionnée et remonte les suivantes
Public Sub RetirerLigneBordereau()
    Dim ws As Worksheet, r As Long, i As Long, c As Long
    Set ws = ThisWorkbook.Worksheets(F_BORD)
    If Not ActiveSheet Is ws Then Exit Sub
    If Trim$(CStr(ws.Range(CELL_NUM).Value)) <> "" Then
        MsgBox "Ce bordereau est déjà validé : il ne peut plus être modifié.", vbExclamation
        Exit Sub
    End If
    r = ActiveCell.Row
    If r < L_B_DEBUT Or r > L_B_FIN Then
        MsgBox "Cliquez d'abord sur une ligne du tableau.", vbExclamation
        Exit Sub
    End If
    For i = r To L_B_FIN - 1
        For c = B_NOM To B_TYPE
            ws.Cells(i, c).Value = ws.Cells(i + 1, c).Value
        Next c
    Next i
    ws.Range(ws.Cells(L_B_FIN, B_NUMLIG), ws.Cells(L_B_FIN, B_TYPE)).ClearContents
    Renumeroter
End Sub

' Vide le bordereau (le nom du groupe est conservé)
Public Sub NouveauBordereau()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(F_BORD)
    If Trim$(CStr(ws.Range(CELL_NUM).Value)) = "" And Trim$(CStr(ws.Cells(L_B_DEBUT, B_ID).Value)) <> "" Then
        If MsgBox("Le bordereau en cours n'a pas été validé. L'effacer quand même ?", _
                  vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If
    ws.Range("C4:E7").ClearContents
    ws.Range(ws.Cells(L_B_DEBUT, B_NUMLIG), ws.Cells(L_B_FIN, B_TYPE)).ClearContents
    ws.Range(CELL_DATE).Value = Date
End Sub


' =====================================================================
'  VALIDATION : numérotation, journal des sorties, PDF
' =====================================================================
Public Sub ValiderBordereau()
    Dim wsB As Worksheet, wsP As Worksheet
    Dim i As Long, r As Long, nb As Long, num As Long, idObjet As Long, dispo As Long
    Dim responsable As String, dossier As String, fichier As String, code As String, d As Double

    Set wsB = ThisWorkbook.Worksheets(F_BORD)
    Set wsP = ThisWorkbook.Worksheets(F_PRETS)

    If Trim$(CStr(wsB.Range(CELL_NUM).Value)) <> "" Then
        MsgBox "Ce bordereau est déjà validé (n° " & wsB.Range(CELL_NUM).Value & ")." & vbCrLf & _
               "Cliquez sur « Nouveau bordereau » pour en préparer un autre.", vbExclamation
        Exit Sub
    End If
    responsable = Trim$(CStr(wsB.Range(CELL_RESP).Value))
    If responsable = "" Then
        MsgBox "Indiquez le nom du responsable.", vbExclamation
        Exit Sub
    End If

    ' --- Contrôles ligne par ligne
    For i = L_B_DEBUT To L_B_FIN
        If Trim$(CStr(wsB.Cells(i, B_ID).Value)) <> "" Then
            nb = nb + 1
            If Not IsNumeric(wsB.Cells(i, B_ID).Value) Or Not IsNumeric(wsB.Cells(i, B_QTE).Value) Then
                MsgBox "Ligne " & i & " : ID ou quantité invalide.", vbExclamation
                Exit Sub
            End If
            d = CDbl(wsB.Cells(i, B_QTE).Value)
            If d < 1 Or d <> Fix(d) Then
                MsgBox "« " & wsB.Cells(i, B_NOM).Value & " » : la quantité doit être un entier d'au moins 1.", vbExclamation
                Exit Sub
            End If
            idObjet = CLng(wsB.Cells(i, B_ID).Value)
            dispo = QuantiteStock(idObjet) - QuantiteSortie(idObjet)
            If d > dispo Then
                MsgBox "« " & wsB.Cells(i, B_NOM).Value & " » : " & d & " demandé(s) mais seulement " & dispo & " disponible(s).", vbExclamation
                Exit Sub
            End If
            If Trim$(CStr(wsB.Cells(i, B_ETAT).Value)) = "" Then
                MsgBox "« " & wsB.Cells(i, B_NOM).Value & " » : indiquez l'état.", vbExclamation
                Exit Sub
            End If
        End If
    Next i
    If nb = 0 Then
        MsgBox "Le bordereau est vide. Ajoutez des objets depuis la feuille Inventaire.", vbExclamation
        Exit Sub
    End If

    ' --- Numéro et date
    r = DerniereLignePrets()
    num = 0
    For i = 2 To r
        If PartieNumerique(wsP.Cells(i, P_NUM).Value) > num Then num = PartieNumerique(wsP.Cells(i, P_NUM).Value)
    Next i
    num = num + 1
    code = PREFIXE_NUM & Format$(num, "000")
    If Not IsDate(wsB.Range(CELL_DATE).Value) Then wsB.Range(CELL_DATE).Value = Date
    wsB.Range(CELL_NUM).Value = code

    ' --- Écriture dans le journal des sorties
    For i = L_B_DEBUT To L_B_FIN
        If Trim$(CStr(wsB.Cells(i, B_ID).Value)) <> "" Then
            r = r + 1
            wsP.Cells(r, P_NUM).Value = code
            wsP.Cells(r, P_DATE).Value = wsB.Range(CELL_DATE).Value
            wsP.Cells(r, P_RESP).Value = responsable
            wsP.Cells(r, P_GROUPE).Value = wsB.Range(CELL_GROUPE).Value
            wsP.Cells(r, P_DEST).Value = wsB.Range(CELL_DEST).Value
            wsP.Cells(r, P_ID).Value = CLng(wsB.Cells(i, B_ID).Value)
            wsP.Cells(r, P_TYPE).Value = wsB.Cells(i, B_TYPE).Value
            wsP.Cells(r, P_NOM).Value = wsB.Cells(i, B_NOM).Value
            wsP.Cells(r, P_QTE).Value = CLng(wsB.Cells(i, B_QTE).Value)
            wsP.Cells(r, P_ETAT).Value = wsB.Cells(i, B_ETAT).Value
            wsP.Cells(r, P_OBS).Value = wsB.Cells(i, B_OBS).Value
        End If
    Next i

    ' --- PDF
    dossier = ThisWorkbook.Path
    If dossier = "" Or LCase$(Left$(dossier, 4)) = "http" Then dossier = Environ$("USERPROFILE") & "\Documents"
    fichier = dossier & Application.PathSeparator & "Bordereau_" & code & "_" & _
              NomFichierPropre(responsable) & ".pdf"
    On Error GoTo ErreurPDF
    wsB.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fichier, Quality:=xlQualityStandard, _
                            IgnorePrintAreas:=False, OpenAfterPublish:=True
    On Error GoTo 0
    MsgBox "Bordereau n° " & code & " enregistré (" & nb & " ligne(s))." & vbCrLf & vbCrLf & _
           "PDF créé : " & fichier, vbInformation
    Exit Sub

ErreurPDF:
    MsgBox "Bordereau n° " & code & " enregistré dans le journal des sorties," & vbCrLf & _
           "mais le PDF n'a pas pu être créé (" & Err.Description & ")." & vbCrLf & vbCrLf & _
           "Vous pouvez imprimer la feuille Bordereau avec Fichier > Imprimer.", vbExclamation
End Sub


' =====================================================================
'  RETOUR DU MATÉRIEL
' =====================================================================
' Sur la feuille Prets : enregistre le retour de la ligne sélectionnée ou de tout son bordereau
Public Sub EnregistrerRetour()
    Dim ws As Worksheet, r As Long, i As Long, der As Long, autres As Long
    Dim choix As VbMsgBoxResult, etat As String, num As String

    Set ws = ThisWorkbook.Worksheets(F_PRETS)
    If Not ActiveSheet Is ws Then
        MsgBox "Placez-vous d'abord sur la feuille Prets.", vbExclamation
        Exit Sub
    End If
    r = ActiveCell.Row
    If r < 2 Or Trim$(CStr(ws.Cells(r, P_NUM).Value)) = "" Then
        MsgBox "Cliquez d'abord sur une ligne de sortie.", vbExclamation
        Exit Sub
    End If
    If Trim$(CStr(ws.Cells(r, P_RETOUR).Value)) <> "" Then
        MsgBox "Cette ligne est déjà marquée comme rendue.", vbInformation
        Exit Sub
    End If

    num = Trim$(CStr(ws.Cells(r, P_NUM).Value))
    der = DerniereLignePrets()
    For i = 2 To der
        If i <> r And Trim$(CStr(ws.Cells(i, P_NUM).Value)) = num And Trim$(CStr(ws.Cells(i, P_RETOUR).Value)) = "" Then
            autres = autres + 1
        End If
    Next i

    choix = vbNo
    If autres > 0 Then
        choix = MsgBox("Le bordereau n° " & num & " contient " & (autres + 1) & " ligne(s) non rendue(s)." & vbCrLf & vbCrLf & _
                       "Oui : tout est rendu, dans l'état de départ." & vbCrLf & _
                       "Non : enregistrer seulement « " & ws.Cells(r, P_NOM).Value & " ».", _
                       vbQuestion + vbYesNoCancel, "Retour du matériel")
        If choix = vbCancel Then Exit Sub
    End If

    If choix = vbYes Then
        For i = 2 To der
            If Trim$(CStr(ws.Cells(i, P_NUM).Value)) = num And Trim$(CStr(ws.Cells(i, P_RETOUR).Value)) = "" Then
                ws.Cells(i, P_RETOUR).Value = Date
                ws.Cells(i, P_ETATRET).Value = ws.Cells(i, P_ETAT).Value
            End If
        Next i
        MsgBox "Bordereau n° " & num & " : tout le matériel est noté comme rendu.", vbInformation
    Else
        etat = InputBox("État au retour de « " & ws.Cells(r, P_NOM).Value & " » ?" & vbCrLf & _
                        "(Bon, À réparer, Hors service...)", "Retour du matériel", CStr(ws.Cells(r, P_ETAT).Value))
        If Trim$(etat) = "" Then Exit Sub
        ws.Cells(r, P_RETOUR).Value = Date
        ws.Cells(r, P_ETATRET).Value = Trim$(etat)
        If LCase$(Trim$(etat)) <> LCase$(Trim$(CStr(ws.Cells(r, P_ETAT).Value))) Then
            MsgBox "Retour enregistré." & vbCrLf & vbCrLf & _
                   "L'état a changé : pensez à mettre à jour l'objet dans l'inventaire" & vbCrLf & _
                   "(feuille Inventaire, bouton « Modifier la ligne »).", vbInformation
        End If
    End If
End Sub
