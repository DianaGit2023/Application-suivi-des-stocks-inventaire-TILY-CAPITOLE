Attribute VB_Name = "InventaireScout"
Option Explicit

' =====================================================================
'  INVENTAIRE SCOUT - gestion de stock historisée
'  Chaque modification ferme l'ancienne ligne (Date_Fin = aujourd'hui,
'  Actif = 0) et crée une nouvelle ligne (Date_Fin = 31/12/9999, Actif = 1).
'  Rien n'est jamais effacé de la feuille Data.
' =====================================================================

' --- Noms des feuilles
Private Const F_ACCUEIL As String = "Accueil"
Private Const F_INVENTAIRE As String = "Inventaire"
Private Const F_DATA As String = "Data"
Private Const F_PREIMPORT As String = "PreImport"

' --- Colonnes de la feuille Data
Private Const C_IDLIGNE As Long = 1
Private Const C_IDOBJET As Long = 2
Private Const C_TYPE As Long = 3
Private Const C_NOM As Long = 4
Private Const C_QTE As Long = 5
Private Const C_ETAT As Long = 6
Private Const C_EMPL As Long = 7
Private Const C_COMM As Long = 8
Private Const C_DEBUT As Long = 9
Private Const C_FIN As Long = 10
Private Const C_ACTIF As Long = 11
Private Const C_PAR As Long = 12

' --- Feuille Accueil : formulaire (colonne C) et zone de notes (colonne E)
Private Const L_FORM_ID As Long = 10
Private Const L_FORM_TYPE As Long = 11
Private Const L_FORM_NOM As Long = 12
Private Const L_FORM_QTE As Long = 13
Private Const L_FORM_ETAT As Long = 14
Private Const L_FORM_EMPL As Long = 15
Private Const L_FORM_COMM As Long = 16
Private Const COL_FORM As Long = 3
Private Const COL_NOTES As Long = 5
Private Const L_NOTES_DEBUT As Long = 11
Private Const L_NOTES_FIN As Long = 60

' --- Feuille Inventaire
Private Const L_INV_ENTETE As Long = 5
Private Const CELL_RECHERCHE As String = "C2"

' --- Feuille PreImport
Private Const L_PRE_ENTETE As Long = 5
Private Const C_PRE_STATUT As Long = 7


' =====================================================================
'  INSTALLATION (à lancer une seule fois après l'import du module)
' =====================================================================
Public Sub InstallerApplication()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(F_ACCUEIL)
    AjouterBouton ws, "btnVoir", "Voir l'inventaire", "AllerInventaire", ws.Range("B7"), 0
    AjouterBouton ws, "btnImport", "Importer un fichier", "ImporterFichier", ws.Range("B7"), 1
    AjouterBouton ws, "btnEnregistrer", "Enregistrer", "EnregistrerSaisie", ws.Range("B18"), 0
    AjouterBouton ws, "btnVider", "Vider le formulaire", "ViderFormulaire", ws.Range("B18"), 1
    AjouterBouton ws, "btnNotes", "Analyser mes notes", "AnalyserNotes", ws.Range("E7"), 0

    Set ws = ThisWorkbook.Worksheets(F_INVENTAIRE)
    AjouterBouton ws, "btnActualiser", "Actualiser / chercher", "ActualiserInventaire", ws.Range("A3"), 0
    AjouterBouton ws, "btnModifier", "Modifier la ligne", "ModifierLigne", ws.Range("A3"), 1
    AjouterBouton ws, "btnSupprimer", "Supprimer la ligne", "SupprimerLigne", ws.Range("A3"), 2
    AjouterBouton ws, "btnHisto", "Historique de l'objet", "VoirHistorique", ws.Range("A3"), 3
    AjouterBouton ws, "btnAccueilInv", "Retour à l'accueil", "AllerAccueil", ws.Range("A3"), 4

    Set ws = ThisWorkbook.Worksheets(F_PREIMPORT)
    AjouterBouton ws, "btnControler", "Contrôler", "LancerControle", ws.Range("A3"), 0
    AjouterBouton ws, "btnIntegrer", "Valider l'intégration", "IntegrerPreImport", ws.Range("A3"), 1
    AjouterBouton ws, "btnViderPre", "Vider le pré-import", "ViderPreImport", ws.Range("A3"), 2
    AjouterBouton ws, "btnAccueilPre", "Retour à l'accueil", "AllerAccueil", ws.Range("A3"), 3

    ActualiserInventaire
    ThisWorkbook.Worksheets(F_ACCUEIL).Activate
    MsgBox "Installation terminée." & vbCrLf & vbCrLf & _
           "Enregistrez maintenant le classeur au format" & vbCrLf & _
           "« Classeur Excel prenant en charge les macros (*.xlsm) ».", vbInformation, "Inventaire scout"
End Sub

Private Sub AjouterBouton(ws As Worksheet, ByVal nom As String, ByVal texte As String, _
                          ByVal macro As String, ancre As Range, ByVal rang As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nom).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ancre.Left + rang * 148, ancre.Top + 4, 140, 26)
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


' =====================================================================
'  NAVIGATION
' =====================================================================
Public Sub AllerInventaire()
    ActualiserInventaire
    ThisWorkbook.Worksheets(F_INVENTAIRE).Activate
End Sub

Public Sub AllerAccueil()
    ThisWorkbook.Worksheets(F_ACCUEIL).Activate
End Sub


' =====================================================================
'  OUTILS
' =====================================================================
Private Function WsData() As Worksheet
    Set WsData = ThisWorkbook.Worksheets(F_DATA)
End Function

Private Function DerniereLigne(ws As Worksheet, ByVal colDebut As Long, ByVal colFin As Long, _
                               ByVal ligneEntete As Long) As Long
    Dim c As Long, r As Long, maxi As Long
    maxi = ligneEntete
    For c = colDebut To colFin
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > maxi Then maxi = r
    Next c
    DerniereLigne = maxi
End Function

Private Function ProchainID(ByVal col As Long) As Long
    Dim ws As Worksheet, der As Long
    Set ws = WsData
    der = DerniereLigne(ws, C_IDLIGNE, C_IDLIGNE, 1)
    If der < 2 Then
        ProchainID = 1
    Else
        ProchainID = CLng(Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, col), ws.Cells(der, col)))) + 1
    End If
End Function

' Clé de comparaison : minuscules, espaces nettoyés, "s" final ignoré (tente = Tentes)
Private Function Cle(ByVal s As String) As String
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If Len(s) > 3 And Right$(s, 1) = "s" Then s = Left$(s, Len(s) - 1)
    Cle = s
End Function

Private Function SansAccent(ByVal s As String) As String
    Dim avec As String, sans As String, i As Long
    avec = "éèêëàâäîïôöùûüç"
    sans = "eeeeaaaiioouuuc"
    s = LCase$(s)
    For i = 1 To Len(avec)
        s = Replace(s, Mid$(avec, i, 1), Mid$(sans, i, 1))
    Next i
    SansAccent = s
End Function

Private Function Majuscule(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) > 0 Then s = UCase$(Left$(s, 1)) & Mid$(s, 2)
    Majuscule = s
End Function

' Vrai si v est un entier positif ou nul ; la valeur est renvoyée dans q
Private Function QuantiteValide(ByVal v As Variant, ByRef q As Long) As Boolean
    Dim d As Double
    QuantiteValide = False
    If IsError(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    d = CDbl(v)
    If d < 0 Or d <> Fix(d) Or d > 1000000 Then Exit Function
    q = CLng(d)
    QuantiteValide = True
End Function


' =====================================================================
'  COEUR : CRUD HISTORISÉ
' =====================================================================
' Read : ligne active d'un objet (0 si introuvable)
Private Function TrouverLigneActive(ByVal idObjet As Long) As Long
    Dim ws As Worksheet, r As Long, der As Long
    Set ws = WsData
    der = DerniereLigne(ws, C_IDLIGNE, C_IDLIGNE, 1)
    For r = 2 To der
        If ws.Cells(r, C_ACTIF).Value = 1 Then
            If ws.Cells(r, C_IDOBJET).Value = idObjet Then
                TrouverLigneActive = r
                Exit Function
            End If
        End If
    Next r
End Function

' Read : ligne active portant ce nom (et ce type s'il est renseigné), 0 sinon
Private Function TrouverActifParNom(ByVal typeObj As String, ByVal nom As String) As Long
    Dim ws As Worksheet, r As Long, der As Long, k As String, kt As String
    Set ws = WsData
    k = Cle(nom)
    kt = Cle(typeObj)
    If k = "" Then Exit Function
    der = DerniereLigne(ws, C_IDLIGNE, C_IDLIGNE, 1)
    For r = 2 To der
        If ws.Cells(r, C_ACTIF).Value = 1 Then
            If Cle(CStr(ws.Cells(r, C_NOM).Value)) = k Then
                If kt = "" Or Cle(CStr(ws.Cells(r, C_TYPE).Value)) = kt Then
                    TrouverActifParNom = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Sub EcrireLigne(ByVal idObjet As Long, ByVal typeObj As String, ByVal nom As String, _
                        ByVal qte As Long, ByVal etat As String, ByVal empl As String, ByVal comm As String)
    Dim ws As Worksheet, r As Long, idLigne As Long
    Set ws = WsData
    idLigne = ProchainID(C_IDLIGNE)
    r = DerniereLigne(ws, C_IDLIGNE, C_IDLIGNE, 1) + 1
    ws.Cells(r, C_IDLIGNE).Value = idLigne
    ws.Cells(r, C_IDOBJET).Value = idObjet
    ws.Cells(r, C_TYPE).Value = typeObj
    ws.Cells(r, C_NOM).Value = nom
    ws.Cells(r, C_QTE).Value = qte
    ws.Cells(r, C_ETAT).Value = etat
    ws.Cells(r, C_EMPL).Value = empl
    ws.Cells(r, C_COMM).Value = comm
    ws.Cells(r, C_DEBUT).Value = Date
    ws.Cells(r, C_FIN).Value = DateSerial(9999, 12, 31)
    ws.Cells(r, C_DEBUT).NumberFormat = "dd/mm/yyyy"
    ws.Cells(r, C_FIN).NumberFormat = "dd/mm/yyyy"
    ws.Cells(r, C_ACTIF).Value = 1
    ws.Cells(r, C_PAR).Value = Application.UserName
End Sub

Private Sub CloturerLigne(ByVal r As Long)
    With WsData
        .Cells(r, C_FIN).Value = Date
        .Cells(r, C_FIN).NumberFormat = "dd/mm/yyyy"
        .Cells(r, C_ACTIF).Value = 0
    End With
End Sub

' Create : renvoie l'ID du nouvel objet
Private Function CreerObjet(ByVal typeObj As String, ByVal nom As String, ByVal qte As Long, _
                            ByVal etat As String, ByVal empl As String, ByVal comm As String) As Long
    Dim idObjet As Long
    idObjet = ProchainID(C_IDOBJET)
    EcrireLigne idObjet, typeObj, nom, qte, etat, empl, comm
    CreerObjet = idObjet
End Function

' Update : ferme l'ancienne version et en crée une nouvelle.
' Renvoie Faux si rien n'a changé (aucune ligne créée) ou si l'objet est introuvable.
Private Function MettreAJourObjet(ByVal idObjet As Long, ByVal typeObj As String, ByVal nom As String, _
                                  ByVal qte As Long, ByVal etat As String, ByVal empl As String, _
                                  ByVal comm As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = WsData
    r = TrouverLigneActive(idObjet)
    If r = 0 Then Exit Function
    If CStr(ws.Cells(r, C_TYPE).Value) = typeObj And CStr(ws.Cells(r, C_NOM).Value) = nom _
       And CStr(ws.Cells(r, C_QTE).Value) = CStr(qte) And CStr(ws.Cells(r, C_ETAT).Value) = etat _
       And CStr(ws.Cells(r, C_EMPL).Value) = empl And CStr(ws.Cells(r, C_COMM).Value) = comm Then
        Exit Function
    End If
    CloturerLigne r
    EcrireLigne idObjet, typeObj, nom, qte, etat, empl, comm
    MettreAJourObjet = True
End Function

' Delete : suppression logique (la ligne est fermée, pas effacée)
Private Function SupprimerObjet(ByVal idObjet As Long) As Boolean
    Dim r As Long
    r = TrouverLigneActive(idObjet)
    If r = 0 Then Exit Function
    CloturerLigne r
    SupprimerObjet = True
End Function


' =====================================================================
'  ACCUEIL : FORMULAIRE DE SAISIE
' =====================================================================
Public Sub EnregistrerSaisie()
    Dim ws As Worksheet
    Dim idTxt As String, typeObj As String, nom As String, etat As String, empl As String, comm As String
    Dim qte As Long, r As Long, idObjet As Long

    Set ws = ThisWorkbook.Worksheets(F_ACCUEIL)
    idTxt = Trim$(CStr(ws.Cells(L_FORM_ID, COL_FORM).Value))
    typeObj = Trim$(CStr(ws.Cells(L_FORM_TYPE, COL_FORM).Value))
    nom = Majuscule(CStr(ws.Cells(L_FORM_NOM, COL_FORM).Value))
    etat = Trim$(CStr(ws.Cells(L_FORM_ETAT, COL_FORM).Value))
    empl = Trim$(CStr(ws.Cells(L_FORM_EMPL, COL_FORM).Value))
    comm = Trim$(CStr(ws.Cells(L_FORM_COMM, COL_FORM).Value))

    If typeObj = "" Or nom = "" Then
        MsgBox "Le type et le nom de l'objet sont obligatoires.", vbExclamation, "Saisie incomplète"
        Exit Sub
    End If
    If Not QuantiteValide(ws.Cells(L_FORM_QTE, COL_FORM).Value, qte) Then
        MsgBox "La quantité doit être un nombre entier positif ou nul.", vbExclamation, "Quantité invalide"
        Exit Sub
    End If
    If etat = "" Then etat = "Bon"

    If idTxt <> "" Then
        ' --- Modification d'un objet existant
        If Not IsNumeric(idTxt) Then
            MsgBox "L'ID objet n'est pas valide. Videz le formulaire et recommencez.", vbExclamation
            Exit Sub
        End If
        idObjet = CLng(idTxt)
        If TrouverLigneActive(idObjet) = 0 Then
            MsgBox "L'objet n° " & idObjet & " n'existe pas ou a été supprimé.", vbExclamation
            Exit Sub
        End If
        If MettreAJourObjet(idObjet, typeObj, nom, qte, etat, empl, comm) Then
            MsgBox "Objet n° " & idObjet & " mis à jour (l'ancienne version est conservée dans l'historique).", vbInformation
        Else
            MsgBox "Aucun changement détecté : rien n'a été modifié.", vbInformation
        End If
    Else
        ' --- Création, avec détection de doublon
        r = TrouverActifParNom(typeObj, nom)
        If r > 0 Then
            If MsgBox("« " & nom & " » existe déjà dans l'inventaire (quantité actuelle : " & _
                      WsData.Cells(r, C_QTE).Value & ")." & vbCrLf & vbCrLf & _
                      "Voulez-vous mettre à jour cet objet avec les valeurs saisies ?", _
                      vbQuestion + vbYesNo, "Objet déjà présent") <> vbYes Then Exit Sub
            idObjet = CLng(WsData.Cells(r, C_IDOBJET).Value)
            If MettreAJourObjet(idObjet, typeObj, nom, qte, etat, empl, comm) Then
                MsgBox "Objet n° " & idObjet & " mis à jour.", vbInformation
            Else
                MsgBox "Aucun changement détecté : rien n'a été modifié.", vbInformation
            End If
        Else
            idObjet = CreerObjet(typeObj, nom, qte, etat, empl, comm)
            MsgBox "Objet ajouté à l'inventaire (n° " & idObjet & ").", vbInformation
        End If
    End If

    ViderFormulaire
    ActualiserInventaire
End Sub

Public Sub ViderFormulaire()
    With ThisWorkbook.Worksheets(F_ACCUEIL)
        .Range(.Cells(L_FORM_ID, COL_FORM), .Cells(L_FORM_COMM, COL_FORM)).ClearContents
    End With
End Sub


' =====================================================================
'  INVENTAIRE : VUE DES LIGNES ACTIVES
' =====================================================================
Public Sub ActualiserInventaire()
    Dim wsI As Worksheet, wsD As Worksheet
    Dim r As Long, n As Long, derD As Long, derI As Long
    Dim filtre As String, texte As String

    Set wsI = ThisWorkbook.Worksheets(F_INVENTAIRE)
    Set wsD = WsData
    Application.ScreenUpdating = False

    If wsI.AutoFilterMode Then wsI.AutoFilterMode = False
    derI = DerniereLigne(wsI, 1, 9, L_INV_ENTETE)
    If derI > L_INV_ENTETE Then
        wsI.Range(wsI.Cells(L_INV_ENTETE + 1, 1), wsI.Cells(derI, 9)).ClearContents
    End If

    filtre = LCase$(Trim$(CStr(wsI.Range(CELL_RECHERCHE).Value)))
    derD = DerniereLigne(wsD, C_IDLIGNE, C_IDLIGNE, 1)
    n = L_INV_ENTETE
    For r = 2 To derD
        If wsD.Cells(r, C_ACTIF).Value = 1 Then
            texte = LCase$(wsD.Cells(r, C_TYPE).Value & " " & wsD.Cells(r, C_NOM).Value & " " & _
                           wsD.Cells(r, C_ETAT).Value & " " & wsD.Cells(r, C_EMPL).Value & " " & _
                           wsD.Cells(r, C_COMM).Value)
            If filtre = "" Or InStr(texte, filtre) > 0 Then
                n = n + 1
                wsI.Cells(n, 1).Value = wsD.Cells(r, C_IDOBJET).Value
                wsI.Cells(n, 2).Value = wsD.Cells(r, C_TYPE).Value
                wsI.Cells(n, 3).Value = wsD.Cells(r, C_NOM).Value
                wsI.Cells(n, 4).Value = wsD.Cells(r, C_QTE).Value
                wsI.Cells(n, 5).Value = wsD.Cells(r, C_ETAT).Value
                wsI.Cells(n, 6).Value = wsD.Cells(r, C_EMPL).Value
                wsI.Cells(n, 7).Value = wsD.Cells(r, C_COMM).Value
                wsI.Cells(n, 8).Value = wsD.Cells(r, C_DEBUT).Value
                wsI.Cells(n, 8).NumberFormat = "dd/mm/yyyy"
                wsI.Cells(n, 9).Value = wsD.Cells(r, C_PAR).Value
            End If
        End If
    Next r

    If n > L_INV_ENTETE + 1 Then
        wsI.Range(wsI.Cells(L_INV_ENTETE, 1), wsI.Cells(n, 9)).Sort _
            Key1:=wsI.Cells(L_INV_ENTETE, 2), Order1:=xlAscending, _
            Key2:=wsI.Cells(L_INV_ENTETE, 3), Order2:=xlAscending, Header:=xlYes
    End If
    Application.ScreenUpdating = True
End Sub

' ID de l'objet sur la ligne sélectionnée de la feuille Inventaire (0 si aucune)
Private Function ObjetSelectionne() As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(F_INVENTAIRE)
    If Not ActiveSheet Is ws Then
        MsgBox "Placez-vous d'abord sur la feuille Inventaire.", vbExclamation
        Exit Function
    End If
    r = ActiveCell.Row
    If r <= L_INV_ENTETE Or Not IsNumeric(ws.Cells(r, 1).Value) Or Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then
        MsgBox "Cliquez d'abord sur une ligne de l'inventaire.", vbExclamation
        Exit Function
    End If
    ObjetSelectionne = CLng(ws.Cells(r, 1).Value)
End Function

' Charge la ligne sélectionnée dans le formulaire de l'accueil
Public Sub ModifierLigne()
    Dim idObjet As Long, r As Long, wsA As Worksheet, wsD As Worksheet
    idObjet = ObjetSelectionne()
    If idObjet = 0 Then Exit Sub
    r = TrouverLigneActive(idObjet)
    If r = 0 Then
        MsgBox "Objet introuvable. Actualisez l'inventaire.", vbExclamation
        Exit Sub
    End If
    Set wsA = ThisWorkbook.Worksheets(F_ACCUEIL)
    Set wsD = WsData
    wsA.Cells(L_FORM_ID, COL_FORM).Value = idObjet
    wsA.Cells(L_FORM_TYPE, COL_FORM).Value = wsD.Cells(r, C_TYPE).Value
    wsA.Cells(L_FORM_NOM, COL_FORM).Value = wsD.Cells(r, C_NOM).Value
    wsA.Cells(L_FORM_QTE, COL_FORM).Value = wsD.Cells(r, C_QTE).Value
    wsA.Cells(L_FORM_ETAT, COL_FORM).Value = wsD.Cells(r, C_ETAT).Value
    wsA.Cells(L_FORM_EMPL, COL_FORM).Value = wsD.Cells(r, C_EMPL).Value
    wsA.Cells(L_FORM_COMM, COL_FORM).Value = wsD.Cells(r, C_COMM).Value
    wsA.Activate
    wsA.Cells(L_FORM_QTE, COL_FORM).Select
End Sub

Public Sub SupprimerLigne()
    Dim idObjet As Long, r As Long
    idObjet = ObjetSelectionne()
    If idObjet = 0 Then Exit Sub
    r = TrouverLigneActive(idObjet)
    If r = 0 Then
        MsgBox "Objet introuvable. Actualisez l'inventaire.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Retirer « " & WsData.Cells(r, C_NOM).Value & " » de l'inventaire ?" & vbCrLf & vbCrLf & _
              "La ligne reste consultable dans l'historique (feuille Data).", _
              vbQuestion + vbYesNo, "Confirmer la suppression") <> vbYes Then Exit Sub
    SupprimerObjet idObjet
    ActualiserInventaire
End Sub

Public Sub VoirHistorique()
    Dim idObjet As Long, ws As Worksheet, r As Long, der As Long, n As Long
    Dim msg As String, fin As String
    idObjet = ObjetSelectionne()
    If idObjet = 0 Then Exit Sub
    Set ws = WsData
    der = DerniereLigne(ws, C_IDLIGNE, C_IDLIGNE, 1)
    For r = der To 2 Step -1
        If ws.Cells(r, C_IDOBJET).Value = idObjet Then
            n = n + 1
            If n <= 10 Then
                If ws.Cells(r, C_ACTIF).Value = 1 Then
                    fin = "en cours"
                Else
                    fin = "au " & Format$(ws.Cells(r, C_FIN).Value, "dd/mm/yyyy")
                End If
                msg = msg & "Du " & Format$(ws.Cells(r, C_DEBUT).Value, "dd/mm/yyyy") & " " & fin & " : " & _
                      ws.Cells(r, C_QTE).Value & " x " & Left$(CStr(ws.Cells(r, C_NOM).Value), 25) & _
                      " (" & ws.Cells(r, C_ETAT).Value & ")" & vbCrLf
            End If
        End If
    Next r
    If n > 10 Then msg = msg & "... et " & (n - 10) & " version(s) plus ancienne(s) dans la feuille Data."
    MsgBox msg, vbInformation, "Historique de l'objet n° " & idObjet & " (du plus récent au plus ancien)"
End Sub


' =====================================================================
'  PRÉ-IMPORT : sas de contrôle avant intégration
' =====================================================================
Private Sub AjouterPreImport(ByVal typeObj As String, ByVal nom As String, ByVal qte As Variant, _
                             ByVal etat As String, ByVal empl As String, ByVal comm As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(F_PREIMPORT)
    r = DerniereLigne(ws, 1, 6, L_PRE_ENTETE) + 1
    ws.Cells(r, 1).Value = Majuscule(typeObj)
    ws.Cells(r, 2).Value = Majuscule(nom)
    ws.Cells(r, 3).Value = qte
    ws.Cells(r, 4).Value = Majuscule(etat)
    ws.Cells(r, 5).Value = Trim$(empl)
    ws.Cells(r, 6).Value = Trim$(comm)
End Sub

Public Sub LancerControle()
    Dim nbErr As Long
    nbErr = ControlerPreImport()
    If nbErr > 0 Then
        MsgBox nbErr & " ligne(s) en erreur (en rouge). Corrigez-les puis relancez le contrôle.", vbExclamation
    Else
        MsgBox "Contrôle terminé : aucune erreur.", vbInformation
    End If
End Sub

' Remplit la colonne Contrôle et renvoie le nombre de lignes en erreur
Private Function ControlerPreImport() As Long
    Dim ws As Worksheet, r As Long, der As Long, rD As Long
    Dim typeObj As String, nom As String, statut As String, q As Long, nbErr As Long

    Set ws = ThisWorkbook.Worksheets(F_PREIMPORT)
    der = DerniereLigne(ws, 1, 6, L_PRE_ENTETE)
    For r = L_PRE_ENTETE + 1 To der
        typeObj = Trim$(CStr(ws.Cells(r, 1).Value))
        nom = Trim$(CStr(ws.Cells(r, 2).Value))

        ' Type manquant : on le reprend d'un objet existant du même nom
        If typeObj = "" And nom <> "" Then
            rD = TrouverActifParNom("", nom)
            If rD > 0 Then
                typeObj = CStr(WsData.Cells(rD, C_TYPE).Value)
                ws.Cells(r, 1).Value = typeObj
            End If
        End If

        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 6))) = 0 Then
            statut = ""
        ElseIf nom = "" Then
            statut = "ERREUR : nom manquant"
        ElseIf Not QuantiteValide(ws.Cells(r, 3).Value, q) Then
            statut = "ERREUR : quantité manquante ou invalide"
        ElseIf typeObj = "" Then
            statut = "ERREUR : type à compléter"
        Else
            rD = TrouverActifParNom(typeObj, nom)
            If rD > 0 Then
                statut = "MAJ : existe déjà (quantité actuelle " & WsData.Cells(rD, C_QTE).Value & ")"
            Else
                statut = "OK : nouvel objet"
            End If
        End If

        ws.Cells(r, C_PRE_STATUT).Value = statut
        If Left$(statut, 6) = "ERREUR" Then
            ws.Cells(r, C_PRE_STATUT).Font.Color = RGB(192, 0, 0)
            nbErr = nbErr + 1
        ElseIf Left$(statut, 3) = "MAJ" Then
            ws.Cells(r, C_PRE_STATUT).Font.Color = RGB(197, 90, 17)
        Else
            ws.Cells(r, C_PRE_STATUT).Font.Color = RGB(0, 128, 0)
        End If
    Next r
    ControlerPreImport = nbErr
End Function

Public Sub IntegrerPreImport()
    Dim ws As Worksheet, wsD As Worksheet, r As Long, der As Long, rD As Long
    Dim typeObj As String, nom As String, etat As String, empl As String, comm As String
    Dim q As Long, nbErr As Long, nbNew As Long, nbMaj As Long, nbId As Long

    Set ws = ThisWorkbook.Worksheets(F_PREIMPORT)
    Set wsD = WsData
    der = DerniereLigne(ws, 1, 6, L_PRE_ENTETE)
    If der <= L_PRE_ENTETE Then
        MsgBox "Le pré-import est vide.", vbInformation
        Exit Sub
    End If

    nbErr = ControlerPreImport()
    If nbErr > 0 Then
        MsgBox nbErr & " ligne(s) en erreur (en rouge)." & vbCrLf & _
               "Corrigez-les ou supprimez-les avant de valider : rien n'a été intégré.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Intégrer ces lignes dans l'inventaire ?" & vbCrLf & vbCrLf & _
              "Pour les objets déjà présents, la quantité importée REMPLACE la quantité actuelle.", _
              vbQuestion + vbYesNo, "Valider l'intégration") <> vbYes Then Exit Sub

    If Not Sauvegarder() Then
        If MsgBox("La copie de sauvegarde n'a pas pu être créée. Continuer quand même ?", _
                  vbExclamation + vbYesNo) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = L_PRE_ENTETE + 1 To der
        nom = Trim$(CStr(ws.Cells(r, 2).Value))
        If nom <> "" Then
            typeObj = Trim$(CStr(ws.Cells(r, 1).Value))
            etat = Trim$(CStr(ws.Cells(r, 4).Value))
            empl = Trim$(CStr(ws.Cells(r, 5).Value))
            comm = Trim$(CStr(ws.Cells(r, 6).Value))
            QuantiteValide ws.Cells(r, 3).Value, q

            rD = TrouverActifParNom(typeObj, nom)
            If rD > 0 Then
                ' Champs laissés vides dans l'import : on garde les valeurs actuelles
                If etat = "" Then etat = CStr(wsD.Cells(rD, C_ETAT).Value)
                If empl = "" Then empl = CStr(wsD.Cells(rD, C_EMPL).Value)
                If comm = "" Then comm = CStr(wsD.Cells(rD, C_COMM).Value)
                If MettreAJourObjet(CLng(wsD.Cells(rD, C_IDOBJET).Value), CStr(wsD.Cells(rD, C_TYPE).Value), _
                                    CStr(wsD.Cells(rD, C_NOM).Value), q, etat, empl, comm) Then
                    nbMaj = nbMaj + 1
                Else
                    nbId = nbId + 1
                End If
            Else
                If etat = "" Then etat = "Bon"
                CreerObjet typeObj, nom, q, etat, empl, comm
                nbNew = nbNew + 1
            End If
        End If
    Next r

    ws.Range(ws.Cells(L_PRE_ENTETE + 1, 1), ws.Cells(der, C_PRE_STATUT)).ClearContents
    ActualiserInventaire
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(F_INVENTAIRE).Activate
    MsgBox "Intégration terminée :" & vbCrLf & _
           " - " & nbNew & " nouvel(s) objet(s)" & vbCrLf & _
           " - " & nbMaj & " objet(s) mis à jour" & vbCrLf & _
           " - " & nbId & " objet(s) inchangé(s)", vbInformation
End Sub

Public Sub ViderPreImport()
    Dim ws As Worksheet, der As Long
    Set ws = ThisWorkbook.Worksheets(F_PREIMPORT)
    der = DerniereLigne(ws, 1, C_PRE_STATUT, L_PRE_ENTETE)
    If der > L_PRE_ENTETE Then
        ws.Range(ws.Cells(L_PRE_ENTETE + 1, 1), ws.Cells(der, C_PRE_STATUT)).ClearContents
    End If
End Sub

' Copie de sécurité du classeur avant toute intégration en masse
Private Function Sauvegarder() As Boolean
    Dim dossier As String
    On Error GoTo Echec
    dossier = ThisWorkbook.Path
    If dossier = "" Or LCase$(Left$(dossier, 4)) = "http" Then
        dossier = Environ$("USERPROFILE") & "\Documents"
    End If
    ThisWorkbook.SaveCopyAs dossier & Application.PathSeparator & "Sauvegarde_" & _
                            Format$(Now, "yyyymmdd_hhnnss") & "_" & ThisWorkbook.Name
    Sauvegarder = True
    Exit Function
Echec:
    Sauvegarder = False
End Function


' =====================================================================
'  NOTES DU TÉLÉPHONE : texte libre vers pré-import
' =====================================================================
' Comprend : "3 tentes", "3x tentes", "tente x3", "Tentes : 3", "tentes 3",
'            "Type;Nom;Quantité;État;Emplacement;Commentaire"
'            et les titres de rubrique : une ligne "Cuisine :" donne le type des lignes suivantes.
' Renvoie Vrai si une ligne a été ajoutée au pré-import.
Private Function AnalyserLigneNote(ByVal s As String, ByRef typeCourant As String) As Boolean
    Dim re As Object, m As Object, motifs As Variant, i As Long
    Dim nom As String, qte As Variant, champs() As String, c As String

    s = Trim$(Replace(s, vbTab, " "))
    ' Puces en début de ligne
    Do While Len(s) > 0
        c = Left$(s, 1)
        If c = "-" Or c = "*" Or c = ChrW$(8226) Or c = ChrW$(8211) Or c = " " Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    s = Trim$(s)
    If s = "" Then Exit Function

    ' Ligne déjà structurée
    If InStr(s, ";") > 0 Then
        champs = Split(s & ";;;;;", ";")
        AjouterPreImport champs(0), champs(1), Trim$(champs(2)), champs(3), champs(4), champs(5)
        AnalyserLigneNote = True
        Exit Function
    End If

    ' Titre de rubrique
    If Right$(s, 1) = ":" Then
        typeCourant = Majuscule(Left$(s, Len(s) - 1))
        Exit Function
    End If

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    motifs = Array("^(\d+)\s*[x*]?\s+(.+)$", "^(.+?)\s+[x*]\s*(\d+)$", "^(.+?)\s*[:=\-]\s*(\d+)$", "^(.+?)\s+(\d+)$")
    nom = s
    qte = ""
    For i = 0 To 3
        re.Pattern = motifs(i)
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            If i = 0 Then
                qte = CLng(m.SubMatches(0))
                nom = m.SubMatches(1)
            Else
                nom = m.SubMatches(0)
                qte = CLng(m.SubMatches(1))
            End If
            Exit For
        End If
    Next i

    AjouterPreImport typeCourant, nom, qte, "", "", ""
    AnalyserLigneNote = True
End Function

Public Sub AnalyserNotes()
    Dim ws As Worksheet, r As Long, n As Long, typeCourant As String, s As String
    Set ws = ThisWorkbook.Worksheets(F_ACCUEIL)
    Application.ScreenUpdating = False
    For r = L_NOTES_DEBUT To L_NOTES_FIN
        s = CStr(ws.Cells(r, COL_NOTES).Value)
        If Trim$(s) <> "" Then
            If AnalyserLigneNote(s, typeCourant) Then n = n + 1
        End If
    Next r
    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "Aucune note à analyser. Collez vos notes dans la zone jaune (une ligne par objet).", vbInformation
        Exit Sub
    End If
    ws.Range(ws.Cells(L_NOTES_DEBUT, COL_NOTES), ws.Cells(L_NOTES_FIN, COL_NOTES)).ClearContents
    ControlerPreImport
    ThisWorkbook.Worksheets(F_PREIMPORT).Activate
    MsgBox n & " ligne(s) envoyée(s) dans le pré-import." & vbCrLf & _
           "Vérifiez-les, corrigez les lignes en rouge, puis cliquez sur « Valider l'intégration ».", vbInformation
End Sub


' =====================================================================
'  IMPORT D'UN FICHIER (CSV ou texte)
' =====================================================================
Private Function LireAvecJeu(ByVal chemin As String, ByVal jeu As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = jeu
    stm.Open
    stm.LoadFromFile chemin
    LireAvecJeu = stm.ReadText
    stm.Close
End Function

' Lit le fichier en UTF-8 ; s'il n'est pas en UTF-8 (accents illisibles), relit en Windows-1252
Private Function LireFichierTexte(ByVal chemin As String) As String
    Dim txt As String
    txt = LireAvecJeu(chemin, "utf-8")
    If InStr(txt, ChrW$(&HFFFD)) > 0 Then txt = LireAvecJeu(chemin, "windows-1252")
    If Left$(txt, 1) = ChrW$(&HFEFF) Then txt = Mid$(txt, 2)
    LireFichierTexte = txt
End Function

Private Function DetecterSeparateur(ByVal ligne As String) As String
    Dim nPv As Long, nTab As Long, nVirg As Long
    nPv = Len(ligne) - Len(Replace(ligne, ";", ""))
    nTab = Len(ligne) - Len(Replace(ligne, vbTab, ""))
    nVirg = Len(ligne) - Len(Replace(ligne, ",", ""))
    If nPv > 0 And nPv >= nTab And nPv >= nVirg Then
        DetecterSeparateur = ";"
    ElseIf nTab > 0 And nTab >= nVirg Then
        DetecterSeparateur = vbTab
    ElseIf nVirg > 0 Then
        DetecterSeparateur = ","
    Else
        DetecterSeparateur = ""
    End If
End Function

' Position (0 à 5) du champ correspondant à un titre de colonne, -1 si inconnu
Private Function ChampDepuisTitre(ByVal titre As String) As Long
    Dim t As String
    t = SansAccent(Trim$(Replace(titre, """", "")))
    ChampDepuisTitre = -1
    If Left$(t, 4) = "type" Or Left$(t, 5) = "categ" Then
        ChampDepuisTitre = 0
    ElseIf Left$(t, 3) = "nom" And Left$(t, 6) <> "nombre" Then
        ChampDepuisTitre = 1
    ElseIf Left$(t, 5) = "objet" Or Left$(t, 5) = "desig" Or Left$(t, 5) = "libel" Then
        ChampDepuisTitre = 1
    ElseIf Left$(t, 5) = "quant" Or Left$(t, 3) = "qte" Or Left$(t, 3) = "qty" Or Left$(t, 6) = "nombre" Then
        ChampDepuisTitre = 2
    ElseIf Left$(t, 4) = "etat" Then
        ChampDepuisTitre = 3
    ElseIf Left$(t, 5) = "empla" Or Left$(t, 4) = "lieu" Then
        ChampDepuisTitre = 4
    ElseIf Left$(t, 5) = "comme" Or Left$(t, 5) = "remar" Or Left$(t, 4) = "note" Then
        ChampDepuisTitre = 5
    End If
End Function

Public Sub ImporterFichier()
    Dim chemin As Variant, txt As String, lignes() As String, champs() As String
    Dim i As Long, j As Long, k As Long, premiere As Long, n As Long
    Dim sep As String, typeCourant As String
    Dim pos(0 To 5) As Long, valeurs(0 To 5) As String, aEntete As Boolean

    chemin = Application.GetOpenFilename("Fichiers CSV ou texte (*.csv;*.txt),*.csv;*.txt", , _
                                         "Choisir le fichier d'inventaire à importer")
    If VarType(chemin) = vbBoolean Then Exit Sub

    On Error GoTo ErreurLecture
    txt = LireFichierTexte(CStr(chemin))
    On Error GoTo 0

    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    lignes = Split(txt, vbLf)

    ' Première ligne non vide
    premiere = -1
    For i = LBound(lignes) To UBound(lignes)
        If Trim$(lignes(i)) <> "" Then
            premiere = i
            Exit For
        End If
    Next i
    If premiere = -1 Then
        MsgBox "Le fichier est vide.", vbExclamation
        Exit Sub
    End If

    sep = DetecterSeparateur(lignes(premiere))
    Application.ScreenUpdating = False

    If sep = "" Then
        ' Pas de colonnes : le fichier est traité comme des notes libres
        For i = premiere To UBound(lignes)
            If AnalyserLigneNote(lignes(i), typeCourant) Then n = n + 1
        Next i
    Else
        ' Ordre par défaut : Type ; Nom ; Quantité ; État ; Emplacement ; Commentaire
        For j = 0 To 5
            pos(j) = j
        Next j
        ' Si la première ligne contient des titres, on s'en sert pour repérer les colonnes
        champs = Split(lignes(premiere), sep)
        For j = LBound(champs) To UBound(champs)
            k = ChampDepuisTitre(champs(j))
            If k = 1 Or k = 2 Then aEntete = True
        Next j
        If aEntete Then
            For j = 0 To 5
                pos(j) = -1
            Next j
            For j = LBound(champs) To UBound(champs)
                k = ChampDepuisTitre(champs(j))
                If k >= 0 Then
                    If pos(k) = -1 Then pos(k) = j
                End If
            Next j
            premiere = premiere + 1
        End If

        For i = premiere To UBound(lignes)
            If Trim$(Replace(lignes(i), sep, "")) <> "" Then
                champs = Split(lignes(i), sep)
                For j = 0 To 5
                    valeurs(j) = ""
                    If pos(j) >= 0 And pos(j) <= UBound(champs) Then
                        valeurs(j) = Trim$(Replace(champs(pos(j)), """", ""))
                    End If
                Next j
                AjouterPreImport valeurs(0), valeurs(1), valeurs(2), valeurs(3), valeurs(4), valeurs(5)
                n = n + 1
            End If
        Next i
    End If

    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "Aucune ligne exploitable dans ce fichier.", vbExclamation
        Exit Sub
    End If
    ControlerPreImport
    ThisWorkbook.Worksheets(F_PREIMPORT).Activate
    MsgBox n & " ligne(s) lue(s) et placée(s) dans le pré-import." & vbCrLf & _
           "Vérifiez-les, corrigez les lignes en rouge, puis cliquez sur « Valider l'intégration ».", vbInformation
    Exit Sub

ErreurLecture:
    Application.ScreenUpdating = True
    MsgBox "Impossible de lire le fichier : " & Err.Description, vbExclamation
End Sub
