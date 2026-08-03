'###########################################################
'#  MODULE STANDARD : DIAG_3464
'#
'#  Module de diagnostic temporaire, a supprimer une fois le
'#  probleme localise.
'#
'#  INSTALLATION
'#    Alt+F11 -> Insertion -> Module
'#    Coller ce code, enregistrer sous le nom  DIAG_3464
'#
'#  UTILISATION
'#    Ctrl+G pour ouvrir la fenetre Execution, puis taper :
'#
'#        TesterRequetesPilotage
'#        ScannerObjetsApresMigration
'#
'#    Les resultats s'affichent dans la fenetre Execution.
'###########################################################

Option Explicit
Option Compare Database


'===========================================================
' TEST 1 - Executer une par une les requetes de 1_INITIAL
'-----------------------------------------------------------
' Identifie laquelle echoue, et avec quel numero d'erreur.
'===========================================================
Public Sub TesterRequetesPilotage()

    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print String(60, "=")
    Debug.Print "TEST DES REQUETES DE 1_INITIAL"
    Debug.Print String(60, "=")

    TesterUneRequete db, "Comptage par etat", _
        "SELECT ETAT, Count(*) AS Nombre FROM [DATABASE] " & _
        "WHERE ETAT IN ('A prendre en charge', 'A l etude', 'En suspens', " & _
        "'Transmis au signataire DRC', 'Creation de la trame', 'A notifier', " & _
        "'En cours de verification') GROUP BY ETAT"

    TesterUneRequete db, "Moyenne DELAIS ANALYSTE", _
        "SELECT Avg([DELAIS ANALYSTE]) AS M FROM [DATABASE] WHERE [DELAIS ANALYSTE] Is Not Null;"

    TesterUneRequete db, "Moyenne DELAIS SIGNATAIRE", _
        "SELECT Avg([DELAIS SIGNATAIRE]) AS M FROM [DATABASE] WHERE [DELAIS SIGNATAIRE] Is Not Null;"

    TesterUneRequete db, "Moyenne DELAIS HORS DRC", _
        "SELECT Avg([DELAIS HORS DRC]) AS M FROM [DATABASE] WHERE [DELAIS HORS DRC] Is Not Null;"

    TesterUneRequete db, "Moyenne DELAIS NOTIFICATION", _
        "SELECT Avg([DELAIS NOTIFICATION]) AS M FROM [DATABASE] WHERE [DELAIS NOTIFICATION] Is Not Null;"

    Debug.Print String(60, "-")
    Debug.Print "TYPES REELS DES CHAMPS"
    Debug.Print String(60, "-")

    AfficherTypeChamp db, "ETAT"
    AfficherTypeChamp db, "DELAIS ANALYSTE"
    AfficherTypeChamp db, "DELAIS SIGNATAIRE"
    AfficherTypeChamp db, "DELAIS HORS DRC"
    AfficherTypeChamp db, "DELAIS NOTIFICATION"
    AfficherTypeChamp db, "DELAIS DRC"
    AfficherTypeChamp db, "DELAIS AD"
    AfficherTypeChamp db, "FAST"
    AfficherTypeChamp db, "CONSULTATION CPLE Recommandé"
    AfficherTypeChamp db, "ALERTE DEFAUT"
    AfficherTypeChamp db, "VALIDATION AINDEP"
    AfficherTypeChamp db, "URGENCE"
    AfficherTypeChamp db, "TYPE_LIGNE"
    AfficherTypeChamp db, "TYPE_GARANTIE"
    AfficherTypeChamp db, "IDSG"

    Set db = Nothing
    Debug.Print String(60, "=")
    Debug.Print "FIN DU TEST 1"
End Sub


Private Sub TesterUneRequete(db As DAO.Database, libelle As String, sql As String)
    Dim rs As DAO.Recordset
    On Error GoTo Echec

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        Debug.Print "[ OK  ] " & libelle & "  (aucune ligne)"
    Else
        Debug.Print "[ OK  ] " & libelle & "  -> " & Nz(rs(0), "Null")
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub

Echec:
    Debug.Print "[ KO  ] " & libelle & "  -> erreur " & Err.Number & " : " & Err.Description
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
    End If
End Sub


Private Sub AfficherTypeChamp(db As DAO.Database, fieldName As String)
    On Error GoTo Absent
    Dim t As Integer
    t = db.TableDefs("DATABASE").Fields(fieldName).Type
    Debug.Print "  [" & fieldName & "] -> " & NomTypeDAO(t)
    Exit Sub
Absent:
    Debug.Print "  [" & fieldName & "] -> CHAMP INTROUVABLE"
End Sub


Private Function NomTypeDAO(t As Integer) As String
    Select Case t
        Case dbBoolean:    NomTypeDAO = "Oui/Non"
        Case dbByte:       NomTypeDAO = "Numerique (Octet)"
        Case dbInteger:    NomTypeDAO = "Numerique (Entier)"
        Case dbLong:       NomTypeDAO = "Numerique (Entier long)"
        Case dbCurrency:   NomTypeDAO = "Monetaire"
        Case dbSingle:     NomTypeDAO = "Numerique (Reel simple)"
        Case dbDouble:     NomTypeDAO = "Numerique (Reel double)"
        Case dbDate:       NomTypeDAO = "Date/Heure"
        Case dbText:       NomTypeDAO = "TEXTE"
        Case dbMemo:       NomTypeDAO = "Memo / Texte long"
        Case dbDecimal:    NomTypeDAO = "Numerique (Decimal)"
        Case Else:         NomTypeDAO = "Autre (code " & t & ")"
    End Select
End Function


'===========================================================
' TEST 2 - Traquer les comparaisons textuelles residuelles
'-----------------------------------------------------------
' Apres la migration Texte -> Oui/Non, toute requete qui
' compare encore un de ces champs a "Oui" ou a une chaine vide
' leve une erreur 3464. Cette procedure parcourt :
'   - les requetes enregistrees
'   - la source, le filtre et le tri de chaque formulaire
'   - la source de chaque controle (dont les graphiques)
'   - la source et le filtre de chaque etat
'===========================================================
Public Sub ScannerObjetsApresMigration()

    Dim champs As Variant
    champs = Array("FAST", "CONSULTATION CPLE Recommandé", "ALERTE DEFAUT", _
                   "VALIDATION AINDEP", "URGENCE")

    Debug.Print String(60, "=")
    Debug.Print "SCAN DES OBJETS REFERENCANT LES CHAMPS MIGRES"
    Debug.Print String(60, "=")

    ScannerRequetes champs
    ScannerFormulaires champs
    ScannerEtats champs

    Debug.Print String(60, "=")
    Debug.Print "FIN DU SCAN"
    Debug.Print "Examiner chaque ligne signalee : un champ Oui/Non ne"
    Debug.Print "doit plus etre compare a ""Oui"", a """" ou a Null,"
    Debug.Print "mais a True / False."
End Sub


Private Sub ScannerRequetes(champs As Variant)
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim i As Integer

    Set db = CurrentDb()
    Debug.Print "--- REQUETES ENREGISTREES ---"

    For Each qd In db.QueryDefs
        If Left(qd.Name, 1) <> "~" Then          ' ignorer les requetes systeme
            For i = LBound(champs) To UBound(champs)
                If InStr(1, qd.sql, champs(i), vbTextCompare) > 0 Then
                    Debug.Print "  Requete [" & qd.Name & "] reference [" & champs(i) & "]"
                    Debug.Print "     " & Left(Replace(qd.sql, vbCrLf, " "), 200)
                    Exit For
                End If
            Next i
        End If
    Next qd

    Set qd = Nothing
    Set db = Nothing
End Sub


Private Sub ScannerFormulaires(champs As Variant)
    Dim ao As AccessObject
    Dim frm As Form
    Dim ctl As Control
    Dim i As Integer
    Dim etaitOuvert As Boolean

    Debug.Print "--- FORMULAIRES ---"

    For Each ao In CurrentProject.AllForms
        etaitOuvert = ao.IsLoaded
        On Error Resume Next
        If Not etaitOuvert Then DoCmd.OpenForm ao.Name, acDesign, , , , acHidden
        Set frm = Forms(ao.Name)

        If Not frm Is Nothing Then
            For i = LBound(champs) To UBound(champs)
                VerifierChaine "Form [" & ao.Name & "] RecordSource", frm.RecordSource, champs(i)
                VerifierChaine "Form [" & ao.Name & "] Filter", frm.Filter, champs(i)
                VerifierChaine "Form [" & ao.Name & "] OrderBy", frm.OrderBy, champs(i)
            Next i

            For Each ctl In frm.Controls
                For i = LBound(champs) To UBound(champs)
                    VerifierChaine "Form [" & ao.Name & "] ctrl [" & ctl.Name & "] ControlSource", _
                                   CStr(ctl.Properties("ControlSource")), champs(i)
                    VerifierChaine "Form [" & ao.Name & "] ctrl [" & ctl.Name & "] RowSource", _
                                   CStr(ctl.Properties("RowSource")), champs(i)
                Next i
            Next ctl
        End If

        If Not etaitOuvert Then DoCmd.Close acForm, ao.Name, acSaveNo
        Set frm = Nothing
        On Error GoTo 0
    Next ao
End Sub


Private Sub ScannerEtats(champs As Variant)
    Dim ao As AccessObject
    Dim rpt As Report
    Dim i As Integer

    Debug.Print "--- ETATS ---"

    For Each ao In CurrentProject.AllReports
        On Error Resume Next
        DoCmd.OpenReport ao.Name, acDesign, , , acHidden
        Set rpt = Reports(ao.Name)
        If Not rpt Is Nothing Then
            For i = LBound(champs) To UBound(champs)
                VerifierChaine "Etat [" & ao.Name & "] RecordSource", rpt.RecordSource, champs(i)
                VerifierChaine "Etat [" & ao.Name & "] Filter", rpt.Filter, champs(i)
            Next i
        End If
        DoCmd.Close acReport, ao.Name, acSaveNo
        Set rpt = Nothing
        On Error GoTo 0
    Next ao
End Sub


Private Sub VerifierChaine(contexte As String, contenu As String, champ As String)
    If Len(contenu) = 0 Then Exit Sub
    If InStr(1, contenu, champ, vbTextCompare) > 0 Then
        Debug.Print "  " & contexte & " reference [" & champ & "]"
        Debug.Print "     " & Left(Replace(contenu, vbCrLf, " "), 200)
    End If
End Sub
