'###########################################################
'#  MODULE DE CLASSE DU FORMULAIRE  5_MODIFICATIONS
'#  Version corrigee - POINT 2 : cases a cocher / champs Oui-Non
'#
'#  PREREQUIS : la migration Texte -> Oui/Non des cinq champs
'#  FAST, CONSULTATION CPLE Recommande, ALERTE DEFAUT,
'#  VALIDATION AINDEP, URGENCE doit avoir ete effectuee AVANT
'#  de coller ce code (voir le fichier MIGRATION_TEXTE_VERS_OUINON).
'#
'#  RESTENT A TRAITER, hors de ce lot :
'#    Point 1 - les _BeforeUpdate ecrivent l'ancienne valeur
'#    Point 3 - CalculerDelaiSansChevauchement (module DELAIS_)
'#    Point 4 - Option Explicit
'#              Ne l'ajouter qu'apres avoir traite le point 1,
'#              la compilation remontera probablement d'autres
'#              variables non declarees.
'###########################################################

Public Sub CalculerEtMettreAJourDelais()
    ' Déclarations des variables (si nécessaire, pour éviter les erreurs de portée)
    Dim IDSGValue As String
    IDSGValue = Nz(Me.Texte162.Value, "")
    If IDSGValue = "" Then Exit Sub
    ' Appeler la fonction de calcul des délais
    Call DELAIS_.DELAIS(IDSGValue)
    ' Mettre à jour les champs de délai
    Call UpdateField("DELAIS ANALYSTE", Me.Texte834)
    Call UpdateField("DELAIS SIGNATAIRE", Me.Texte849)
    Call UpdateField("DELAIS HORS DRC", Me.Texte851)
    Call UpdateField("DELAIS NOTIFICATION", Me.Texte853)
    Call UpdateField("DELAIS DRC", Me.Texte862)
    Call UpdateField("DELAIS AD", Me.Texte958)
    Call UpdateField("DELAIS TOTAL", Me.Texte855)
End Sub
Public Sub Commande744_Click()
    Dim notation As Variant
    Dim valeur As Variant
    Dim Modifiable1 As Variant
    Dim isPME As Boolean
    Dim isValidNotation As Boolean
    notation = Nz(Me.Modifiable5.Value, "")
    Modifiable1 = Nz(Me.Modifiable1.Value, "")
    ' --- Vérification si la Notation est vide ---
    If IsNull(notation) Or Trim(notation) = "" Then
        MsgBox "Veuillez entrer une notation.", vbExclamation
        Exit Sub
    End If
    ' --- Vérification si la Valeur est vide ou n'est pas un nombre ---
    valeur = Me.Texte11.Value
    If IsNull(valeur) Or Trim(valeur) = "" Then
        MsgBox "Veuillez entrer un montant (MXOF).", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(valeur) Then
        MsgBox "Veuillez entrer un montant valide (chiffre) dans Texte11.", vbExclamation
        Exit Sub
    End If
    ' --- Conversion en Long après validation ---
    valeur = CLng(valeur)
    ' --- La valeur doit être supérieure à zéro
    If valeur <= 0 Then
        MsgBox "Le montant doit être supérieur à zéro.", vbExclamation
        Exit Sub
    End If
    ' --- Détermine si modifiable1 est PME ---
    isPME = (UCase$(Trim$(CStr(Modifiable1))) = "PME")
    Select Case UCase$(Trim$(CStr(notation)))
        Case "7", "7-"
            If valeur <= 740 Then
                Me.Modifiable12.Value = "DGA / DARC"
            ElseIf valeur <= 1311 Then
                Me.Modifiable12.Value = "DG / DRC"
            ElseIf valeur < 5900 Then
                Me.Modifiable12.Value = "DG / RISQ CRE"
            ElseIf valeur <= 39000 Then
                Me.Modifiable12.Value = "CORISQ"
            Else
                Me.Modifiable12.Value = "CA"
            End If
        Case "5-"
            If isPME Then
                If valeur < 230 Then
                    Me.Modifiable12.Value = "SC / ARS"
                ElseIf valeur < 250 Then
                    Me.Modifiable12.Value = "SC / SCO"
                ElseIf valeur < 780 Then
                    Me.Modifiable12.Value = "COO COV / SCO"
                ElseIf valeur < 1140 Then
                    Me.Modifiable12.Value = "COO COV / DARC"
                ElseIf valeur < 2290 Then
                    Me.Modifiable12.Value = "DCE / DARC"
                ElseIf valeur < 3440 Then
                    Me.Modifiable12.Value = "DCE / RRC"
                ElseIf valeur < 6559 Then
                    Me.Modifiable12.Value = "DGA / DRC"
                ElseIf valeur < 13500 Then
                    Me.Modifiable12.Value = "DGA / RISQ CRE"
                ElseIf valeur < 52479 Then
                    Me.Modifiable12.Value = "DG / RISQ CRE"
                Else
                    Me.Modifiable12.Value = "CA : CORISQ"
                End If
            Else 'Not PME
                If valeur < 230 Then
                    Me.Modifiable12.Value = "SC / ARS"
                ElseIf valeur < 1140 Then
                    Me.Modifiable12.Value = "SC / SCO"
                ElseIf valeur < 1740 Then
                    Me.Modifiable12.Value = "COO / DARC"
                ElseIf valeur < 2290 Then
                    Me.Modifiable12.Value = "DCE / DARC"
                ElseIf valeur < 3440 Then
                    Me.Modifiable12.Value = "DCE / RRC"
                ElseIf valeur < 6559 Then
                    Me.Modifiable12.Value = "DGA / DRC"
                ElseIf valeur < 13500 Then
                    Me.Modifiable12.Value = "DGA / RISQ CRE"
                ElseIf valeur < 52479 Then
                    Me.Modifiable12.Value = "DG / RISQ CRE"
                Else
                    Me.Modifiable12.Value = "CA / CORISQ"
                End If
            End If
        Case "6+", "6", "6-", "7+"
            If isPME Then
                If valeur < 165 Then
                    Me.Modifiable12.Value = "SC / ARS"
                ElseIf valeur < 250 Then
                    Me.Modifiable12.Value = "SC / SCO"
                ElseIf valeur < 780 Then
                    Me.Modifiable12.Value = "COO COV / SCO"
                ElseIf valeur < 820 Then
                    Me.Modifiable12.Value = "COO COV / DARC"
                ElseIf valeur < 1640 Then
                    Me.Modifiable12.Value = "DCE / DARC"
                ElseIf valeur < 2460 Then
                    Me.Modifiable12.Value = "DCE / RRC"
                ElseIf valeur < 4591 Then
                    Me.Modifiable12.Value = "DGA / DRC"
                ElseIf valeur < 13500 Then
                    Me.Modifiable12.Value = "DGA / RISQ CRE"
                ElseIf valeur < 39357 Then
                    Me.Modifiable12.Value = "DG / RISQ CRE"
                Else
                    Me.Modifiable12.Value = "CA / CORISQ"
                End If
            Else 'Not PME
                If valeur < 165 Then
                    Me.Modifiable12.Value = "SC / ARS"
                ElseIf valeur < 820 Then
                    Me.Modifiable12.Value = "SC / SCO"
                ElseIf valeur < 1640 Then
                    Me.Modifiable12.Value = "DCE / DARC"
                ElseIf valeur < 1700 Then
                    Me.Modifiable12.Value = "DCE / COO"
                ElseIf valeur < 2460 Then
                    Me.Modifiable12.Value = "DCE / RRC"
                ElseIf valeur < 4591 Then
                    Me.Modifiable12.Value = "DGA / DRC"
                ElseIf valeur < 13500 Then
                    Me.Modifiable12.Value = "DGA / RISQ CRE"
                ElseIf valeur < 39357 Then
                    Me.Modifiable12.Value = "DG / RISQ CRE"
                Else
                    Me.Modifiable12.Value = "CA / CORISQ"
                End If
            End If
    End Select
    Call UpdateTableField("LAD", Me.Modifiable12)
End Sub
'UPDATE FIELD ENTRY
Public Sub UpdateField(fieldName As String, targetControl As Control)
    On Error GoTo ErrorHandler
    Dim IDValue As String
    Dim LookupValue As Variant
    IDValue = Nz(Me.Texte162.Value, "")
    If Len(IDValue) > 0 Then
        LookupValue = Nz(DLookup("[" & fieldName & "]", "DATABASE", "IDSG='" & Replace(IDValue, "'", "''") & "'"), "")
        ' Vérifier si LookupValue est Null
        If IsNull(LookupValue) Then
            targetControl.Value = "" ' Définir sur une chaîne vide si c'est Null
        Else
            targetControl.Value = LookupValue
        End If
    Else
        targetControl.Value = ""
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Erreur N°: " & Err.number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbExclamation
End Sub
'-----------------------------------------------------------
' ECRITURE : case a cocher -> champ Oui/Non
'-----------------------------------------------------------
' CORRIGE : ecrivait 'Oui' / '' (du texte) dans le champ, sans
' dbFailOnError ni gestionnaire. L'UPDATE echouait en silence.
' Le Nz(fieldValue, False) corrige en outre le test
' "If fieldValue = True" qui renvoyait Null sur une case en
' etat indetermine et tombait sans bruit dans le Else.
Public Sub UpdateDatabaseField(fieldName As String, fieldValue As Variant)
    On Error GoTo ErrorHandler
    Dim IDValue As String
    Dim sql As String

    IDValue = Nz(Me.Texte162.Value, "")
    If Len(IDValue) = 0 Then Exit Sub

    sql = "UPDATE [DATABASE] SET [" & fieldName & "] = " & _
          IIf(Nz(fieldValue, False), "True", "False") & _
          " WHERE [IDSG] = '" & Replace(IDValue, "'", "''") & "'"

    CurrentDb.Execute sql, dbFailOnError
    Exit Sub

ErrorHandler:
    MsgBox "Erreur N° " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "SQL : " & sql, vbCritical, "UpdateDatabaseField"
End Sub

'-----------------------------------------------------------
' LECTURE : champ Oui/Non -> case a cocher   (NOUVEAU)
'-----------------------------------------------------------
Private Function LireBooleen(fieldName As String) As Boolean
    Dim IDValue As String

    IDValue = Nz(Me.Texte162.Value, "")
    If Len(IDValue) = 0 Then Exit Function   ' renvoie False

    LireBooleen = Nz(DLookup("[" & fieldName & "]", "[DATABASE]", _
                     "[IDSG]='" & Replace(IDValue, "'", "''") & "'"), False)
End Function
'-----------------------------------------------------------
' TYPE_LIGNE (champ Texte, liste concatenee) - durci
'-----------------------------------------------------------
Private Sub SaveTypeLigne()
    On Error GoTo ErrorHandler
    Dim valeur As String
    Dim sql As String
    Dim IDValue As String

    IDValue = Nz(Me.Texte162.Value, "")
    If Len(IDValue) = 0 Then Exit Sub

    valeur = ""
    If Nz(Me.ChkCaisse, False) Then valeur = valeur & "Caisse, "
    If Nz(Me.ChkSignature, False) Then valeur = valeur & "Signature, "
    If Nz(Me.ChkMLT, False) Then valeur = valeur & "Moyen/Long terme, "
    ' Supprimer la derniere virgule
    If Len(valeur) > 0 Then valeur = Left(valeur, Len(valeur) - 2)

    sql = "UPDATE [DATABASE] SET [TYPE_LIGNE] = '" & valeur & "'" & _
          " WHERE [IDSG] = '" & Replace(IDValue, "'", "''") & "'"
    CurrentDb.Execute sql, dbFailOnError
    Exit Sub

ErrorHandler:
    MsgBox "Erreur N° " & Err.Number & vbCrLf & Err.Description, vbCritical, "SaveTypeLigne"
End Sub
Public Sub Update_TypeLigne()
    Dim IDValue As String
    Dim valeur As String

    IDValue = Nz(Me.Texte162.Value, "")
    valeur = Nz(DLookup("[TYPE_LIGNE]", "[DATABASE]", _
                "[IDSG]='" & Replace(IDValue, "'", "''") & "'"), "")

    Me.ChkCaisse = (InStr(valeur, "Caisse") > 0)
    Me.ChkSignature = (InStr(valeur, "Signature") > 0)
    Me.ChkMLT = (InStr(valeur, "Moyen/Long terme") > 0)
End Sub
'-----------------------------------------------------------
' TYPE_GARANTIE (champ Texte, liste concatenee)
'-----------------------------------------------------------
' CORRIGE : l'ecriture produisait "AGF, " alors que
' Update_TypeGarantie relit InStr(valeur, "AFG"). Les deux
' lettres etaient inversees, la case ne revenait donc jamais
' cochee a la relecture, quelle que soit la valeur en base.
Private Sub SaveTypeGarantie()
    On Error GoTo ErrorHandler
    Dim valeur As String
    Dim sql As String
    Dim IDValue As String

    IDValue = Nz(Me.Texte162.Value, "")
    If Len(IDValue) = 0 Then Exit Sub

    valeur = ""
    If Nz(Me.ChkARIZ, False) Then valeur = valeur & "ARIZ, "
    If Nz(Me.ChkAFG, False) Then valeur = valeur & "AFG, "      ' <-- AFG, plus AGF
    If Nz(Me.ChkIFC, False) Then valeur = valeur & "IFC, "
    If Len(valeur) > 0 Then valeur = Left(valeur, Len(valeur) - 2)

    sql = "UPDATE [DATABASE] SET [TYPE_GARANTIE] = '" & valeur & "'" & _
          " WHERE [IDSG] = '" & Replace(IDValue, "'", "''") & "'"
    CurrentDb.Execute sql, dbFailOnError
    Exit Sub

ErrorHandler:
    MsgBox "Erreur N° " & Err.Number & vbCrLf & Err.Description, vbCritical, "SaveTypeGarantie"
End Sub
Public Sub Update_TypeGarantie()
    Dim IDValue As String
    Dim valeur As String

    IDValue = Nz(Me.Texte162.Value, "")
    valeur = Nz(DLookup("[TYPE_GARANTIE]", "[DATABASE]", _
                "[IDSG]='" & Replace(IDValue, "'", "''") & "'"), "")

    Me.ChkARIZ = (InStr(valeur, "ARIZ") > 0)
    Me.ChkAFG = (InStr(valeur, "AFG") > 0)
    Me.ChkIFC = (InStr(valeur, "IFC") > 0)
End Sub
Public Sub Update1(): UpdateField "RCT", Me.Texte1
End Sub
Public Sub Update2(): UpdateField "MARCHE", Me.Modifiable1
End Sub
Public Sub Update3(): UpdateField "PCRU", Me.Texte3
End Sub
Public Sub Update4(): UpdateField "CLIENT", Me.Texte4 'Nom
End Sub
Public Sub Update5(): UpdateField "GROUPE", Me.Texte5
End Sub
Public Sub Update6(): UpdateField "SECTEUR", Me.Modifiable2
End Sub
Public Sub Update7(): UpdateField "DEMANDE", Me.Modifiable3
End Sub
Public Sub Update8(): UpdateField "RECEPTION", Me.Texte8
End Sub
Public Sub Update9(): UpdateField "ANALYSTE", Me.Modifiable4
End Sub
Public Sub Update10(): UpdateField "NOTE", Me.Modifiable5
End Sub
Public Sub Update11(): UpdateField "MXOF", Me.Texte11
End Sub
Public Sub Update12(): UpdateField "LAD", Me.Modifiable12
End Sub
Public Sub Update13(): UpdateField "SIGNATAIRE", Me.Modifiable13
End Sub
Public Sub Update14(): UpdateField "ATTRIBUTION", Me.Texte14
End Sub
Public Sub Update15(): UpdateField "PRISE EN CHARGE", Me.Texte15
End Sub
Public Sub Update16(): UpdateField "ETAT", Me.Modifiable16
End Sub
Public Sub Update17(): UpdateField "DVA NOTE", Me.Texte17
End Sub
Public Sub Update18(): UpdateField "A/T", Me.Texte19
End Sub
Public Sub Update19(): UpdateField "ENVOIE 1 MARCHE", Me.Texte20
End Sub
Public Sub Update20(): UpdateField "RETOUR 1 MARCHE", Me.Texte21
End Sub
Public Sub Update21(): UpdateField "ENVOIE 2 MARCHE", Me.Texte22
End Sub
Public Sub Update22(): UpdateField "RETOUR 2 MARCHE", Me.Texte23
End Sub
Public Sub Update23(): UpdateField "AUTRES J MARCHE", Me.Texte24
End Sub
Public Sub Update24(): UpdateField "DATE ENVOIE AU SIGNATAIRE", Me.Texte505
End Sub
Public Sub Update25(): UpdateField "DATE SIGNATURE", Me.Texte25
End Sub
Public Sub Update26(): UpdateField "ENVOIE DCE/COO", Me.Texte26
End Sub
Public Sub Update27(): UpdateField "RETOUR DCE/COO", Me.Texte27
End Sub
Public Sub Update28(): UpdateField "AUTRESJ DCEE/COO", Me.Texte28
End Sub
Public Sub Update29(): UpdateField "ENVOIE DGA", Me.Texte29
End Sub
Public Sub Update30(): UpdateField "RETOUR DGA", Me.Texte30
End Sub
Public Sub Update31(): UpdateField "AUTRESJ DGA", Me.Texte31
End Sub
Public Sub Update32(): UpdateField "ENVOIE DG", Me.Texte32
End Sub
Public Sub Update33(): UpdateField "RETOUR DG", Me.Texte654
End Sub
Public Sub Update34(): UpdateField "AUTRESJ DG", Me.Texte33
End Sub
Public Sub Update35(): UpdateField "ENVOIE RISQCRE", Me.Texte34
End Sub
Public Sub Update36(): UpdateField "RETOUR RISQCRE", Me.Texte35
End Sub
Public Sub Update37(): UpdateField "AUTRESJ RISQCRE", Me.Texte36
End Sub
Public Sub Update38(): UpdateField "ENVOIE CA", Me.Texte37
End Sub
Public Sub Update39(): UpdateField "RETOUR CA", Me.Texte38
End Sub
Public Sub Update40(): UpdateField "AUTRESJ CA", Me.Texte39
End Sub
Public Sub Update41(): UpdateField "ENVOIE AUTRESPCRU", Me.Texte40
End Sub
Public Sub Update42(): UpdateField "RETOUR AUTRESPCRU", Me.Texte41
End Sub
Public Sub Update43(): UpdateField "AUTRESJ AUTRESPCRU", Me.Texte42
End Sub
Public Sub Update44(): UpdateField "DECISION", Me.Modifiable43
End Sub
Public Sub Update45(): UpdateField "MONTANT ACCORDE", Me.Texte44
End Sub
Public Sub Update46(): UpdateField "DATE SIGNATURE FINAL", Me.Texte45
End Sub
Public Sub Update47(): UpdateField "DATE ENVOIE TRAME", Me.Texte46
End Sub
Public Sub Update48(): UpdateField "Date de notification", Me.Texte47
End Sub
Public Sub Update49(): UpdateField "ANALYSTE NOTIFICATION", Me.Modifiable48
End Sub
Public Sub Update50(): UpdateField "ENVOIE CORISQ", Me.Texte815
End Sub
Public Sub Update51(): UpdateField "RETOUR CORISQ", Me.Texte818
End Sub
Public Sub Update52(): UpdateField "AUTRESJ CORISQ", Me.Texte821
End Sub
Public Sub Update53(): UpdateField "DELAIS ANALYSTE", Me.Texte834
End Sub
Public Sub Update54(): UpdateField "DELAIS SIGNATAIRE", Me.Texte849
End Sub
Public Sub Update55(): UpdateField "DELAIS HORS DRC", Me.Texte851
End Sub
Public Sub Update56(): UpdateField "DELAIS NOTIFICATION", Me.Texte853
End Sub
Public Sub Update57(): UpdateField "DELAIS TOTAL", Me.Texte855
End Sub
Public Sub Update58(): UpdateField "DELAIS DRC", Me.Texte862
End Sub
Public Sub Update59(): UpdateField "ENVOIE 1 SIGN MARCHE", Me.Texte897
End Sub
Public Sub Update60(): UpdateField "RETOUR 1 SIGN MARCHE", Me.Texte899
End Sub
Public Sub Update61(): UpdateField "ENVOIE 2 SIGN MARCHE", Me.Texte901
End Sub
Public Sub Update62(): UpdateField "RETOUR 2 SIGN MARCHE", Me.Texte903
End Sub
Public Sub Update63(): UpdateField "AUTRES J SIGN MARCHE", Me.Texte905
End Sub
Public Sub Update64(): UpdateField "DELAIS AD", Me.Texte958
End Sub
' CORRIGE : appelait UpdateDatabaseField, qui est le routeur des
' champs Oui/Non. [DELAIS AD] est un nombre de jours : il doit
' passer par UpdateTableField, qui type la valeur.
Private Sub Texte958_AfterUpdate()
    Call UpdateTableField("DELAIS AD", Me.Texte958)
End Sub
Private Sub Option362_AfterUpdate()
    Call UpdateDatabaseField("CONSULTATION CPLE Recommandé", Me.Option362)
End Sub
Private Sub Option364_AfterUpdate()
    Call UpdateDatabaseField("ALERTE DEFAUT", Me.Option364)
End Sub
Private Sub Option371_AfterUpdate()
    Call UpdateDatabaseField("VALIDATION AINDEP", Me.Option371)
End Sub
Private Sub Option373_AfterUpdate()
    Call UpdateDatabaseField("FAST", Me.Option373)
End Sub
Private Sub Option735_AfterUpdate()
    Call UpdateDatabaseField("URGENCE", Me.Option735)
End Sub
Private Sub ChkCaisse_AfterUpdate()
    Call SaveTypeLigne
End Sub
Private Sub ChkSignature_AfterUpdate()
    Call SaveTypeLigne
End Sub
Private Sub ChkMLT_AfterUpdate()
    Call SaveTypeLigne
End Sub
Private Sub ChkARIZ_AfterUpdate()
    Call SaveTypeGarantie
End Sub
Private Sub ChkAFG_AfterUpdate()
    Call SaveTypeGarantie
End Sub
Private Sub ChkIFC_AfterUpdate()
    Call SaveTypeGarantie
End Sub
'UPDATE CHANGE
Public Sub Update_FAST()
    Me.Option373 = LireBooleen("FAST")
End Sub

Public Sub Update_CPLE()
    Me.Option362 = LireBooleen("CONSULTATION CPLE Recommandé")
End Sub

Public Sub Update_DEFAUT()
    Me.Option364 = LireBooleen("ALERTE DEFAUT")
End Sub

Public Sub Update_VALDIND()
    Me.Option371 = LireBooleen("VALIDATION AINDEP")
End Sub

Public Sub Update_URGENCE()
    Me.Option735 = LireBooleen("URGENCE")
End Sub
' DELAIS ANALYSTE
Public Sub Texte834_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS ANALYSTE", Me.Texte834)
End Sub
Public Sub Texte862_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS DRC", Me.Texte862)
End Sub
' DELAIS SIGNATAIRE", Me.Texte849
Public Sub Texte849_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS SIGNATAIRE", Me.Texte849)
End Sub
' DELAIS HORS DRC", Me.Texte851
Public Sub Texte851_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS HORS DRC", Me.Texte851)
End Sub
' DELAIS NOTIFICATION", Me.Texte853
Public Sub Texte853_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS NOTIFICATION", Me.Texte853)
End Sub
' DELAIS TOTAL", Me.Texte855
Public Sub Texte855_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DELAIS TOTAL", Me.Texte855)
End Sub
' RCT
Public Sub Texte1_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("RCT", Me.Texte1)
End Sub
' MARCHE
Public Sub Modifiable1_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("MARCHE", Me.Modifiable1)
End Sub
' PCRU
Public Sub Texte3_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("PCRU", Me.Texte3)
End Sub
' CLIENT
Public Sub Texte4_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("CLIENT", Me.Texte4)
End Sub
' GROUPE
Public Sub Texte5_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("GROUPE", Me.Texte5)
End Sub
' SECTEUR
Public Sub Modifiable2_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("SECTEUR", Me.Modifiable2)
End Sub
' DEMANDE
Public Sub Modifiable3_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DEMANDE", Me.Modifiable3)
End Sub
' RECEPTION
Public Sub Texte8_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("RECEPTION", Me.Texte8)
            Call CalculerEtMettreAJourDelais
End Sub
' ANALYSTE
Public Sub Modifiable4_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("ANALYSTE", Me.Modifiable4)
End Sub
' NOTE
Public Sub Modifiable5_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("NOTE", Me.Modifiable5)
End Sub
' MXOF
Public Sub Texte11_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("MXOF", Me.Texte11)
End Sub
' LAD
Public Sub Modifiable12_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("LAD", Me.Modifiable12)
End Sub
' SIGNATAIRE
Public Sub Modifiable13_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("SIGNATAIRE", Me.Modifiable13)
End Sub
Public Sub Texte14_BeforeUpdate(Cancel As Integer) ' Attribution
    Dim receptionDate As Variant
    Dim attributionDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    receptionDate = Nz(Me.Texte8.Value, "") ' Reception
    attributionDate = Nz(Me.Texte14.Value, "") ' Attribution
    ' Vérifier si Attribution est rempli mais pas Réception
    If attributionDate <> "" And receptionDate = "" Then
        MsgBox "Attention : La date de réception n'est pas remplie, supprimez cette valeur.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si Attribution est antérieure à Réception
    If IsDate(receptionDate) And IsDate(attributionDate) Then
        If CDate(attributionDate) < CDate(receptionDate) Then
            MsgBox "Attention : La date d'attribution est antérieure à la date de réception, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
        Call UpdateTableField("ATTRIBUTION", Me.Texte14)
        Call CalculerEtMettreAJourDelais
End Sub
' PRISE EN CHARGE
Public Sub Texte15_BeforeUpdate(Cancel As Integer)
    Dim attributionDate As Variant
    Dim priseEnChargeDate As Variant
    Dim message As String
    Dim champVide As Boolean
    champVide = False
    message = ""
    ' Vérification des champs obligatoires
    If IsNull(Me.Texte8) Or Me.Texte8 = "" Then
        message = message & "- Veuillez remplir le champ RECEPTION pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If IsNull(Me.Modifiable4) Or Me.Modifiable4 = "" Then
        message = message & "- Veuillez remplir le champ ANALYSTE pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If IsNull(Me.Modifiable5) Or Me.Modifiable5 = "" Then
        message = message & "- Veuillez remplir le champ NOTE pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If IsNull(Me.Texte11) Or Me.Texte11 = "" Then
        message = message & "- Veuillez remplir le champ MXOF pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If IsNull(Me.Texte3) Or Me.Texte3 = "" Then
        message = message & "- Veuillez remplir le champ PCRU pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If IsNull(Me.Modifiable1) Or Me.Modifiable1 = "" Then
        message = message & "- Veuillez remplir le champ MARCHE pour pouvoir renseigner celui-ci." & vbCrLf
        champVide = True
    End If
    If champVide Then
        MsgBox message, vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    attributionDate = Nz(Me.Texte14.Value, "") ' Attribution
    priseEnChargeDate = Nz(Me.Texte15.Value, "") ' Prise en charge
    ' Vérifier si Prise en charge est remplie mais pas Attribution
    If priseEnChargeDate <> "" And attributionDate = "" Then
        MsgBox "Attention : La date de prise en charge ne peut pas être remplie si la date d'attribution est vide.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si Prise en charge est antérieure à Attribution
    If IsDate(attributionDate) And IsDate(priseEnChargeDate) Then
        If CDate(priseEnChargeDate) < CDate(attributionDate) Then
            MsgBox "Attention : La date de prise en charge est antérieure à la date d'attribution, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("PRISE EN CHARGE", Me.Texte15)
    Call CalculerEtMettreAJourDelais
End Sub
' ETAT
Public Sub Modifiable16_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("ETAT", Me.Modifiable16)
End Sub
' DVA NOTE
Public Sub Texte17_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("DVA NOTE", Me.Texte17)
End Sub
' A/T
Public Sub Texte19_BeforeUpdate(Cancel As Integer)
    Dim dvaNoteDate As Variant
    Dim atDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dvaNoteDate = Nz(Me.Texte17.Value, "") ' DVA NOTE
    atDate = Nz(Me.Texte19.Value, "") ' A/T
    ' Vérifier si A/T est antérieure à DVA NOTE
    If IsDate(dvaNoteDate) And IsDate(atDate) Then
        If CDate(atDate) < CDate(dvaNoteDate) Then
            MsgBox "Attention : La date d'A/T est antérieure à la date de DVA NOTE, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("A/T", Me.Texte19)
End Sub
' ENVOIE 1 MARCHE
Public Sub Texte20_BeforeUpdate(Cancel As Integer)
    Dim priseEnChargeDate As Variant
    Dim envoie1Date As Variant
    Dim retour1Date As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    priseEnChargeDate = Nz(Me.Texte15.Value, "") ' PRISE EN CHARGE
    envoie1Date = Nz(Me.Texte20.Value, "") ' ENVOIE 1 MARCHE
    retour1Date = Nz(Me.Texte21.Value, "") ' RETOUR 1 MARCHE
    ' Vérifier si PRISE EN CHARGE est vide et ENVOIE 1 est en cours de modification
    If priseEnChargeDate = "" And envoie1Date <> "" Then
        MsgBox "Attention : La date d'envoi N°1 ne peut pas être remplie si la date de prise en charge est vide.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si ENVOIE 1 est antérieure à PRISE EN CHARGE
    If IsDate(priseEnChargeDate) And IsDate(envoie1Date) Then
        If CDate(envoie1Date) < CDate(priseEnChargeDate) Then
            MsgBox "Attention : La date d'envoi N°1 est antérieure à la date de prise en charge, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    ' Vérifier si RETOUR 1 est rempli mais pas ENVOIE 1
    If retour1Date <> "" And envoie1Date = "" Then
        MsgBox "Attention : La date de retour N°1 ne peut pas être remplie si la date d'envoi N°1 est vide.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si RETOUR 1 est antérieure à ENVOIE 1
    If IsDate(envoie1Date) And IsDate(retour1Date) Then
        If CDate(retour1Date) < CDate(envoie1Date) Then
            MsgBox "Attention : La date de retour N°1 est antérieure à la date d'envoi N°1, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE 1 MARCHE", Me.Texte20)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR 1 MARCHE
Public Sub Texte21_BeforeUpdate(Cancel As Integer)
    Dim envoie1Date As Variant
    Dim retour1Date As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoie1Date = Nz(Me.Texte20.Value, "") ' ENVOIE 1 MARCHE
    retour1Date = Nz(Me.Texte21.Value, "") ' RETOUR 1 MARCHE
    ' Vérifier si RETOUR 1 est rempli mais pas ENVOIE 1
    If retour1Date <> "" And envoie1Date = "" Then
        MsgBox "Attention : La date de retour N°1 ne peut pas être remplie si la date d'envoi N°1 est vide.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si RETOUR 1 est antérieure à ENVOIE 1
    If IsDate(envoie1Date) And IsDate(retour1Date) Then
        If CDate(retour1Date) < CDate(envoie1Date) Then
            MsgBox "Attention : La date de retour N°1 est antérieure à la date d'envoi N°1, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR 1 MARCHE", Me.Texte21)
    Call CalculerEtMettreAJourDelais
End Sub
Public Sub Texte22_BeforeUpdate(Cancel As Integer)
    Dim envoie1Date As Variant
    Dim envoie2Date As Variant
    Dim retour1Date As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoie1Date = Nz(Me.Texte20.Value, "") ' ENVOIE 1 MARCHE
    envoie2Date = Nz(Me.Texte22.Value, "") ' ENVOIE 2 MARCHE
    retour1Date = Nz(Me.Texte21.Value, "")   ' RETOUR 1 MARCHE
    ' Vérifier si ENVOIE 1 est rempli, RETOUR 1 est vide, et ENVOIE 2 est en cours de modification
    If IsDate(envoie1Date) And Not IsDate(retour1Date) And Not IsNull(Me.Texte22.Value) Then
        MsgBox "Impossible de remplir 'ENVOIE 2 MARCHE' si 'ENVOIE 1 MARCHE' est rempli et 'RETOUR 1 MARCHE' est vide. Veuillez d'abord renseigner 'RETOUR 1 MARCHE'.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si ENVOIE 2 est antérieure à RETOUR 1
    If IsDate(retour1Date) And IsDate(envoie2Date) Then
        If CDate(envoie2Date) < CDate(retour1Date) Then
            MsgBox "Attention : La date d'envoi N°2 est antérieure à la date de retour N°1, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    ' Vérifier si ENVOIE 2 est antérieure à ENVOIE 1
    If IsDate(envoie1Date) And IsDate(envoie2Date) Then
        If CDate(envoie2Date) < CDate(envoie1Date) Then
            MsgBox "Attention : La date d'envoi N°2 est antérieure à la date d'envoi N°1, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE 2 MARCHE", Me.Texte22)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR 2 MARCHE
Public Sub Texte23_BeforeUpdate(Cancel As Integer)
    Dim envoie2Date As Variant
    Dim retour2Date As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoie2Date = Nz(Me.Texte22.Value, "") ' ENVOIE 2 MARCHE
    retour2Date = Nz(Me.Texte23.Value, "") ' RETOUR 2 MARCHE
    ' Vérifier si RETOUR 2 est rempli mais pas ENVOIE 2
    If retour2Date <> "" And envoie2Date = "" Then
        MsgBox "Attention : La date de retour N°2 ne peut pas être remplie si la date d'envoi N°2 est vide.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    ' Vérifier si RETOUR 2 est antérieure à ENVOIE 2
    If IsDate(envoie2Date) And IsDate(retour2Date) Then
        If CDate(retour2Date) < CDate(envoie2Date) Then
            MsgBox "Attention : La date de retour N°2 est antérieure à la date d'envoi N°2, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR 2 MARCHE", Me.Texte23)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRES J MARCHE
Public Sub Texte24_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRES J MARCHE", Me.Texte24)
    Call CalculerEtMettreAJourDelais
End Sub
' DATE ENVOIE AU SIGNATAIRE
Public Sub Texte505_BeforeUpdate(Cancel As Integer)
    Dim priseEnChargeDate As Variant
    Dim envoie1Date As Variant, retour1Date As Variant
    Dim envoie2Date As Variant, retour2Date As Variant
    Dim envoieSignataireDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    priseEnChargeDate = Nz(Me.Texte15.Value, "") ' PRISE EN CHARGE
    envoie1Date = Nz(Me.Texte20.Value, "") ' ENVOIE 1 MARCHE
    retour1Date = Nz(Me.Texte21.Value, "") ' RETOUR 1 MARCHE
    envoie2Date = Nz(Me.Texte22.Value, "") ' ENVOIE 2 MARCHE
    retour2Date = Nz(Me.Texte23.Value, "") ' RETOUR 2 MARCHE
    envoieSignataireDate = Nz(Me.Texte505.Value, "") ' DATE ENVOIE AU SIGNATAIRE
    ' *** NOUVELLE LOGIQUE ***
    ' 0. Vérifier si PRISE EN CHARGE est vide
    If priseEnChargeDate = "" And envoieSignataireDate <> "" Then
        MsgBox "Attention : La date d'envoi au signataire ne peut pas être remplie si la date de prise en charge est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' 1. Vérifier si un envoi est rempli sans son retour correspondant
    If (envoie1Date <> "" And retour1Date = "") Or (envoie2Date <> "" And retour2Date = "") Then
        MsgBox "Attention : Un envoi ne peut pas être rempli sans son retour correspondant.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' 2. Si pas d'envois 1 et 2, vérifier que la date d'envoi au signataire n'est pas antérieure à la date de prise en charge
    If envoie1Date = "" And envoie2Date = "" And envoieSignataireDate <> "" And priseEnChargeDate <> "" Then
        If IsDate(envoieSignataireDate) And IsDate(priseEnChargeDate) Then
            If CDate(envoieSignataireDate) < CDate(priseEnChargeDate) Then
                MsgBox "Attention : En l'absence d'envois 1 et 2, la date d'envoi au signataire ne peut pas être antérieure à la date de prise en charge.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    ' 3. Vérifier si l'envoi au signataire est plus ancien que le retour N°1 ou N°2 (si existants)
    If envoieSignataireDate <> "" Then
        If retour1Date <> "" And IsDate(retour1Date) And IsDate(envoieSignataireDate) Then
            If CDate(envoieSignataireDate) < CDate(retour1Date) Then
                MsgBox "Attention : L'envoi au signataire ne peut pas être antérieur au retour N°1.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retour2Date <> "" And IsDate(retour2Date) And IsDate(envoieSignataireDate) Then
            If CDate(envoieSignataireDate) < CDate(retour2Date) Then
                MsgBox "Attention : L'envoi au signataire ne peut pas être antérieur au retour N°2.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    Call UpdateTableField("DATE ENVOIE AU SIGNATAIRE", Me.Texte505)
    Call CalculerEtMettreAJourDelais
End Sub
' DATE SIGNATURE
Public Sub Texte25_BeforeUpdate(Cancel As Integer)
    Dim envoieSignataireDate As Variant
    Dim dateSignatureDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieSignataireDate = Nz(Me.Texte505.Value, "") ' DATE ENVOIE AU SIGNATAIRE
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    ' Vérifier si la date de signature est remplie alors que la date d'envoi au signataire est vide
    If dateSignatureDate <> "" And envoieSignataireDate = "" Then
        MsgBox "Attention : La date de signature ne peut pas être remplie si la date d'envoi au signataire est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si la date d'envoi au signataire est remplie
    If envoieSignataireDate <> "" Then
        ' Vérifier si la date de signature est plus ancienne que la date d'envoi au signataire
        If IsDate(envoieSignataireDate) And IsDate(dateSignatureDate) Then
            If CDate(dateSignatureDate) < CDate(envoieSignataireDate) Then
                MsgBox "Attention : La date de signature ne peut pas être antérieure à la date d'envoi au signataire.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    Call UpdateTableField("DATE SIGNATURE", Me.Texte25)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE DCE/COO
Public Sub Texte26_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieDCECOODate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieDCECOODate = Nz(Me.Texte26.Value, "") ' ENVOIE DCE/COO
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieDCECOODate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi DGA ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE DGA est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieDCECOODate) Then
        If CDate(envoieDCECOODate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi DGA est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE DCE/COO", Me.Texte26)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR DCE/COO
Public Sub Texte27_BeforeUpdate(Cancel As Integer)
    Dim envoieDCECOODate As Variant
    Dim retourDCECOODate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieDCECOODate = Nz(Me.Texte26.Value, "") ' ENVOIE DCE/COO
    retourDCECOODate = Nz(Me.Texte27.Value, "") ' RETOUR DCE/COO
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourDCECOODate <> "" And envoieDCECOODate = "" Then
        MsgBox "Attention : La date de retour DCE/COO ne peut pas être remplie si la date d'envoi DCE/COO est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR DCE/COO est antérieure à ENVOIE DCE/COO
    If IsDate(envoieDCECOODate) And IsDate(retourDCECOODate) Then
        If CDate(retourDCECOODate) < CDate(envoieDCECOODate) Then
            MsgBox "Attention : La date de retour DCE/COO est antérieure à la date d'envoi DCE/COO, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR DCE/COO", Me.Texte27)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ DCEE/COO
Public Sub Texte28_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ DCEE/COO", Me.Texte28)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE DGA
Public Sub Texte29_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieDGADate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieDGADate = Nz(Me.Texte29.Value, "") ' ENVOIE DGA
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieDGADate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi DGA ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE DGA est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieDGADate) Then
        If CDate(envoieDGADate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi DGA est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE DGA", Me.Texte29)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR DGA
Public Sub Texte30_BeforeUpdate(Cancel As Integer)
    Dim envoieDGADate As Variant
    Dim retourDGADate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieDGADate = Nz(Me.Texte29.Value, "") ' ENVOIE DGA
    retourDGADate = Nz(Me.Texte30.Value, "") ' RETOUR DGA
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourDGADate <> "" And envoieDGADate = "" Then
        MsgBox "Attention : La date de retour DGA ne peut pas être remplie si la date d'envoi DGA est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR DGA est antérieure à ENVOIE DGA
    If IsDate(envoieDGADate) And IsDate(retourDGADate) Then
        If CDate(retourDGADate) < CDate(envoieDGADate) Then
            MsgBox "Attention : La date de retour DGA est antérieure à la date d'envoi DGA, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR DGA", Me.Texte30)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ DGA
Public Sub Texte31_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ DGA", Me.Texte31)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE DG
Public Sub Texte32_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieDGDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieDGDate = Nz(Me.Texte32.Value, "") ' ENVOIE DG
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieDGDate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi DG ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE DG est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieDGDate) Then
        If CDate(envoieDGDate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi DG est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE DG", Me.Texte32)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR DG
Public Sub Texte654_BeforeUpdate(Cancel As Integer)
    Dim envoieDGDate As Variant
    Dim retourDGDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieDGDate = Nz(Me.Texte32.Value, "") ' ENVOIE DG
    retourDGDate = Nz(Me.Texte654.Value, "") ' RETOUR DG
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourDGDate <> "" And envoieDGDate = "" Then
        MsgBox "Attention : La date de retour DG ne peut pas être remplie si la date d'envoi DG est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR DG est antérieure à ENVOIE DG
    If IsDate(envoieDGDate) And IsDate(retourDGDate) Then
        If CDate(retourDGDate) < CDate(envoieDGDate) Then
            MsgBox "Attention : La date de retour DG est antérieure à la date d'envoi DG, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR DG", Me.Texte654)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ DG
Public Sub Texte33_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ DG", Me.Texte33)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE RISQCRE
Public Sub Texte34_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieRISQCREDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieRISQCREDate = Nz(Me.Texte34.Value, "") ' ENVOIE RISQCRE
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieRISQCREDate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi RISQCRE ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE RISQCRE est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieRISQCREDate) Then
        If CDate(envoieRISQCREDate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi RISQCRE est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE RISQCRE", Me.Texte34)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR RISQCRE
Public Sub Texte35_BeforeUpdate(Cancel As Integer)
    Dim envoieRISQCREDate As Variant
    Dim retourRISQCREDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieRISQCREDate = Nz(Me.Texte34.Value, "") ' ENVOIE RISQCRE
    retourRISQCREDate = Nz(Me.Texte35.Value, "") ' RETOUR RISQCRE
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourRISQCREDate <> "" And envoieRISQCREDate = "" Then
        MsgBox "Attention : La date de retour RISQCRE ne peut pas être remplie si la date d'envoi RISQCRE est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR RISQCRE est antérieure à ENVOIE RISQCRE
    If IsDate(envoieRISQCREDate) And IsDate(retourRISQCREDate) Then
        If CDate(retourRISQCREDate) < CDate(envoieRISQCREDate) Then
            MsgBox "Attention : La date de retour RISQCRE est antérieure à la date d'envoi RISQCRE, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR RISQCRE", Me.Texte35)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ RISQCRE
Public Sub Texte36_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ RISQCRE", Me.Texte36)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE CA
Public Sub Texte37_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieCADate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieCADate = Nz(Me.Texte37.Value, "") ' ENVOIE CA
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieCADate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi CA ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE CA est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieCADate) Then
        If CDate(envoieCADate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi CA est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE CA", Me.Texte37)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR CA
Public Sub Texte38_BeforeUpdate(Cancel As Integer)
    Dim envoieCADate As Variant
    Dim retourCADate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieCADate = Nz(Me.Texte37.Value, "") ' ENVOIE CA
    retourCADate = Nz(Me.Texte38.Value, "") ' RETOUR CA
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourCADate <> "" And envoieCADate = "" Then
        MsgBox "Attention : La date de retour CA ne peut pas être remplie si la date d'envoi CA est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR CA est antérieure à ENVOIE CA
    If IsDate(envoieCADate) And IsDate(retourCADate) Then
        If CDate(retourCADate) < CDate(envoieCADate) Then
            MsgBox "Attention : La date de retour CA est antérieure à la date d'envoi CA, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR CA", Me.Texte38)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ CA
Public Sub Texte39_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ CA", Me.Texte39)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE AUTRESPCRU
Public Sub Texte40_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieAUTRESPCRUDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieAUTRESPCRUDate = Nz(Me.Texte40.Value, "") ' ENVOIE AUTRESPCRU
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieAUTRESPCRUDate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi AUTRESPCRU ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE AUTRESPCRU est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieAUTRESPCRUDate) Then
        If CDate(envoieAUTRESPCRUDate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi AUTRESPCRU est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE AUTRESPCRU", Me.Texte40)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR AUTRESPCRU
Public Sub Texte41_BeforeUpdate(Cancel As Integer)
    Dim envoieAUTRESPCRUDate As Variant
    Dim retourAUTRESPCRUDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieAUTRESPCRUDate = Nz(Me.Texte40.Value, "") ' ENVOIE AUTRESPCRU
    retourAUTRESPCRUDate = Nz(Me.Texte41.Value, "") ' RETOUR AUTRESPCRU
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourAUTRESPCRUDate <> "" And envoieAUTRESPCRUDate = "" Then
        MsgBox "Attention : La date de retour AUTRESPCRU ne peut pas être remplie si la date d'envoi AUTRESPCRU est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR AUTRESPCRU est antérieure à ENVOIE AUTRESPCRU
    If IsDate(envoieAUTRESPCRUDate) And IsDate(retourAUTRESPCRUDate) Then
        If CDate(retourAUTRESPCRUDate) < CDate(envoieAUTRESPCRUDate) Then
            MsgBox "Attention : La date de retour AUTRESPCRU est antérieure à la date d'envoi AUTRESPCRU, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
           Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR AUTRESPCRU", Me.Texte41)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ AUTRESPCRU
Public Sub Texte42_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ AUTRESPCRU", Me.Texte42)
    Call CalculerEtMettreAJourDelais
End Sub
' DECISION
Public Sub Modifiable43_BeforeUpdate(Cancel As Integer)
    Dim etatValue As String
    ' Récupérer la valeur du champ ETAT
    etatValue = Nz(Me.Modifiable16.Value, "")
    ' Vérifier si ETAT est différent de "Notifie OK" et "Dossier non valide / autres"
    If etatValue <> "Notifie OK" And etatValue <> "Dossier non valide / autres" Then
        MsgBox "Vous ne pouvez pas modifier la décision si l'état n'est pas 'Notifie OK' ou 'Dossier non valide / autres'.", vbExclamation
        Cancel = True ' Empêcher la mise à jour
        Exit Sub
    End If
    Call UpdateTableField("DECISION", Me.Modifiable43)
End Sub
' MONTANT ACCORDE
Public Sub Texte44_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("MONTANT ACCORDE", Me.Texte44)
End Sub
Public Sub Texte45_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureFinale As Variant
    Dim envoieDCECOODate As Variant, retourDCECOODate As Variant
    Dim envoieDGADate As Variant, retourDGADate As Variant
    Dim envoieDGDate As Variant, retourDGDate As Variant
    Dim envoieRISQCREDate As Variant, retourRISQCREDate As Variant
    Dim envoieCADate As Variant, retourCADate As Variant
    Dim envoieAUTRESPCRUDate As Variant, retourAUTRESPCRUDate As Variant
    Dim SIGNATUREDate As Variant
    Dim auMoinsUnePaireRemplie As Boolean
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureFinale = Nz(Me.Texte45.Value, "") ' DATE SIGNATURE FINAL
    envoieDCECOODate = Nz(Me.Texte26.Value, "") ' ENVOIE DCE/COO
    retourDCECOODate = Nz(Me.Texte27.Value, "") ' RETOUR DCE/COO
    envoieDGADate = Nz(Me.Texte29.Value, "") ' ENVOIE DGA
    retourDGADate = Nz(Me.Texte30.Value, "") ' RETOUR DGA
    envoieDGDate = Nz(Me.Texte32.Value, "") ' ENVOIE DG
    retourDGDate = Nz(Me.Texte654.Value, "") ' RETOUR DG
    envoieRISQCREDate = Nz(Me.Texte34.Value, "") ' ENVOIE RISQCRE
    retourRISQCREDate = Nz(Me.Texte35.Value, "") ' RETOUR RISQCRE
    envoieCADate = Nz(Me.Texte37.Value, "") ' ENVOIE CA
    retourCADate = Nz(Me.Texte38.Value, "") ' RETOUR CA
    envoieAUTRESPCRUDate = Nz(Me.Texte40.Value, "") ' ENVOIE AUTRESPCRU
    retourAUTRESPCRUDate = Nz(Me.Texte41.Value, "") ' RETOUR AUTRESPCRU
    SIGNATUREDate = Nz(Me.Texte25.Value, "") ' SIGNATURE SIGNATAIRE
    ' Vérification de la date de signature
    If SIGNATUREDate = "" Then
        MsgBox "Attention : Vous devez spécifier une date de signature.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si la date de signature finale n'est pas antérieure à la date de signature
    If IsDate(dateSignatureFinale) And IsDate(SIGNATUREDate) Then
        If CDate(dateSignatureFinale) < CDate(SIGNATUREDate) Then
            MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de signature.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
    ' Initialiser la variable pour vérifier si au moins une paire est remplie
    auMoinsUnePaireRemplie = False
    ' Vérifications des paires Envoi/Retour et mise à jour de la variable auMoinsUnePaireRemplie
    If (envoieDCECOODate <> "" And retourDCECOODate = "") Or (envoieDCECOODate = "" And retourDCECOODate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour DCE/COO.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieDCECOODate <> "" And retourDCECOODate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    If (envoieDGADate <> "" And retourDGADate = "") Or (envoieDGADate = "" And retourDGADate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour DGA.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieDGADate <> "" And retourDGADate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    If (envoieDGDate <> "" And retourDGDate = "") Or (envoieDGDate = "" And retourDGDate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour DG.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieDGDate <> "" And retourDGDate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    If (envoieRISQCREDate <> "" And retourRISQCREDate = "") Or (envoieRISQCREDate = "" And retourRISQCREDate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour RISQCRE.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieRISQCREDate <> "" And retourRISQCREDate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    If (envoieCADate <> "" And retourCADate = "") Or (envoieCADate = "" And retourCADate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour CA.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieCADate <> "" And retourCADate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    If (envoieAUTRESPCRUDate <> "" And retourAUTRESPCRUDate = "") Or (envoieAUTRESPCRUDate = "" And retourAUTRESPCRUDate <> "") Then
        MsgBox "Attention : Vous devez spécifier une date d'envoi ET une date de retour pour AUTRESPCRU.", vbExclamation
        Cancel = True
        Exit Sub
    ElseIf (envoieAUTRESPCRUDate <> "" And retourAUTRESPCRUDate <> "") Then
        auMoinsUnePaireRemplie = True
    End If
    ' Vérifier si au moins une paire Envoi/Retour est remplie
    ' **Suppression de cette vérification**
    ' Vérifier si la date de signature finale est antérieure à l'une des dates d'envoi ou de retour
    If IsDate(dateSignatureFinale) Then
        If envoieDCECOODate <> "" And IsDate(envoieDCECOODate) Then
            If CDate(dateSignatureFinale) < CDate(envoieDCECOODate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi DCE/COO.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourDCECOODate <> "" And IsDate(retourDCECOODate) Then
            If CDate(dateSignatureFinale) < CDate(retourDCECOODate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour DCE/COO.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If envoieDGADate <> "" And IsDate(envoieDGADate) Then
            If CDate(dateSignatureFinale) < CDate(envoieDGADate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi DGA.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourDGADate <> "" And IsDate(retourDGADate) Then
            If CDate(dateSignatureFinale) < CDate(retourDGADate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour DGA.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If envoieDGDate <> "" And IsDate(envoieDGDate) Then
            If CDate(dateSignatureFinale) < CDate(envoieDGDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi DG.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourDGDate <> "" And IsDate(retourDGDate) Then
            If CDate(dateSignatureFinale) < CDate(retourDGDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour DG.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If envoieRISQCREDate <> "" And IsDate(envoieRISQCREDate) Then
            If CDate(dateSignatureFinale) < CDate(envoieRISQCREDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi RISQCRE.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourRISQCREDate <> "" And IsDate(retourRISQCREDate) Then
            If CDate(dateSignatureFinale) < CDate(retourRISQCREDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour RISQCRE.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If envoieCADate <> "" And IsDate(envoieCADate) Then
            If CDate(dateSignatureFinale) < CDate(envoieCADate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi CA.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourCADate <> "" And IsDate(retourCADate) Then
            If CDate(dateSignatureFinale) < CDate(retourCADate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour CA.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If envoieAUTRESPCRUDate <> "" And IsDate(envoieAUTRESPCRUDate) Then
            If CDate(dateSignatureFinale) < CDate(envoieAUTRESPCRUDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date d'envoi AUTRESPCRU.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        If retourAUTRESPCRUDate <> "" And IsDate(retourAUTRESPCRUDate) Then
            If CDate(dateSignatureFinale) < CDate(retourAUTRESPCRUDate) Then
                MsgBox "Attention : La date de signature finale ne peut pas être antérieure à la date de retour AUTRESPCRU.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    Call UpdateTableField("DATE SIGNATURE FINAL", Me.Texte45)
    Call CalculerEtMettreAJourDelais
End Sub
' DATE ENVOIE TRAME
Public Sub Texte46_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureFinale As Variant
    Dim dateEnvoieTrame As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureFinale = Nz(Me.Texte45.Value, "") ' DATE SIGNATURE FINAL
    dateEnvoieTrame = Nz(Me.Texte46.Value, "") ' DATE ENVOIE TRAME
    ' Vérifier si la date d'envoi de la trame est remplie alors que la date de signature finale est vide
    If dateEnvoieTrame <> "" And dateSignatureFinale = "" Then
        MsgBox "Attention : La date d'envoi de la trame ne peut pas être remplie si la date de signature finale est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si la date d'envoi de la trame est antérieure à la date de signature finale
    If IsDate(dateSignatureFinale) And IsDate(dateEnvoieTrame) Then
        If CDate(dateEnvoieTrame) < CDate(dateSignatureFinale) Then
            MsgBox "Attention : La date d'envoi de la trame ne peut pas être antérieure à la date de signature finale.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
    Call UpdateTableField("DATE ENVOIE TRAME", Me.Texte46)
    Call CalculerEtMettreAJourDelais
End Sub
' Date de notification
Public Sub Texte47_BeforeUpdate(Cancel As Integer)
    Dim dateEnvoieTrame As Variant
    Dim dateNotification As Variant
    Dim analysteNotification As String
    ' Récupérer les valeurs des champs (en gérant les valeurs Null)
    dateEnvoieTrame = Nz(Me.Texte46.Value, "") ' DATE ENVOIE TRAME
    dateNotification = Nz(Me.Texte47.Value, "") ' Date de notification
    analysteNotification = Nz(Me.Modifiable48.Value, "") ' ANALYSTE NOTIFICATION
    ' Vérifier si la date de notification est remplie alors que la date d'envoi de la trame est vide
    If dateNotification <> "" And dateEnvoieTrame = "" Then
        MsgBox "Attention : La date de notification ne peut pas être remplie si la date d'envoi de la trame est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si la date de notification est antérieure à la date d'envoi de la trame
    If IsDate(dateEnvoieTrame) And IsDate(dateNotification) Then
        If CDate(dateNotification) < CDate(dateEnvoieTrame) Then
            MsgBox "Attention : La date de notification ne peut pas être antérieure à la date d'envoi de la trame.", vbExclamation
            Cancel = True
            Exit Sub
        End If
    End If
    ' Vérifier si l'analyste de notification est vide
    If analysteNotification = "" And dateNotification <> "" Then
        MsgBox "Attention : Vous devez sélectionner un analyste de notification avant de définir la date de notification.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    Call UpdateTableField("Date de notification", Me.Texte47)
    Call CalculerEtMettreAJourDelais
End Sub
' ANALYSTE NOTIFICATION
Public Sub Modifiable48_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("ANALYSTE NOTIFICATION", Me.Modifiable48)
End Sub
' ENVOIE CORISQ
Public Sub Texte815_BeforeUpdate(Cancel As Integer)
    Dim dateSignatureDate As Variant
    Dim envoieCORISQDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    dateSignatureDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    envoieCORISQDate = Nz(Me.Texte815.Value, "") ' ENVOIE CORISQ
    ' Vérifier si la date d'envoi est remplie alors que la date de signature est vide
    If envoieCORISQDate <> "" And dateSignatureDate = "" Then
        MsgBox "Attention : La date d'envoi CORISQ ne peut pas être remplie si la date de signature est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si ENVOIE CORISQ est antérieure à DATE SIGNATURE
    If IsDate(dateSignatureDate) And IsDate(envoieCORISQDate) Then
        If CDate(envoieCORISQDate) < CDate(dateSignatureDate) Then
            MsgBox "Attention : La date d'envoi CORISQ est antérieure à la date de signature, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("ENVOIE CORISQ", Me.Texte815)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR CORISQ
Public Sub Texte818_BeforeUpdate(Cancel As Integer)
    Dim envoieCORISQDate As Variant
    Dim retourCORISQDate As Variant
    ' Récupérer les dates depuis les champs (en gérant les valeurs Null)
    envoieCORISQDate = Nz(Me.Texte815.Value, "") ' ENVOIE CORISQ
    retourCORISQDate = Nz(Me.Texte818.Value, "") ' RETOUR CORISQ
    ' Vérifier si la date de retour est remplie alors que la date d'envoi est vide
    If retourCORISQDate <> "" And envoieCORISQDate = "" Then
        MsgBox "Attention : La date de retour CORISQ ne peut pas être remplie si la date d'envoi CORISQ est vide.", vbExclamation
        Cancel = True
        Exit Sub
    End If
    ' Vérifier si RETOUR CORISQ est antérieure à ENVOIE CORISQ
    If IsDate(envoieCORISQDate) And IsDate(retourCORISQDate) Then
        If CDate(retourCORISQDate) < CDate(envoieCORISQDate) Then
            MsgBox "Attention : La date de retour CORISQ est antérieure à la date d'envoi CORISQ, supprimez ou corrigez.", vbExclamation
            Cancel = True ' Empêcher la mise à jour
            Exit Sub
        End If
    End If
    Call UpdateTableField("RETOUR CORISQ", Me.Texte818)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRESJ CORISQ
Public Sub Texte821_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRESJ CORISQ", Me.Texte821)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE 1 SIGN MARCHE", Me.Texte897
Public Sub Texte897_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("ENVOIE 1 SIGN MARCHE", Me.Texte897)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR 1 SIGN MARCHE", Me.Texte899
Public Sub Texte899_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("RETOUR 1 SIGN MARCHE", Me.Texte899)
    Call CalculerEtMettreAJourDelais
End Sub
' ENVOIE 2 SIGN MARCHE", Me.Texte901
Public Sub Texte901_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("ENVOIE 2 SIGN MARCHE", Me.Texte901)
    Call CalculerEtMettreAJourDelais
End Sub
' RETOUR 2 SIGN MARCHE", Me.Texte903
Public Sub Texte903_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("RETOUR 2 SIGN MARCHE", Me.Texte903)
    Call CalculerEtMettreAJourDelais
End Sub
' AUTRES J SIGN MARCHE", Me.Texte905
Public Sub Texte905_BeforeUpdate(Cancel As Integer)
    Call UpdateTableField("AUTRES J SIGN MARCHE", Me.Texte905)
    Call CalculerEtMettreAJourDelais
End Sub
Public Sub UpdateTableField(fieldName As String, targetControl As Control)
    On Error GoTo ErrorHandler
    Dim IDValue As String
    Dim NewFieldValue As String
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim IDSGFieldType As DAO.DataTypeEnum
    Dim TableName As String
    Dim FieldType As DAO.DataTypeEnum
    Dim formattedDate As String
    ' Get the ID value from Texte162
    IDValue = Nz(Me.Texte162.Value, "")
    ' Get the new field value from the target control
    NewFieldValue = Nz(targetControl.Value, "")
    ' Set the table name
    TableName = "[DATABASE]" ' REMPLACEZ CECI PAR LE NOM DE VOTRE TABLE entre []
    ' Check if the ID is not empty
    If IDValue <> "" Then
        ' Determine the data type of the IDSG field
        Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 IDSG, [" & fieldName & "] FROM " & TableName)
        IDSGFieldType = rs.Fields("IDSG").Type
        FieldType = rs.Fields(fieldName).Type ' Get the type of the field being updated
        rs.Close
        Set rs = Nothing
        ' Construct the SQL UPDATE statement based on the type IDSG
        ' Use the FieldType to determine how to handle the update
        Select Case FieldType
            Case dbText, dbMemo
                ' Text/Memo
                sql = "UPDATE " & TableName & " SET [" & fieldName & "] = '" & Replace(NewFieldValue, "'", "''") & "' WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
            Case dbInteger, dbLong, dbCurrency, dbDouble, dbSingle
                ' Numeric types
                If NewFieldValue = "" Then
                    sql = "UPDATE " & TableName & " SET [" & fieldName & "] = Null WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
                Else
                    sql = "UPDATE " & TableName & " SET [" & fieldName & "] = " & Replace(NewFieldValue, "'", "''") & " WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
                End If
            Case dbDate
                ' Date type
                If NewFieldValue = "" Then
                    sql = "UPDATE " & TableName & " SET [" & fieldName & "] = Null WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
                Else
                    ' Format the date as yyyy-mm-dd for database storage
                    formattedDate = Format(CDate(NewFieldValue), "yyyy-mm-dd")
                    sql = "UPDATE " & TableName & " SET [" & fieldName & "] = #" & formattedDate & "# WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
                End If
            Case dbBoolean
                ' Oui/Non - filet de securite si un champ booleen
                ' venait a transiter par cette procedure
                sql = "UPDATE " & TableName & " SET [" & fieldName & "] = " & _
                      IIf(Nz(targetControl.Value, False), "True", "False") & _
                      " WHERE IDSG = '" & Replace(IDValue, "'", "''") & "'"
            Case Else
                MsgBox "Type de données non supporté pour le champ " & fieldName, vbCritical, "Erreur de Type de Données"
                Exit Sub
        End Select
        ' Execute the query
        CurrentDb.Execute sql, dbFailOnError
        Debug.Print "Updated " & fieldName & " to: " & NewFieldValue & " for IDSG: " & IDValue
        Debug.Print "SQL: " & sql
    Else
        MsgBox "IDSG is empty. Cannot update the database.", vbExclamation, "Update Failed"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error Number: " & Err.number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "SQL: " & sql, vbCritical, "Update Error"
End Sub
Public Sub Commande825_Click()
    Dim dateSignatureFinale As Date
    Dim retourDCECOODate As Variant, retourDGADate As Variant, retourDGDate As Variant
    Dim retourRISQCREDate As Variant, retourCORISQDate As Variant, retourCADate As Variant
    Dim retourAUTRESPCRUDate As Variant
    Dim SIGNATUREDate As Variant ' Ajout de la variable pour DATE SIGNATURE
    Dim auMoinsUneDateRetour As Boolean
    ' Récupérer les dates de retour depuis les champs (en gérant les valeurs Null)
    retourDCECOODate = Nz(Me.Texte27.Value, "") ' RETOUR DCE/COO
    retourDGADate = Nz(Me.Texte30.Value, "") ' RETOUR DGA
    retourDGDate = Nz(Me.Texte654.Value, "") ' RETOUR DG
    retourRISQCREDate = Nz(Me.Texte35.Value, "") ' RETOUR RISQCRE
    retourCORISQDate = Nz(Me.Texte818.Value, "") ' RETOUR CORISQ
    retourCADate = Nz(Me.Texte38.Value, "") ' RETOUR CA
    retourAUTRESPCRUDate = Nz(Me.Texte41.Value, "") ' RETOUR AUTRESPCRU
    SIGNATUREDate = Nz(Me.Texte25.Value, "") ' DATE SIGNATURE
    ' Initialiser la date de signature finale avec la date la plus ancienne possible
    dateSignatureFinale = DateSerial(1900, 1, 1)
    ' Initialiser la variable pour vérifier si au moins une date de retour est renseignée
    auMoinsUneDateRetour = False
    ' Comparer chaque date de retour avec la date de signature finale et mettre à jour si nécessaire
    If IsDate(retourDCECOODate) Then
        auMoinsUneDateRetour = True
        If CDate(retourDCECOODate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourDCECOODate)
        End If
    End If
    If IsDate(retourDGADate) Then
        auMoinsUneDateRetour = True
        If CDate(retourDGADate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourDGADate)
        End If
    End If
    If IsDate(retourDGDate) Then
        auMoinsUneDateRetour = True
        If CDate(retourDGDate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourDGDate)
        End If
    End If
    If IsDate(retourRISQCREDate) Then
        auMoinsUneDateRetour = True
        If CDate(retourRISQCREDate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourRISQCREDate)
        End If
    End If
    If IsDate(retourCORISQDate) Then
        auMoinsUneDateRetour = True
        If CDate(retourCORISQDate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourCORISQDate)
        End If
    End If
    If IsDate(retourCADate) Then
        auMoinsUneDateRetour = True
        If CDate(retourCADate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourCADate)
        End If
    End If
    If IsDate(retourAUTRESPCRUDate) Then
        auMoinsUneDateRetour = True
        If CDate(retourAUTRESPCRUDate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(retourAUTRESPCRUDate)
        End If
    End If
    ' Si la date de signature est spécifiée, comparer avec la date finale actuelle
    If IsDate(SIGNATUREDate) Then
        If CDate(SIGNATUREDate) > dateSignatureFinale Then
            dateSignatureFinale = CDate(SIGNATUREDate)
        End If
    End If
    ' Ecrire la date la plus récente dans le champ Texte45
    If IsDate(SIGNATUREDate) Or auMoinsUneDateRetour Then
        If dateSignatureFinale > DateSerial(1900, 1, 1) Then
            Me.Texte45.Value = dateSignatureFinale
        Else
            Me.Texte45.Value = Null ' Ou "" si vous préférez une chaîne vide
        End If
    Else
        MsgBox "Il n'y a aucune date de signature.", vbInformation
        Me.Texte45.Value = Null ' Effacer le champ Texte45 si aucune date n'est renseignée
    End If
End Sub
Public Sub Commande798_Click() 'ENVOIE NOTE POUR NOVA
    Dim olApp As Object
    Dim olMail As Object
    Dim res As VbMsgBoxResult
    Dim fd As Object
    Dim fpath As String
    Dim corps As String
    Const msoFileDialogFilePicker = 3
    Dim test1 As String
    On Error Resume Next
    ' Ouvrir Outlook (ou le créer s’il n’est pas ouvert)
    Set olApp = GetObject(, "Outlook.Application")
    If Err.number <> 0 Then
        MsgBox "Il est nécessaire d'ouvrir outlook en amont"
        Exit Sub
    End If
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
        If Err.number <> 0 Then
            MsgBox "Erreur lors de la création d'Outlook : " & Err.Description
            Exit Sub
        End If
    End If
    On Error GoTo 0
    ' Demande d’attachement (AVANT la création de l'e-mail)
    res = MsgBox("Souhaitez-vous joindre un document complémentaire à cet e-mail ? (Accord Consignataire SGCI)", vbQuestion + vbYesNo, "Joindre un fichier")
    If res = vbYes Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        If Not fd Is Nothing Then
            With fd
                .AllowMultiSelect = False
                .Title = "Sélectionnez le document à joindre"
                If .Show = -1 Then
                    fpath = .SelectedItems(1)
                End If
            End With
            Set fd = Nothing
        End If
    End If
    ' Créer un nouveau mail
    Set olMail = olApp.CreateItem(0) ' 0 = olMailItem
    With olMail
        ' Sujet
        .Subject = "VALIDATION DE LA NOTE DANS NOVA - RCT N° " & Nz(Me.Texte1.Value, "")
        ' Corps du mail
        corps = "Bonjour," & "<br><br>"
        corps = corps & "Je vous prie de bien vouloir valider la note suivante dans NOVA, afin d'envoyer le dossier aux signataires :" & "<br><br>"
        corps = corps & "<ul>"
        corps = corps & "  <li><b>RCT :</b> " & Nz(Me.Texte1.Value, "") & "</li>"
        corps = corps & "  <li><b>CLIENT :</b> " & Nz(Me.Texte4.Value, "") & "</li>"
        corps = corps & "  <li><b>NOTE :</b> " & Nz(Me.Modifiable5.Value, "") & "</li>"
        corps = corps & "  <li><b>DVA :</b> " & Nz(Me.Texte17.Value, "") & "</li>"
        corps = corps & "</ul>"
        corps = corps & "Je vous remercie par avance pour votre diligence." & "<br><br>"
        corps = corps & "Cordialement,"
        .htmlBody = corps
        'Ajout de la pièce jointe (SI il y en a une)
         If Len(fpath) > 0 Then
            olMail.Attachments.Add fpath
        End If
        .Display
    End With
    ' Nettoyage
    Set olMail = Nothing
    Set olApp = Nothing
End Sub
Public Sub Commande106_Click() 'DEMANDE D'AT
Dim olApp As Object
Dim olMail As Object
Dim f As Object
Dim v199 As String '’ Nom du client (Texte199)
Dim v174 As String ' Numéro client (Texte174)
Dim v162 As String
Dim v233 As String '’ DVA (Texte233)
Dim duree As String '’ Nombre de mois (depuis InputBox)
' Récupérer les données depuis le formulaire "5_Modifications"
Set f = Forms("5_Modifications")
v199 = Nz(f!Texte4, "")
v174 = Nz(f!Texte1, "")
v233 = Nz(f!Texte17, "")
v162 = Nz(f!Texte162, "")
' Demander le nombre de mois via InputBox
duree = InputBox("Entrez le nombre de mois pour l'AT :", "Nombre de mois", "X")
If duree = "" Then
    MsgBox "Opération annulée.", vbExclamation
    Exit Sub
End If
' Sujet
Dim sujet As String
sujet = "DEMANDE D'AT - ID SGCI N°" & v162 & " - " & v199
' Corps du mail en HTML
Dim htmlBody As String
htmlBody = "<p>Bonjour Yha, Danielle,</p>" & _
           "<p>Je vous sollicite pour la mise en place d'une AT (<b>" & duree & " mois</b>) pour :</p>" & _
           "<ul><li>le client <b>" & v199 & "</b> <span>, ID RCT</span> <b>" & v174 & "</b>, DVA au <b>" & v233 & "</b>.</li></ul>" & _
           "<p>Vous remerciant par avance pour le traitement,</p>" & _
           "<p>Bien à vous</p>"
' Adresses (à remplacer par vos destinataires réels ou les récupérer du formulaire)
Dim toAddr As String
Dim ccAddr As String
Dim bccAddr As String
toAddr = "danielle.kpegne@socgen.com; yha.traore@socgen.com"                           ' premier destinataire
ccAddr = "RISQ_CDR@bhfm-ci.fr.socgen.com"   ' deux personnes en copie
' Ouvre Outlook et crée le mail
On Error Resume Next
Set olApp = GetObject(, "Outlook.Application")
If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
End If
On Error GoTo 0
If olApp Is Nothing Then
    MsgBox "Outlook n'est pas disponible.", vbExclamation
    Exit Sub
End If
Set olMail = olApp.CreateItem(0)
With olMail
    .Subject = sujet
    .To = toAddr
    .CC = ccAddr
    .BCC = bccAddr
    .htmlBody = htmlBody
    .Display  ' Ou .Send si vous voulez envoyer directement
End With
Set olMail = Nothing
Set olApp = Nothing
Set f = Nothing
End Sub
Public Sub Commande797_Click() 'ENVOIE TRAME
    Dim olApp As Object
    Dim olMail As Object
    Dim f As Object
    Dim v199 As String ' Nom du client (Texte199)
    Dim v174 As String ' Numéro client (Texte174)
    Dim duree As String ' Lien du fichier Share
    Dim fd As FileDialog
    Dim varFile As Variant
    Dim signatureFile As String
    Dim recommendationFile As String
    Dim mailFile As String
    Dim trameFile As String
    Dim Modifiable48Value As String ' Variable to store the value of Modifiable48
    Dim response As Integer ' Variable to store the response from the MsgBox
    ' Récupérer les données depuis le formulaire "5_Modifications"
    Set f = Forms("5_Modifications")
    v199 = Nz(f!Texte4, "")
    v174 = Nz(f!Texte1, "")
    Modifiable48Value = Nz(f!Modifiable48, "") ' Get the value of Modifiable48
    ' Check if Modifiable48 has a value
    If Modifiable48Value = "" Then
        MsgBox "Veuillez entrer le nom de l'analyste pour la notification (Anlst. NOTIF).", vbExclamation
        Exit Sub
    End If
    ' Confirm that the documents are prepared
    response = MsgBox("Vous allez devoir joindre les documents suivants par piece jointe, les avez vous préparés :" & vbCrLf & _
                       "page de signature, recommandation, échanges par mail, trame ?", vbYesNo, "Confirmation")
    ' If the user clicks "No", exit the subroutine
    If response = vbNo Then
        Exit Sub
    End If
    ' Définir le sujet
    Dim sujet As String
    sujet = "POUR NOTIFICATION - " & v199
    ' Corps du mail en HTML
    Dim htmlBody As String
    htmlBody = "<p>Bonjour, " & Modifiable48Value & "</p>" & _
               "<p>Comme échangé, tu trouveras ci-joint les documents nécessaires afin de réaliser la notification de " & v199 & " (RCT: <b>" & v174 & "</b>) :</p>" & _
               "<p>Lien du fichier sur le share : <span style=""color:red; font-weight:bold;"">A INDIQUER</span></p>" & _
               "<p>Bien à toi</p>"
    ' Adresses (à remplacer par vos destinataires réels ou les récupérer du formulaire)
    Dim toAddr As String
    Dim ccAddr As String
    Dim bccAddr As String
    toAddr = ";"                           ' premier destinataire
    ccAddr = ""
    bccAddr = ""
    ' Boîte de dialogue de sélection de fichiers
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False ' On demande un fichier à la fois
        .Title = "Sélectionnez les fichiers (un par un) dans l'ordre : Signature, Recommandation, Mail, Trame"
        .Filters.Clear
        .Filters.Add "Tous les fichiers", "*.*"
        ' Sélection du fichier de signature
        MsgBox ("Sélectionnez la page de signature (PDF)")
        If .Show = -1 Then
            signatureFile = .SelectedItems(1)
        Else
            MsgBox "Sélection du fichier de Signature annulée.", vbExclamation
            Exit Sub
        End If
        ' Sélection du fichier de recommandation
        MsgBox ("Sélectionnez la recommandation (Word)")
        If .Show = -1 Then
            recommendationFile = .SelectedItems(1)
        Else
            MsgBox "Sélection du fichier de Recommandation annulée.", vbExclamation
            Exit Sub
        End If
        ' Sélection du fichier Mail
        MsgBox ("Sélectionnez les échanges par mails")
        If .Show = -1 Then
            mailFile = .SelectedItems(1)
        Else
            MsgBox "Sélection du fichier Mail annulée.", vbExclamation
            Exit Sub
        End If
        ' Sélection du fichier Trame
        MsgBox ("Sélectionnez la trame (Word)")
        If .Show = -1 Then
            trameFile = .SelectedItems(1)
        Else
            MsgBox "Sélection du fichier Trame annulée.", vbExclamation
            Exit Sub
        End If
    End With
    ' Ouvre Outlook et crée le mail
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlook n'est pas disponible.", vbExclamation
        Exit Sub
    End If
    Set olMail = olApp.CreateItem(0)
    With olMail
        .Subject = sujet
        .To = toAddr
        .CC = ccAddr
        .BCC = bccAddr
        .htmlBody = htmlBody
        ' Ajouter les pièces jointes
        .Attachments.Add signatureFile
        .Attachments.Add recommendationFile
        .Attachments.Add mailFile
        .Attachments.Add trameFile
        .Display  ' Ou .Send si vous voulez envoyer directement
    End With
    ' Nettoyage des objets
    Set olMail = Nothing
    Set olApp = Nothing
    Set f = Nothing
    Set fd = Nothing
End Sub
Public Sub Commande809_Click()
    Dim objOutlook As Object
    Dim objMail As Object
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim strBody As String
    Dim strClient As String
    ' Récupérer la valeur du champ Texte4
    strClient = Nz(Me.Texte4.Value, "")
    strBody = "Bonsoir à tous,<br><br>" & _
          "Nous vous prions de trouver ci-joint la notification sur le client : <b>" & strClient & "</b><br><br>" & _
          "Bonne réception,"
    ' Afficher un message pour demander de sélectionner le fichier PDF
    MsgBox "Veuillez sélectionner la notification PDF.", vbInformation, "Sélection du fichier PDF"
    ' Utiliser le FileDialog pour sélectionner le fichier PDF
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Sélectionner le fichier PDF à joindre"
        .Filters.Clear
        .Filters.Add "Fichiers PDF", "*.pdf"
        If .Show = -1 Then ' Si l'utilisateur a sélectionné un fichier
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "Aucun fichier sélectionné.", vbExclamation
            Exit Sub
        End If
    End With
    ' Créer l'objet Outlook
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0) ' 0 pour un nouvel e-mail
    ' Ajouter la pièce jointe et configurer le mail
    With objMail
        .CC = "" ' Ajouter les adresses mail en copie ici
        .BCC = "" ' Ajouter les adresses mail en copie cachée ici
        .Subject = "NOTIFICATION - " & strClient ' Définir l'objet du mail
        .htmlBody = strBody ' Utiliser HTMLBody pour interpréter le code HTML
        .Attachments.Add selectedFile
        .Display ' Afficher le mail pour relecture et envoi
        '.Send ' Envoyer le mail directement (à utiliser avec précaution)
    End With
    ' Libérer les objets
    Set objMail = Nothing
    Set objOutlook = Nothing
    Set fd = Nothing
End Sub
'Private Sub Texte958_Click()
'    Call CalculerEtMettreAJourDelais
 '   Call UpdateDatabaseField("DELAIS AD", Me.Option958)
'End Sub

