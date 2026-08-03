'###########################################################
'#  MODULE STANDARD  DELAIS_
'#  Version refondue - calcul par phases disjointes
'#
'#  IDENTITES GARANTIES PAR CONSTRUCTION
'#
'#    DELAIS DRC   = AD + ANALYSTE + SIGNATAIRE + NOTIFICATION
'#    DELAIS TOTAL = DELAIS DRC + DELAIS HORS DRC
'#                 = JoursOuvrables(RECEPTION, DATE FINALE)
'#
'#  PRINCIPE
'#
'#    La ligne de temps RECEPTION -> DATE FINALE est decoupee
'#    en phases contigues et disjointes. Chaque phase est
'#    d'abord calculee en brut, puis on lui retire la SEULE
'#    part des periodes hors DRC qui tombe reellement a
'#    l'interieur de ses bornes.
'#
'#    L'analyste n'est donc plus diminue que du temps ou le
'#    dossier n'etait effectivement pas chez lui. Les trois
'#    causes de rognage abusif de l'ancienne version sont
'#    eliminees :
'#      - deduction de m1 et m2 sans fusion des chevauchements
'#      - deduction de periodes debordant hors de la fenetre
'#      - deduction de forfaits sans plafonnement
'#
'#  REMPLACE INTEGRALEMENT L'ANCIEN MODULE DELAIS_
'#  Point d'entree inchange : DELAIS(IDSGValue)
'###########################################################

Option Explicit
Option Compare Database


'===========================================================
'  PARAMETRAGE
'===========================================================

' Segment DATE SIGNATURE -> DATE ENVOIE TRAME.
' Il couvre le circuit de validation externe (DCE/COO, DGA, DG,
' RISQCRE, CA, CORISQ, AUTRESPCRU) et la redaction de la trame.
' La part externe est deja neutralisee via HORS DRC ; ce qui
' reste est du temps DRC et doit etre rattache a une phase pour
' que l'identite tienne.
'
' Valeurs admises : "ANALYSTE" (defaut), "AD", "NOTIFICATION"
Private Const PHASE_TRAME As String = "ANALYSTE"

' Mettre a True pour tracer le detail du calcul dans la fenetre
' Execution (Ctrl+G). A laisser a False en production.
Private Const TRACE_CALCUL As Boolean = False


'===========================================================
'  POINT D'ENTREE
'===========================================================
Public Sub DELAIS(ByVal IDSGValue As String)

    On Error GoTo GestionErreur

    Dim rs As DAO.Recordset
    Dim intervalles() As Variant
    Dim nbInt As Long

    ' Jalons du circuit
    Dim J0 As Date, J1 As Date, J2 As Date, J3 As Date, J4 As Date, J5 As Date

    ' Brut, part hors DRC, forfait retenu, net - par phase
    Dim brutAD As Long, brutAna As Long, brutSig As Long, brutTrame As Long, brutNotif As Long
    Dim hdAD As Long, hdAna As Long, hdSig As Long, hdTrame As Long, hdNotif As Long
    Dim fftAna As Long, fftSig As Long, fftTrame As Long
    Dim netAD As Long, netAna As Long, netSig As Long, netTrame As Long, netNotif As Long

    Dim horsDRC As Long, totalDRC As Long, totalGlobal As Long
    Dim controle As Long

    If Len(Trim(IDSGValue)) = 0 Then Exit Sub

    Set rs = CurrentDb.OpenRecordset( _
             "SELECT * FROM [DATABASE] WHERE " & GetCritere(IDSGValue))

    If rs.EOF Then GoTo Nettoyage

    '-------------------------------------------------------
    ' 1. Jalons, normalises en suite croissante
    '-------------------------------------------------------
    NormaliserJalons rs, J0, J1, J2, J3, J4, J5

    '-------------------------------------------------------
    ' 2. Union des periodes hors DRC, bornee a [J0 ; J5]
    '-------------------------------------------------------
    ConstruireIntervallesHorsDRC rs, intervalles, nbInt, J0, J5
    nbInt = FusionnerIntervalles(intervalles, nbInt)

    '-------------------------------------------------------
    ' 3. Brut de chaque phase
    '-------------------------------------------------------
    brutAD = JoursOuvrables(J0, J1)      ' RECEPTION      -> ATTRIBUTION
    brutAna = JoursOuvrables(J1, J2)     ' ATTRIBUTION    -> ENVOI SIGNATAIRE
    brutSig = JoursOuvrables(J2, J3)     ' ENVOI SIGN.    -> SIGNATURE
    brutTrame = JoursOuvrables(J3, J4)   ' SIGNATURE      -> ENVOI TRAME
    brutNotif = JoursOuvrables(J4, J5)   ' ENVOI TRAME    -> DATE FINALE

    '-------------------------------------------------------
    ' 4. Part hors DRC tombant dans chaque phase
    '-------------------------------------------------------
    hdAD = PartHorsDRC(intervalles, nbInt, J0, J1)
    hdAna = PartHorsDRC(intervalles, nbInt, J1, J2)
    hdSig = PartHorsDRC(intervalles, nbInt, J2, J3)
    hdTrame = PartHorsDRC(intervalles, nbInt, J3, J4)
    hdNotif = PartHorsDRC(intervalles, nbInt, J4, J5)

    '-------------------------------------------------------
    ' 5. Forfaits AUTRES J, plafonnes au disponible
    '-------------------------------------------------------
    ' Un forfait ne peut retirer plus de jours que la phase
    ' n'en contient encore. Sans ce plafond, la phase serait
    ' ramenee a zero alors que HORS DRC recevrait la totalite
    ' du forfait : l'identite serait rompue.
    fftAna = Plafonner(LireForfait(rs, "AUTRES J MARCHE"), brutAna - hdAna)
    fftSig = Plafonner(LireForfait(rs, "AUTRES J SIGN MARCHE"), brutSig - hdSig)

    fftTrame = LireForfait(rs, "AUTRESJ DCEE/COO") _
             + LireForfait(rs, "AUTRESJ DGA") _
             + LireForfait(rs, "AUTRESJ DG") _
             + LireForfait(rs, "AUTRESJ RISQCRE") _
             + LireForfait(rs, "AUTRESJ CA") _
             + LireForfait(rs, "AUTRESJ AUTRESPCRU")
    fftTrame = Plafonner(fftTrame, brutTrame - hdTrame)

    '-------------------------------------------------------
    ' 6. Net par phase
    '-------------------------------------------------------
    netAD = brutAD - hdAD
    netAna = brutAna - hdAna - fftAna
    netSig = brutSig - hdSig - fftSig
    netTrame = brutTrame - hdTrame - fftTrame
    netNotif = brutNotif - hdNotif

    ' Rattachement du segment trame
    Select Case PHASE_TRAME
        Case "AD":           netAD = netAD + netTrame
        Case "NOTIFICATION": netNotif = netNotif + netTrame
        Case Else:           netAna = netAna + netTrame
    End Select

    '-------------------------------------------------------
    ' 7. Agregats
    '-------------------------------------------------------
    horsDRC = hdAD + hdAna + hdSig + hdTrame + hdNotif _
            + fftAna + fftSig + fftTrame

    totalDRC = netAD + netAna + netSig + netNotif
    totalGlobal = totalDRC + horsDRC

    '-------------------------------------------------------
    ' 8. Ecriture
    '-------------------------------------------------------
    rs.Edit
    rs![DELAIS AD] = netAD
    rs![DELAIS ANALYSTE] = netAna
    rs![DELAIS SIGNATAIRE] = netSig
    rs![DELAIS NOTIFICATION] = netNotif
    rs![DELAIS HORS DRC] = horsDRC
    rs![DELAIS DRC] = totalDRC
    rs![DELAIS TOTAL] = totalGlobal
    rs.Update

    '-------------------------------------------------------
    ' 9. Controle d'identite
    '-------------------------------------------------------
    ' TOTAL doit retomber exactement sur l'ecart RECEPTION ->
    ' DATE FINALE. Tout ecart signale une anomalie de donnees
    ' ou une regression du calcul.
    controle = JoursOuvrables(J0, J5)
    If totalGlobal <> controle Then
        Debug.Print "ECART IDENTITE  IDSG=" & IDSGValue & _
                    "  TOTAL=" & totalGlobal & "  attendu=" & controle
    End If

    If TRACE_CALCUL Then
        Debug.Print "--- IDSG " & IDSGValue & " ---"
        Debug.Print "  Jalons  " & J0 & " | " & J1 & " | " & J2 & _
                    " | " & J3 & " | " & J4 & " | " & J5
        Debug.Print "  AD     brut=" & brutAD & " hd=" & hdAD & " net=" & netAD
        Debug.Print "  ANA    brut=" & brutAna & " hd=" & hdAna & " fft=" & fftAna
        Debug.Print "  SIG    brut=" & brutSig & " hd=" & hdSig & " fft=" & fftSig
        Debug.Print "  TRAME  brut=" & brutTrame & " hd=" & hdTrame & " fft=" & fftTrame
        Debug.Print "  NOTIF  brut=" & brutNotif & " hd=" & hdNotif
        Debug.Print "  DRC=" & totalDRC & "  HORS=" & horsDRC & "  TOTAL=" & totalGlobal
    End If

Nettoyage:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub

GestionErreur:
    Debug.Print "DELAIS  IDSG=" & IDSGValue & _
                "  erreur " & Err.Number & " : " & Err.Description
    Resume Nettoyage
End Sub


'===========================================================
'  JALONS
'-----------------------------------------------------------
' Renvoie six jalons formant une suite croissante couvrant
' l'integralite de RECEPTION -> DATE FINALE.
'
' Un jalon manquant prend la valeur du jalon suivant : la phase
' correspondante vaut alors zero, sans creer de trou. Un jalon
' anterieur au precedent est ramene a celui-ci : une saisie
' incoherente ne peut pas produire de duree negative.
'===========================================================
Private Sub NormaliserJalons(rs As DAO.Recordset, _
                             ByRef J0 As Date, ByRef J1 As Date, _
                             ByRef J2 As Date, ByRef J3 As Date, _
                             ByRef J4 As Date, ByRef J5 As Date)

    ' J5 : date finale du processus
    If Not IsNull(rs![Date de notification]) Then
        J5 = CDate(rs![Date de notification])
    ElseIf Not IsNull(rs![DATE SIGNATURE FINAL]) Then
        J5 = CDate(rs![DATE SIGNATURE FINAL])
    ElseIf Not IsNull(rs![DATE SIGNATURE]) Then
        J5 = CDate(rs![DATE SIGNATURE])
    ElseIf Not IsNull(rs![DATE ENVOIE AU SIGNATAIRE]) Then
        J5 = CDate(rs![DATE ENVOIE AU SIGNATAIRE])
    Else
        J5 = Date                      ' dossier en cours
    End If

    ' Remontee : chaque jalon absent s'aligne sur le suivant
    J4 = JalonOuDefaut(rs, "DATE ENVOIE TRAME", J5)
    J3 = JalonOuDefaut(rs, "DATE SIGNATURE", J4)
    J2 = JalonOuDefaut(rs, "DATE ENVOIE AU SIGNATAIRE", J3)
    J1 = JalonOuDefaut(rs, "ATTRIBUTION", J2)
    J0 = JalonOuDefaut(rs, "RECEPTION", J1)

    ' Descente : monotonie croissante
    If J1 < J0 Then J1 = J0
    If J2 < J1 Then J2 = J1
    If J3 < J2 Then J3 = J2
    If J4 < J3 Then J4 = J3
    If J5 < J4 Then J5 = J4
End Sub


Private Function JalonOuDefaut(rs As DAO.Recordset, _
                               fieldName As String, _
                               ByVal valeurDefaut As Date) As Date
    On Error GoTo Absent
    If IsNull(rs(fieldName)) Then
        JalonOuDefaut = valeurDefaut
    ElseIf Not IsDate(rs(fieldName)) Then
        JalonOuDefaut = valeurDefaut
    Else
        JalonOuDefaut = CDate(rs(fieldName))
    End If
    Exit Function
Absent:
    JalonOuDefaut = valeurDefaut
End Function


'===========================================================
'  PERIODES HORS DRC
'-----------------------------------------------------------
' Chaque periode est bornee a [J0 ; J5] des sa construction :
' un aller-retour marche qui deborderait de la fenetre du
' dossier ne peut ainsi pas retirer plus de jours qu'il n'en
' existe reellement.
'===========================================================
Private Sub ConstruireIntervallesHorsDRC(rs As DAO.Recordset, _
                                         ByRef intervalles() As Variant, _
                                         ByRef nbInt As Long, _
                                         ByVal borneMin As Date, _
                                         ByVal borneMax As Date)
    nbInt = 0
    ReDim intervalles(1 To 1)

    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE 1 MARCHE", "RETOUR 1 MARCHE", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE 2 MARCHE", "RETOUR 2 MARCHE", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE 1 SIGN MARCHE", "RETOUR 1 SIGN MARCHE", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE 2 SIGN MARCHE", "RETOUR 2 SIGN MARCHE", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE DCE/COO", "RETOUR DCE/COO", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE DGA", "RETOUR DGA", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE DG", "RETOUR DG", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE RISQCRE", "RETOUR RISQCRE", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE CA", "RETOUR CA", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE CORISQ", "RETOUR CORISQ", borneMin, borneMax
    AjouterIntervalle intervalles, nbInt, rs, "ENVOIE AUTRESPCRU", "RETOUR AUTRESPCRU", borneMin, borneMax
End Sub


Private Sub AjouterIntervalle(ByRef intervalles() As Variant, _
                              ByRef nbInt As Long, _
                              rs As DAO.Recordset, _
                              champEnvoi As String, _
                              champRetour As String, _
                              ByVal borneMin As Date, _
                              ByVal borneMax As Date)
    On Error GoTo Absent

    Dim d As Date, f As Date

    If IsNull(rs(champEnvoi)) Then Exit Sub
    If Not IsDate(rs(champEnvoi)) Then Exit Sub
    d = CDate(rs(champEnvoi))

    ' Periode non close : le decompte court jusqu'a la borne
    If IsNull(rs(champRetour)) Then
        f = borneMax
    ElseIf Not IsDate(rs(champRetour)) Then
        f = borneMax
    Else
        f = CDate(rs(champRetour))
    End If

    ' Bornage sur la fenetre du dossier
    If d < borneMin Then d = borneMin
    If f > borneMax Then f = borneMax
    If f <= d Then Exit Sub

    nbInt = nbInt + 1
    ReDim Preserve intervalles(1 To nbInt)
    intervalles(nbInt) = Array(d, f)
    Exit Sub

Absent:
    ' Champ inexistant dans la table : periode ignoree
End Sub


'===========================================================
'  FUSION DES CHEVAUCHEMENTS
'-----------------------------------------------------------
' Tri par date de debut puis balayage lineaire. Aucune ecriture
' sur un sous-element de Variant, ce que VBA n'autorise pas.
'===========================================================
Private Function FusionnerIntervalles(ByRef intervalles() As Variant, _
                                      ByVal nbInt As Long) As Long
    Dim i As Long, n As Long
    Dim debutCourant As Date, finCourante As Date
    Dim d As Date, f As Date
    Dim resultat() As Variant

    FusionnerIntervalles = 0
    If nbInt < 1 Then Exit Function

    TrierParDebut intervalles, nbInt

    ReDim resultat(1 To nbInt)
    n = 0
    debutCourant = CDate(intervalles(1)(0))
    finCourante = CDate(intervalles(1)(1))

    For i = 2 To nbInt
        d = CDate(intervalles(i)(0))
        f = CDate(intervalles(i)(1))
        If d <= finCourante Then
            If f > finCourante Then finCourante = f
        Else
            n = n + 1
            resultat(n) = Array(debutCourant, finCourante)
            debutCourant = d
            finCourante = f
        End If
    Next i

    n = n + 1
    resultat(n) = Array(debutCourant, finCourante)

    ReDim intervalles(1 To n)
    For i = 1 To n
        intervalles(i) = resultat(i)
    Next i

    FusionnerIntervalles = n
End Function


Private Sub TrierParDebut(ByRef intervalles() As Variant, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = 1 To n - 1
        For j = 1 To n - i
            If CDate(intervalles(j)(0)) > CDate(intervalles(j + 1)(0)) Then
                tmp = intervalles(j)
                intervalles(j) = intervalles(j + 1)
                intervalles(j + 1) = tmp
            End If
        Next j
    Next i
End Sub


'===========================================================
'  PART HORS DRC CONTENUE DANS UNE PHASE
'-----------------------------------------------------------
' Somme des jours ouvrables des intersections entre la phase
' [phaseDebut ; phaseFin] et les periodes hors DRC fusionnees.
'
' Les intervalles etant disjoints apres fusion, et les phases
' etant contigues, la somme sur toutes les phases redonne
' exactement le hors DRC total : l'identite est preservee.
'===========================================================
Private Function PartHorsDRC(ByRef intervalles() As Variant, _
                             ByVal nbInt As Long, _
                             ByVal phaseDebut As Date, _
                             ByVal phaseFin As Date) As Long
    Dim i As Long
    Dim d As Date, f As Date
    Dim total As Long

    PartHorsDRC = 0
    If nbInt < 1 Then Exit Function
    If phaseFin <= phaseDebut Then Exit Function

    total = 0
    For i = 1 To nbInt
        d = CDate(intervalles(i)(0))
        f = CDate(intervalles(i)(1))
        If d < phaseDebut Then d = phaseDebut
        If f > phaseFin Then f = phaseFin
        If f > d Then total = total + JoursOuvrables(d, f)
    Next i

    PartHorsDRC = total
End Function


'===========================================================
'  UTILITAIRES
'===========================================================

Private Function GetCritere(IDSGValue As String) As String
    If IsNumeric(IDSGValue) Then
        GetCritere = "IDSG=" & IDSGValue
    Else
        GetCritere = "IDSG='" & Replace(IDSGValue, "'", "''") & "'"
    End If
End Function


' Jours ouvrables strictement posterieurs a Debut, jusqu'a Fin
' incluse. Cette convention rend le calcul additif : pour trois
' jalons a <= b <= c, JoursOuvrables(a,b) + JoursOuvrables(b,c)
' egale JoursOuvrables(a,c). C'est ce qui garantit l'identite.
'
' Les jours feries ivoiriens ne sont pas pris en compte.
Public Function JoursOuvrables(Debut As Variant, Fin As Variant) As Long
    Dim d As Date
    Dim c As Long

    JoursOuvrables = 0
    If IsNull(Debut) Or IsNull(Fin) Then Exit Function
    If Not IsDate(Debut) Or Not IsDate(Fin) Then Exit Function
    If CDate(Fin) <= CDate(Debut) Then Exit Function

    c = 0
    For d = DateAdd("d", 1, CDate(Debut)) To CDate(Fin)
        If Weekday(d, vbMonday) <= 5 Then c = c + 1
    Next d

    JoursOuvrables = c
End Function


Private Function LireForfait(rs As DAO.Recordset, fieldName As String) As Long
    On Error GoTo Absent
    LireForfait = CLng(Nz(rs(fieldName), 0))
    If LireForfait < 0 Then LireForfait = 0
    Exit Function
Absent:
    LireForfait = 0
End Function


Private Function Plafonner(ByVal valeur As Long, ByVal maximum As Long) As Long
    If maximum < 0 Then maximum = 0
    If valeur < 0 Then valeur = 0
    If valeur > maximum Then
        Plafonner = maximum
    Else
        Plafonner = valeur
    End If
End Function


'###########################################################
'#  RECALCUL DE L'HISTORIQUE
'#
'#  A lancer depuis la fenetre Execution (Ctrl+G) :
'#      RecalculerTousLesDossiers
'#
'#  SAUVEGARDER LA BASE AVANT.
'###########################################################
Public Sub RecalculerTousLesDossiers()

    Dim rs As DAO.Recordset
    Dim n As Long, debut As Single

    If MsgBox("Recalculer les delais de TOUS les dossiers ?" & vbCrLf & vbCrLf & _
              "Cette operation ecrase les delais existants." & vbCrLf & _
              "Une sauvegarde prealable est indispensable.", _
              vbExclamation + vbYesNo, "Recalcul complet") <> vbYes Then Exit Sub

    debut = Timer
    Set rs = CurrentDb.OpenRecordset( _
             "SELECT [IDSG] FROM [DATABASE] WHERE [IDSG] Is Not Null;", dbOpenSnapshot)

    Do While Not rs.EOF
        DELAIS CStr(rs![IDSG])
        n = n + 1
        If n Mod 50 = 0 Then
            SysCmd acSysCmdSetStatus, "Recalcul en cours : " & n & " dossiers..."
            DoEvents
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    SysCmd acSysCmdClearStatus

    MsgBox n & " dossiers recalcules en " & Format(Timer - debut, "0.0") & " s." & vbCrLf & _
           "Verifier la fenetre Execution : tout ecart d'identite y est signale.", _
           vbInformation
End Sub
