'###########################################################
'#  CORRECTIONS - POINT 3
'#  Module standard  DELAIS_
'#
'#  Quatre interventions :
'#    3.1  CalculerDelaiSansChevauchement  -> REMPLACER
'#    3.2  TrierIntervallesParDebut        -> AJOUTER (nouveau)
'#    3.3  AjouterIntervalle               -> REMPLACER
'#    3.4  DELAIS_HORS_DRC                 -> retirer une ligne
'#
'#  Les autres procedures du module ne sont pas concernees.
'###########################################################


'===========================================================
'  3.1  REMPLACER  CalculerDelaiSansChevauchement
'===========================================================
'
' DEFAUTS DE LA VERSION PRECEDENTE
'
'   a) intervalles(i)(0) = DateMin(...)
'      VBA autorise la LECTURE d'un element imbrique dans un
'      Variant contenant un tableau, mais pas son ECRITURE.
'      Erreur de compilation, ou 451 a l'execution.
'
'   b) Fusion incomplete. Quand l'intervalle i absorbe
'      l'intervalle j, sa borne de fin recule. Il peut alors
'      chevaucher un intervalle compris entre i+1 et j-1, deja
'      compare a l'ancienne borne et jamais reexamine. Le
'      chevauchement subsiste et le delai est surevalue.
'
'   c) ReDim Preserve a chaque fusion, dans une double boucle.
'      Inutile puisque nbIntervalles porte deja le compte.
'
' PRINCIPE RETENU
'
'   Trier les intervalles par date de debut, puis les parcourir
'   une seule fois en maintenant un intervalle courant. Chaque
'   nouvel intervalle soit prolonge le courant, soit ouvre le
'   suivant. Aucune ecriture imbriquee n'est necessaire, et la
'   fusion est complete par construction.

Private Function CalculerDelaiSansChevauchement( _
                        ByRef intervalles() As Variant, _
                        ByVal nbIntervalles As Long) As Long

    Dim i As Long
    Dim total As Long
    Dim debutCourant As Date, finCourante As Date
    Dim d As Date, f As Date

    CalculerDelaiSansChevauchement = 0
    If nbIntervalles < 1 Then Exit Function

    ' Tri croissant sur la date de debut
    TrierIntervallesParDebut intervalles, nbIntervalles

    ' Initialisation sur le premier intervalle
    debutCourant = CDate(intervalles(1)(0))
    finCourante = CDate(intervalles(1)(1))
    total = 0

    For i = 2 To nbIntervalles
        d = CDate(intervalles(i)(0))
        f = CDate(intervalles(i)(1))

        If d <= finCourante Then
            ' Chevauchement ou contiguite : on prolonge
            If f > finCourante Then finCourante = f
        Else
            ' Rupture : on solde le bloc courant et on en ouvre un autre
            total = total + JoursOuvrables(debutCourant, finCourante)
            debutCourant = d
            finCourante = f
        End If
    Next i

    ' Solde du dernier bloc
    total = total + JoursOuvrables(debutCourant, finCourante)

    CalculerDelaiSansChevauchement = total
End Function


'===========================================================
'  3.2  AJOUTER  TrierIntervallesParDebut   (procedure nouvelle)
'===========================================================
'
' Tri a bulles. Le choix se justifie ici : DELAIS_HORS_DRC
' alimente au maximum onze intervalles, et un tri simple reste
' plus lisible qu'un tri rapide pour un volume aussi faible.
'
' L'echange passe par une variable intermediaire, ce qui est
' precisement ce que l'ancienne version ne faisait pas : on
' affecte l'element entier du tableau, jamais un sous-element.

Private Sub TrierIntervallesParDebut(ByRef intervalles() As Variant, _
                                     ByVal n As Long)
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
'  3.3  REMPLACER  AjouterIntervalle
'===========================================================
'
' CORRECTIONS
'   - suppression du label orphelin  SetNothing:
'     Aucun GoTo n'y menait ; il ne servait a rien et brouillait
'     la lecture.
'   - suppression du Debug.Print, qui saturait la fenetre
'     Execution a chaque recalcul.
'   - ajout d'un garde-fou : un intervalle dont le retour
'     precede l'envoi est ignore plutot que compte a zero, ce
'     qui evite qu'une saisie aberrante fausse le tri.

Private Sub AjouterIntervalle( _
    ByRef intervalles() As Variant, _
    ByRef nbIntervalles As Long, _
    ByVal DateEnvoi As Variant, _
    ByVal DateRetour As Variant)

    If IsNull(DateEnvoi) Then Exit Sub
    If Not IsDate(DateEnvoi) Then Exit Sub

    ' Dossier encore en cours : on arrete le decompte a aujourd'hui
    If IsNull(DateRetour) Then DateRetour = Date
    If Not IsDate(DateRetour) Then DateRetour = Date

    ' Saisie incoherente : intervalle ignore
    If CDate(DateRetour) < CDate(DateEnvoi) Then Exit Sub

    nbIntervalles = nbIntervalles + 1
    ReDim Preserve intervalles(1 To nbIntervalles)

    intervalles(nbIntervalles) = Array(CDate(DateEnvoi), CDate(DateRetour))
End Sub


'===========================================================
'  3.4  DELAIS_HORS_DRC  -  retirer la ligne dupliquee
'===========================================================
'
' Le corps contient :
'
'         nbIntervalles = 0
'         ReDim intervalles(1 To 1)
'             nbIntervalles = 0        <-- doublon, a supprimer
'
' Sans consequence fonctionnelle, mais l'indentation decalee
' laisse croire a un bloc conditionnel absent.
'
' Conserver :
'
'         nbIntervalles = 0
'         ReDim intervalles(1 To 1)


'###########################################################
'#  DateMin et DateMax
'#
'#  Ces deux fonctions ne sont plus appelees nulle part apres
'#  la reecriture. Les laisser ne cause aucun tort ; si tu
'#  preferes un module propre, elles peuvent etre supprimees
'#  apres avoir verifie qu'aucun autre module ne les utilise
'#  (Ctrl+F, "Tout le projet").
'###########################################################


'###########################################################
'#  DEUX POINTS RELEVES AU PASSAGE, HORS PERIMETRE
'#
'#  a) Fermeture des recordsets
'#     Les procedures du module font  rs.Close  sans
'#     Set rs = Nothing, et sans gestion d'erreur. Si une
'#     erreur survient entre OpenRecordset et Close, le
'#     recordset reste ouvert. Sur un recalcul de masse,
'#     cela finit par saturer les ressources DAO.
'#
'#  b) JoursOuvrables
'#     La boucle demarre a DateAdd("d", 1, Debut) : le jour
'#     d'envoi n'est pas compte, seul le lendemain l'est. C'est
'#     une convention defendable, mais verifie qu'elle
'#     correspond bien a celle attendue par la DRC pour le
'#     reporting des delais. Elle ne tient pas compte non plus
'#     des jours feries ivoiriens, ce qui allonge mecaniquement
'#     les delais affiches autour des periodes de fete.
'###########################################################
