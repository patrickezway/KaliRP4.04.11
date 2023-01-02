Attribute VB_Name = "MSaisie"
Option Explicit

'Type de saisie g�r�e
Public Const SAIS_TYP_TOUT_CAR = 0
Public Const SAIS_TYP_JOUR_SEMAINE = 1
Public Const SAIS_TYP_HEURE = 2
Public Const SAIS_TYP_DATE = 3
Public Const SAIS_TYP_ENTIER = 4
Public Const SAIS_TYP_LETTRE = 5
Public Const SAIS_TYP_LETTRE_PONCT = 6
Public Const SAIS_TYP_ENTIER_NEG = 7
Public Const SAIS_TYP_DATNAIS = 8
Public Const SAIS_TYP_CAR_PARTICULIER = 9
Public Const SAIS_TYP_PRIX = 10
Public Const SAIS_TYP_PERIODE = 11
Public Const SAIS_TYP_CODE = 12
Public Const SAIS_TYP_CHOIXLISTE = 13

'Conversions possibles
Public Const SAIS_CONV_MINUSCULE = 1
Public Const SAIS_CONV_MAJUSCULE = 2
Public Const SAIS_CONV_SECRET = 3

' indique s'il y a des champs � saisir ou seulement des boutons
Public BOOL_YA_DES_CHAMPS As Boolean

'Retour possible
Public Const SAIS_RET_NOMODIF = -1
Public Const SAIS_RET_MODIF = 0

Public Type SAIS_SPRMFRM
    titre As String
    nomhelp As String
    visu_oblig As Boolean
    x As Integer
    y As Integer
    max_nbcar_visible As Integer    '0 => la zone texte = � la taille du texte le plus grand
    reste_charg�e As Boolean
End Type

'Structure permettant l'appel � la form FSAIS_
Public Type SAIS_SCHAMP
    libelle As String
    len As Integer  'Longueur du texte � saisir
    type As Integer
    chaine_type As String
    facu As Boolean 'False => si OK ce champ doit �tre rempli
    conversion As Integer
    sval As String  'Contenu de la zone texte au retour
    validationdirecte As Boolean
    liste_nomtable As String
    liste_multiselect As Boolean
    liste_chpretour As String
    liste_chpnum As String
End Type

'Structure contenant les lignes pour les champs listes
Public Type SAIS_SCHPLISTE
    Liste_Num As Integer
    Item_Num As Integer
    Item_code_retour As String
    Item_LaStr As String
    Item_bSel As Boolean
End Type

Public Type SAIS_SBOUTON
    libelle As String
    image As String
    raccourci_alt As Integer
    raccourci_touche As Integer
    largeur As Long
End Type

Public Type SAIS_SSAISIE
    prmfrm As SAIS_SPRMFRM
    champs() As SAIS_SCHAMP
    item_liste() As SAIS_SCHPLISTE
    boutons() As SAIS_SBOUTON
    retour As Integer
End Type
    
Public SAIS_Saisie As SAIS_SSAISIE

Private Function ctrl_date(ByRef vr_str As String) As Boolean

    Dim stmp As String, si�cle_en_cours As String, sdater As String, s As String
    Dim jj As Integer, mm As Integer, AA As Integer, pos As Integer
    Dim nbj As Integer
    
    If left$(vr_str, 1) = "j" Or left$(vr_str, 1) = "J" Then
        If Len(vr_str) = 1 Then
            nbj = 0
        ElseIf Mid$(vr_str, 2, 1) = "-" Then
            If Len(vr_str) > 6 Then
                GoTo lab_fin_date
            End If
            If Not STR_EstEntierPos(Mid$(vr_str, 3)) Then
                GoTo lab_fin_date
            End If
            nbj = -(CInt(Mid$(vr_str, 3)))
        ElseIf Mid$(vr_str, 2, 1) = "+" Then
            If Len(vr_str) > 6 Then
                GoTo lab_fin_date
            End If
            If Not STR_EstEntierPos(Mid$(vr_str, 3)) Then
                GoTo lab_fin_date
            End If
            nbj = CInt(Mid$(vr_str, 3))
        Else
            GoTo lab_fin_date
        End If
        vr_str = Format(Date + nbj, "dd/mm/yyyy")
        ctrl_date = True
        Exit Function
    End If
    
    If left$(vr_str, 1) = "m" Or left$(vr_str, 1) = "M" Then
        If Len(vr_str) = 1 Then
            s = Date
            s = "01/" & Mid$(s, 4)
        ElseIf Mid$(vr_str, 2, 1) = "-" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                ' A FAIRE ...
            Else
            End If
            GoTo lab_fin_date
        ElseIf Mid$(vr_str, 2, 1) = "+" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                ' A FAIRE ...
            Else
            End If
            GoTo lab_fin_date
        Else
            GoTo lab_fin_date
        End If
        vr_str = Format(CDate(s), "dd/mm/yyyy")
        ctrl_date = True
        Exit Function
    End If
    
    stmp = Format(Date, "dd/mm/yyyy")
    si�cle_en_cours = Mid(stmp, 7, 2)
    
    If STR_EstEntierPos(vr_str) Then
        If Len(vr_str) = 6 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + si�cle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 8 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + Right$(vr_str, 4)
        Else
            sdater = ""
        End If
    Else
        If Not IsDate(vr_str) Then
            sdater = ""
            GoTo lab_fin_date
        End If
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        jj = CInt(Mid$(vr_str, 1, pos - 1))
        vr_str = Mid$(vr_str, pos + 1)
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        mm = CInt(Mid$(vr_str, 1, pos - 1))
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        AA = CInt(Mid$(vr_str, pos + 1))
        sdater = Format(jj, "00") + "/" + Format(mm, "00") + "/"
        If AA < 100 Then
            sdater = sdater + si�cle_en_cours + Format(AA, "00")
        Else
            sdater = sdater + Format(AA, "0000")
        End If
    End If
    
lab_fin_date:
    If sdater = "" Or Not IsDate(sdater) Then
        MsgBox "La saisie ne correspond pas � une date.", vbOKOnly + vbExclamation, "SAIS_ Erronn�e"
        ctrl_date = False
        Exit Function
    End If
    
    vr_str = sdater
    ctrl_date = True

End Function

Private Function ctrl_entier_pos(ByVal v_str As String) As String

    If Not STR_EstEntierPos(v_str) Then
        MsgBox "La saisie ne correspond pas � un nombre positif.", vbOKOnly + vbExclamation, "SAIS_ Erronn�e"
        ctrl_entier_pos = False
        Exit Function
    End If
    ctrl_entier_pos = True
    
End Function

Private Function ctrl_heure(ByRef vr_str As String) As Boolean

    Dim HH As Integer, mm As Integer, pos As Integer
    Dim s As String
    
    If vr_str = "h" Or vr_str = "H" Then
        vr_str = Format(Time, "hh:mm")
        ctrl_heure = True
        Exit Function
    End If
    
    HH = -1
    If STR_EstEntierPos(vr_str) And Len(vr_str) <= 4 Then
        If Len(vr_str) <= 2 Then
            HH = val(vr_str)
            mm = 0
        Else
            HH = val(vr_str) / 100
            mm = val(vr_str) Mod 100
        End If
    Else
        pos = InStr(vr_str, ":")
        If pos > 0 Then
            s = Mid$(vr_str, pos + 1)
            If InStr(s, ":") <= 0 Then
                HH = val(left$(vr_str, pos - 1))
                mm = val(Mid$(vr_str, pos + 1))
            End If
        End If
    End If
    If HH > 24 Then
        HH = -1
    ElseIf mm > 59 Then
        HH = -1
    ElseIf (HH * 100) + mm > 2400 Then
        HH = -1
    End If
    If HH >= 0 Then
        s = ""
        If HH < 10 Then
            s = "0"
        End If
        s = s + Trim$(str(HH)) + ":"
        If mm < 10 Then
            s = s + "0"
        End If
        vr_str = s + Trim$(str(mm))
        ctrl_heure = True
        Exit Function
    End If
    
    MsgBox "La saisie ne correspond pas � une heure.", vbOKOnly + vbExclamation, "SAIS_ Erronn�e"
    ctrl_heure = False

End Function

Public Sub SAIS_AddBouton(ByVal v_libelle As String, _
                          ByVal v_image As String, _
                          ByVal v_rcalt As Integer, _
                          ByVal v_rctouche As Integer, _
                          ByVal v_largeur As Integer)
    
    Dim n As Integer
    
    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.boutons)
    On Error GoTo 0
    
    n = n + 1
    ReDim Preserve SAIS_Saisie.boutons(n)
    SAIS_Saisie.boutons(n).libelle = v_libelle
    SAIS_Saisie.boutons(n).image = v_image
    SAIS_Saisie.boutons(n).raccourci_alt = v_rcalt
    SAIS_Saisie.boutons(n).raccourci_touche = v_rctouche
    SAIS_Saisie.boutons(n).largeur = v_largeur

End Sub

Public Sub SAIS_AddChamp(ByVal v_libelle As String, _
                         ByVal v_len As Integer, _
                         ByVal v_type As Integer, _
                         ByVal v_facu As Boolean, _
                         Optional v_sval As Variant)
    
    Dim n As Integer
    
    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.champs)
    On Error GoTo 0
    
    
    BOOL_YA_DES_CHAMPS = True
    n = n + 1
    ReDim Preserve SAIS_Saisie.champs(n)
    SAIS_Saisie.champs(n).libelle = v_libelle
    SAIS_Saisie.champs(n).len = v_len
    SAIS_Saisie.champs(n).type = v_type
    SAIS_Saisie.champs(n).facu = v_facu
    If Not IsMissing(v_sval) Then
        SAIS_Saisie.champs(n).sval = v_sval
    Else
        SAIS_Saisie.champs(n).sval = ""
    End If
    SAIS_Saisie.champs(n).validationdirecte = False
    
End Sub

Public Sub SAIS_AddItemListe(ByVal v_liste_num As Integer, _
                             ByVal v_Item_Num As Integer, _
                             ByVal v_Item_code_retour As String, _
                             ByVal v_Item_LaStr As String, _
                             ByVal v_Item_bSel As Boolean)
                             
    Dim n As Integer
    
    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.item_liste)
    On Error GoTo 0
    
    n = n + 1
    ReDim Preserve SAIS_Saisie.item_liste(n)
    SAIS_Saisie.item_liste(n).Liste_Num = v_liste_num
    SAIS_Saisie.item_liste(n).Item_Num = v_Item_Num
    SAIS_Saisie.item_liste(n).Item_code_retour = v_Item_code_retour
    SAIS_Saisie.item_liste(n).Item_LaStr = v_Item_LaStr
    SAIS_Saisie.item_liste(n).Item_bSel = v_Item_bSel

End Sub

Public Sub SAIS_AddListe(ByVal v_libelle As String, _
                         ByVal v_liste_nomtable As String, _
                         ByVal v_liste_multiselect As Boolean, _
                         ByRef r_liste_chpretour As String, _
                         ByVal v_liste_chpnum As String, _
                         ByVal v_type As Integer, _
                         ByVal v_facu As Boolean, _
                         Optional v_sval As Variant)
    
    Dim n As Integer
    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.champs)
    On Error GoTo 0
    
    n = n + 1
    ReDim Preserve SAIS_Saisie.champs(n)
    SAIS_Saisie.champs(n).libelle = v_libelle
    SAIS_Saisie.champs(n).liste_nomtable = v_liste_nomtable
    SAIS_Saisie.champs(n).liste_multiselect = v_liste_multiselect
    SAIS_Saisie.champs(n).liste_chpretour = r_liste_chpretour
    SAIS_Saisie.champs(n).liste_chpnum = v_liste_chpnum
    SAIS_Saisie.champs(n).type = v_type
    SAIS_Saisie.champs(n).facu = v_facu
    If Not IsMissing(v_sval) Then
        SAIS_Saisie.champs(n).sval = v_sval
    Else
        SAIS_Saisie.champs(n).sval = ""
    End If
    SAIS_Saisie.champs(n).validationdirecte = False
    
End Sub

Public Sub SAIS_AddChampComplet(ByVal v_libelle As String, _
                                ByVal v_len As Integer, _
                                ByVal v_type As Integer, _
                                ByVal v_str_type As String, _
                                ByVal v_facu As Boolean, _
                                ByVal v_conv As Integer, _
                                ByVal v_valid_direct As Boolean, _
                                Optional v_sval As Variant)
    
    Dim n As Integer
    
    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.champs)
    On Error GoTo 0
    
    BOOL_YA_DES_CHAMPS = True
    n = n + 1
    ReDim Preserve SAIS_Saisie.champs(n)
    SAIS_Saisie.champs(n).libelle = v_libelle
    SAIS_Saisie.champs(n).len = v_len
    SAIS_Saisie.champs(n).type = v_type
    SAIS_Saisie.champs(n).chaine_type = v_str_type
    SAIS_Saisie.champs(n).facu = v_facu
    SAIS_Saisie.champs(n).conversion = v_conv
    SAIS_Saisie.champs(n).validationdirecte = v_valid_direct
    If Not IsMissing(v_sval) Then
        SAIS_Saisie.champs(n).sval = v_sval
    Else
        SAIS_Saisie.champs(n).sval = ""
    End If
    
End Sub

Public Function SAIS_CtrlChamp(ByRef vr_str As String, _
                               ByVal v_typchamp As Integer) As Boolean
                                
    Dim s As String, s2 As String, sc As String
    Dim fok As Boolean
    Dim pos As Integer, n As Integer, I As Integer
    
    Select Case v_typchamp
    Case SAIS_TYP_JOUR_SEMAINE
        Select Case LCase(left$(vr_str, 1))
        Case "l"
            vr_str = "lundi"
            SAIS_CtrlChamp = True
        Case "ma"
            vr_str = "mardi"
            SAIS_CtrlChamp = True
        Case "me"
            vr_str = "mercredi"
            SAIS_CtrlChamp = True
        Case "j"
            vr_str = "jeudi"
            SAIS_CtrlChamp = True
        Case "v"
            vr_str = "vendredi"
            SAIS_CtrlChamp = True
        Case Else
            MsgBox "La saisie ne correspond pas � un jour de la semaine", vbOKOnly + vbExclamation, "Saisie Erronn�e"
            SAIS_CtrlChamp = False
        End Select
        Exit Function
    Case SAIS_TYP_PERIODE
        vr_str = UCase(vr_str)
        s = Right$(vr_str, 1)
        If s = "J" Or s = "S" Or s = "M" Or s = "A" Then
            s2 = left$(vr_str, Len(vr_str) - 1)
            If STR_EstEntierPos(s2) Then
                n = CInt(s2)
                vr_str = n & s
                SAIS_CtrlChamp = True
                Exit Function
            End If
        End If
        MsgBox "La saisie ne correspond pas � une p�riode : nombre suivi de J(ours)/S(emaines)/M(ois)/A(nn�es).", vbOKOnly + vbExclamation, "Saisie Erronn�e"
        SAIS_CtrlChamp = False
        Exit Function
    Case SAIS_TYP_HEURE
        SAIS_CtrlChamp = ctrl_heure(vr_str)
        Exit Function
    Case SAIS_TYP_DATE
        SAIS_CtrlChamp = ctrl_date(vr_str)
        Exit Function
    Case SAIS_TYP_ENTIER_NEG
        If InStr(vr_str, "-") > 1 Then
            MsgBox "La saisie ne correspond pas � un nombre sign�.", vbOKOnly + vbExclamation, "Saisie Erronn�e"
            SAIS_CtrlChamp = False
            Exit Function
        End If
        If left$(vr_str, 1) = "-" And InStr(Mid$(vr_str, 2), "-") > 0 Then
            MsgBox "La saisie ne correspond pas � un nombre sign�.", vbOKOnly + vbExclamation, "Saisie Erronn�e"
            SAIS_CtrlChamp = False
            Exit Function
        End If
        SAIS_CtrlChamp = True
        Exit Function
    Case SAIS_TYP_ENTIER
        SAIS_CtrlChamp = ctrl_entier_pos(vr_str)
        Exit Function
    Case SAIS_TYP_DATNAIS
        SAIS_CtrlChamp = ctrl_ddn(vr_str)
        Exit Function
    Case SAIS_TYP_PRIX
        SAIS_CtrlChamp = ctrl_prix(vr_str)
        Exit Function
    Case SAIS_TYP_CODE
        For I = 1 To Len(vr_str)
            sc = Mid$(vr_str, I, 1)
            If sc >= "A" And sc <= "Z" Then
                fok = True
            ElseIf sc >= "a" And sc <= "z" Then
                fok = True
            ElseIf sc >= "0" And sc <= "9" Then
                fok = True
            ElseIf sc = "-" Or sc = "_" Or sc = "." Then
                fok = True
            Else
                fok = False
            End If
            If Not fok Then
                MsgBox "Seuls les chiffres, lettres (sauf les caract�res accentu�s) et - _ . sont autoris�s.", vbOKOnly + vbExclamation, "Saisie Erronn�e"
                SAIS_CtrlChamp = False
                Exit Function
            End If
        Next I
        SAIS_CtrlChamp = True
        Exit Function
    Case Else
        SAIS_CtrlChamp = True
    End Select

End Function

Private Function ctrl_ddn(ByRef vr_str As String) As Boolean

    Dim stmp As String, si�cle_en_cours As String, sdater As String
    Dim ddn As Date
    
    stmp = Format(Date, "dd/mm/yyyy")
    si�cle_en_cours = Mid$(stmp, 7, 2)
    
    If STR_EstEntierPos(vr_str) Then
        If Len(vr_str) = 6 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + si�cle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 8 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + Right$(vr_str, 4)
        Else
            sdater = ""
        End If
    ElseIf STR_EstEntierPos(left$(vr_str, 2)) And STR_EstEntierPos(Mid$(vr_str, 4, 2)) And STR_EstEntierPos(Mid$(vr_str, 7)) And Mid$(vr_str, 3, 1) = "/" And Mid$(vr_str, 6, 1) = "/" Then
        If Len(vr_str) = 8 Then
            sdater = left$(vr_str, 6) + si�cle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 10 Then
            sdater = vr_str
        Else
            sdater = ""
        End If
    End If
    
    If sdater = "" Or Not IsDate(sdater) Then
        MsgBox "La saisie ne correspond pas � une date.", vbOKOnly + vbExclamation, "SAIS_ Erron�e"
        ctrl_ddn = False
        Exit Function
    End If
    
    ddn = CDate(sdater)
    If ddn > Date Then
        MsgBox "Ce malade n'est pas encore n�.", vbOKOnly + vbExclamation, "SAIS_ Erron�e"
        ctrl_ddn = False
        Exit Function
    End If
    
    vr_str = sdater
    ctrl_ddn = True

End Function

Private Function ctrl_prix(ByRef vr_str As String) As Boolean

    Dim prix As Double
    
    On Error GoTo err_prix
    prix = CDbl(vr_str)
    On Error GoTo 0
    vr_str = STR_Prix(vr_str)
    ctrl_prix = True
    Exit Function
    
err_prix:
    MsgBox "La saisie ne correspond pas � un prix.", vbOKOnly + vbExclamation, "SAIS_ Erron�e"
    ctrl_prix = False
    
End Function

Public Sub SAIS_Init()

    BOOL_YA_DES_CHAMPS = False
    
    SAIS_Saisie.prmfrm.visu_oblig = True
    SAIS_Saisie.prmfrm.titre = ""
    SAIS_Saisie.prmfrm.nomhelp = ""
    SAIS_Saisie.prmfrm.x = 0
    SAIS_Saisie.prmfrm.y = 0
    SAIS_Saisie.prmfrm.max_nbcar_visible = 50
    SAIS_Saisie.prmfrm.reste_charg�e = False
    
    Erase SAIS_Saisie.champs()
    
    Erase SAIS_Saisie.boutons()
    
    Erase SAIS_Saisie.item_liste()
    
End Sub

Public Sub SAIS_InitOblig(ByVal v_oblig As Boolean)

    SAIS_Saisie.prmfrm.visu_oblig = v_oblig

End Sub

Public Sub SAIS_InitPos(ByVal v_posx As Long, _
                        ByVal v_posy As Long)

    SAIS_Saisie.prmfrm.x = v_posx
    SAIS_Saisie.prmfrm.y = v_posy

End Sub

Public Sub SAIS_InitResteCharg�e(ByVal v_restec As Boolean)

    SAIS_Saisie.prmfrm.reste_charg�e = v_restec

End Sub

Public Sub SAIS_InitTitreHelp(ByVal v_nomtitre As String, _
                              ByVal v_nomhelp As String)

    SAIS_Saisie.prmfrm.titre = v_nomtitre
    SAIS_Saisie.prmfrm.nomhelp = v_nomhelp

End Sub



