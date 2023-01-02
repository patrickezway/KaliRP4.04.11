VERSION 5.00
Begin VB.Form Saisie 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSaisie 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton ComboCmd 
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox ComboCmb 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtSaisie 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblOblig 
         BackColor       =   &H00C0C0C0&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   645
         Width           =   135
      End
      Begin VB.Label lblSaisie 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2340
      Width           =   5535
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "Saisie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nbbouton As Integer

Private g_form_active As Boolean
Private g_ind_txt_enabled As Integer

Private Function ctrl_saisie_ok(ByVal v_index As Integer, _
                                ByVal v_lafin As Boolean)

    Dim HH As Integer, mm As Integer, pos As Integer
    Dim stmp As String, texte As String, s As String, s2 As String
    
    If Len(txtSaisie(v_index)) = 0 Then
        If Not v_lafin Then
            ctrl_saisie_ok = True
        Else
            If SAIS_Saisie.champs(v_index).facu Then
                ctrl_saisie_ok = True
            Else
                MsgBox "La saisie de cette rubrique est obligatoire", vbOKOnly + vbExclamation, "Saisie Erronn�e"
                ctrl_saisie_ok = False
            End If
        End If
        Exit Function
    End If
        
    texte = txtSaisie(v_index).Text
    If Not SAIS_CtrlChamp(texte, SAIS_Saisie.champs(v_index).type) Then
        txtSaisie(v_index).Text = ""
        ctrl_saisie_ok = False
    Else
        txtSaisie(v_index).Text = texte
        ctrl_saisie_ok = True
    End If
    
End Function

Private Sub initialiser()

    Dim hauteur As Integer, nb_champ As Integer, Index As Integer, I As Integer
    Dim max_nbcar_txt As Integer, max_size_lbl As Integer, nb_car As Integer
    Dim max_size_txt As Integer, marge As Integer, intervalle As Integer
    Dim size_txt As Integer, IndexText As Integer, IndexListe As Integer
    Dim lg_titre As Long, lg_texte As Long, lg_bouton As Long, lg_bouton1 As Long
    Dim lg As Long, lg_tot As Long
    
    frmSaisie.Caption = SAIS_Saisie.prmfrm.titre
    lg_titre = FRM_LargeurTexte(Me, frmSaisie, SAIS_Saisie.prmfrm.titre) + 255

    'Hauteur de chaque intervalle
    hauteur = 325
    marge = 255
    
    'Nbre de caract�res du plus long libell�
    max_size_lbl = 0
    'Nbre de caract�res du plus long texte
    max_nbcar_txt = 0
    
    ' Textbox + labels
    nb_champ = 0
    On Error Resume Next
    nb_champ = UBound(SAIS_Saisie.champs) + 1
    On Error GoTo 0
    IndexText = 1
    IndexListe = 1
    Index = 1
    g_ind_txt_enabled = -1
    For I = 0 To nb_champ - 1
        If I > 0 Then
            Load lblSaisie(I)
            If SAIS_Saisie.champs(I).type = SAIS_TYP_CHOIXLISTE Then
                Load ComboCmb(I)
                If SAIS_Saisie.champs(I).liste_multiselect Then
                    Load ComboCmd(I)
                End If
            Else
                Load txtSaisie(I)
            End If
            Load lblOblig(I)
        End If
        If SAIS_Saisie.champs(I).type = SAIS_TYP_CHOIXLISTE Then
            lblSaisie(I).visible = True
            lblSaisie(I).Top = 325 + hauteur + ((I * 2) * hauteur)
            lblSaisie(I).Caption = SAIS_Saisie.champs(I).libelle
            lblSaisie(I).TabIndex = IndexText + IndexListe
            If SAIS_Saisie.prmfrm.visu_oblig Then
                If SAIS_Saisie.champs(I).facu = False Then
                    lblOblig(I).visible = True
                    lblOblig(I).Top = lblSaisie(I).Top
                Else
                    lblOblig(I).visible = False
                End If
            Else
                lblOblig(I).visible = False
            End If
            'IndexListe = IndexListe + 1
            If FRM_LargeurTexte(Me, lblSaisie(I), lblSaisie(I).Caption) > max_size_lbl Then
                max_size_lbl = FRM_LargeurTexte(Me, lblSaisie(I), lblSaisie(I).Caption)
            End If
            ComboCmb(I).visible = True
            ComboCmb(I).Top = lblSaisie(I).Top
            If SAIS_Saisie.champs(I).liste_multiselect Then ComboCmd(I).visible = True
            If SAIS_Saisie.champs(I).liste_multiselect Then ComboCmd(I).Top = lblSaisie(I).Top
            ' Remplir le combo
            Call Remplir_Combo(I)
            If SAIS_Saisie.champs(I).liste_multiselect Then
                ComboCmd(I).Caption = SAIS_Saisie.champs(I).libelle
                ComboCmd(I).width = 2 * FRM_LargeurTexte(Me, ComboCmd(I), ComboCmd(I).Caption)
            End If
            If nb_car > max_nbcar_txt Then
                max_nbcar_txt = nb_car
            End If
            ComboCmb(I).TabIndex = Index
            Index = Index + 1
        Else
            If SAIS_Saisie.champs(I).conversion = SAIS_CONV_SECRET Then
                txtSaisie(I).PasswordChar = "*"
            End If
            lblSaisie(I).visible = True
            lblSaisie(I).Top = 325 + hauteur + ((I * 2) * hauteur)
            lblSaisie(I).Caption = SAIS_Saisie.champs(I).libelle
            lblSaisie(I).TabIndex = IndexText
            If SAIS_Saisie.prmfrm.visu_oblig Then
                If SAIS_Saisie.champs(I).facu = False Then
                    lblOblig(I).visible = True
                    lblOblig(I).Top = lblSaisie(I).Top
                Else
                    lblOblig(I).visible = False
                End If
            Else
                lblOblig(I).visible = False
            End If
            'Index = Index + 1
            If FRM_LargeurTexte(Me, lblSaisie(I), lblSaisie(I).Caption) > max_size_lbl Then
                max_size_lbl = FRM_LargeurTexte(Me, lblSaisie(I), lblSaisie(I).Caption)
            End If
            txtSaisie(I).visible = True
            txtSaisie(I).Top = lblSaisie(I).Top
            nb_car = SAIS_Saisie.champs(I).len
            If nb_car > 0 Then
                If g_ind_txt_enabled = -1 Then
                    g_ind_txt_enabled = I
                End If
    '            txtSaisie(i).BackColor = &HFFFFFF
                txtSaisie(I).Enabled = True
            ElseIf nb_car < 0 Then
                'Texte non modifiable
    '            txtSaisie(i).BackColor = &HC0C0C0
                txtSaisie(I).Enabled = False
                nb_car = -nb_car
            Else
                txtSaisie(I).visible = False
            End If
            txtSaisie(I).MaxLength = nb_car
            txtSaisie(I).Text = SAIS_Saisie.champs(I).sval
            If nb_car > max_nbcar_txt Then
                max_nbcar_txt = nb_car
            End If
            txtSaisie(I).TabIndex = Index
            Index = Index + 2
        End If
    Next I
    
    'Nbre de caract�res max du texte
    If SAIS_Saisie.prmfrm.max_nbcar_visible > 0 And max_nbcar_txt > SAIS_Saisie.prmfrm.max_nbcar_visible Then
        max_nbcar_txt = SAIS_Saisie.prmfrm.max_nbcar_visible
    End If
    'Conversion de caract�res en pixels
    max_size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(max_nbcar_txt, "M"))
    lg_texte = 255 + max_size_lbl + 255 + max_size_txt + 255
    
    ' Boutons
    On Error Resume Next
    g_nbbouton = UBound(SAIS_Saisie.boutons) + 1
    On Error GoTo 0
    lg_bouton = 0
    For I = 0 To g_nbbouton - 1
        If I > 0 Then Load cmd(I)
        cmd(I).visible = True
        If SAIS_Saisie.boutons(I).image <> "" Then
            cmd(I).Picture = CM_LoadPicture(SAIS_Saisie.boutons(I).image)
            cmd(I).Caption = ""
            cmd(I).ToolTipText = SAIS_Saisie.boutons(I).libelle
        Else
            cmd(I).Picture = LoadPicture("")
            cmd(I).Caption = SAIS_Saisie.boutons(I).libelle
        End If
        If SAIS_Saisie.boutons(I).largeur > 0 Then
            cmd(I).width = SAIS_Saisie.boutons(I).largeur
        End If
        lg_bouton = lg_bouton + cmd(I).width
    Next I
    lg_bouton1 = lg_bouton
    If lg_bouton > 0 Then
        lg_bouton = 255 + lg_bouton + 255 + (g_nbbouton - 1) * 510
    End If
        
    ' Labels et textes align�s
    For I = 0 To nb_champ - 1
        lblSaisie(I).width = max_size_lbl
        lblSaisie(I).left = marge
        If SAIS_Saisie.champs(I).type = SAIS_TYP_CHOIXLISTE Then
            ComboCmb(I).left = lblSaisie(I).left + max_size_lbl + 255
            If SAIS_Saisie.champs(I).liste_multiselect Then
                ComboCmd(I).left = ComboCmb(I).left + ComboCmb(I).width + 50
            End If
        Else
            txtSaisie(I).left = lblSaisie(I).left + max_size_lbl + 255
            If txtSaisie(I).MaxLength > max_nbcar_txt Then
                size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(max_nbcar_txt, "M"))
            Else
                size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(txtSaisie(I).MaxLength, "M"))
            End If
            txtSaisie(I).width = size_txt
        End If
    Next I

    ' Reglage largeur
    lg = lg_titre
    If lg < lg_bouton Then
        lg = lg_bouton
    End If
    If lg < lg_texte Then
        lg = lg_texte
    End If
    lg_bouton = lg + 512
    lg_tot = lg + 512
    frmSaisie.width = lg_tot
    frmFct.width = lg_tot
    Me.width = lg_tot
    
    ' Positionnement des boutons
    If g_nbbouton = 1 Then
        cmd(0).left = (frmFct.width - 510 - cmd(0).width) / 2
    Else
        intervalle = (frmFct.width - 510 - lg_bouton1) / (g_nbbouton - 1)
        left = 255
        For I = 0 To g_nbbouton - 1
            cmd(I).left = left
            left = left + cmd(I).width + intervalle
        Next I
    End If
        
    ' Calcul de la hauteur
    frmSaisie.Height = 255 + (2 * nb_champ * hauteur) + 300
    frmSaisie.Top = 0
    frmSaisie.ZOrder 0
    frmFct.Top = frmSaisie.Top + frmSaisie.Height - 150
    Me.Height = frmFct.Top + frmFct.Height + 300
    
    If SAIS_Saisie.prmfrm.x > 0 Then
        Me.left = SAIS_Saisie.prmfrm.x
    Else
        Me.left = (Screen.width - Me.width) / 2
    End If
    If SAIS_Saisie.prmfrm.y > 0 Then
        Me.Top = SAIS_Saisie.prmfrm.y
    Else
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    frmSaisie.left = 0
    
End Sub

Private Sub Remplir_Combo(ByVal v_ichp As Integer)
    Dim sql As String, rs As rdoResultset
    Dim ich As Integer
    Dim s As String
    Dim b_sel As Boolean
    
    If Not SAIS_Saisie.champs(ich).liste_multiselect Then ComboCmb(v_ichp).AddItem " "
    
    ComboCmb(v_ichp).Clear
    For ich = 0 To UBound(SAIS_Saisie.item_liste())
        If SAIS_Saisie.item_liste(ich).Liste_Num = v_ichp Then
            If SAIS_Saisie.champs(v_ichp).liste_multiselect Then
                ' on ne met que ceux qui sont s�lectionn�s
                If SAIS_Saisie.item_liste(ich).Item_bSel Then
                    ComboCmb(v_ichp).AddItem SAIS_Saisie.item_liste(ich).Item_LaStr
                    ComboCmb(v_ichp).ItemData(ComboCmb(v_ichp).ListCount - 1) = SAIS_Saisie.item_liste(ich).Item_Num
                End If
            Else
                ' voir si s�lectionn�
                ComboCmb(v_ichp).AddItem SAIS_Saisie.item_liste(ich).Item_LaStr
                ComboCmb(v_ichp).ItemData(ComboCmb(v_ichp).ListCount - 1) = SAIS_Saisie.item_liste(ich).Item_Num
                If SAIS_Saisie.item_liste(ich).Item_bSel Then
                    ComboCmb(v_ichp).ListIndex = ComboCmb(v_ichp).ListCount - 1
                End If
            End If
        End If
    Next ich
End Sub

Private Sub quitter(ByVal Index As Integer)

    Dim modif As Boolean
    Dim I As Integer
    
    Select Case Index
    Case 0
        If BOOL_YA_DES_CHAMPS Then
            modif = False
            For I = 0 To UBound(SAIS_Saisie.champs())
                If SAIS_Saisie.champs(I).type <> SAIS_TYP_CHOIXLISTE Then
                    If Not ctrl_saisie_ok(I, True) Then
                        txtSaisie(I).SetFocus
                        Exit Sub
                    End If
                End If
                If SAIS_Saisie.champs(I).type = SAIS_TYP_CHOIXLISTE Then
                    If SAIS_Saisie.champs(I).liste_multiselect Then
                        If SAIS_Saisie.champs(I).sval <> ComboCmd(I).tag Then modif = True
                        SAIS_Saisie.champs(I).sval = ComboCmd(I).tag
                    Else
                        If SAIS_Saisie.champs(I).sval <> ComboCmb(I).tag Then modif = True
                        SAIS_Saisie.champs(I).sval = ComboCmb(I).tag
                    End If
                Else
                    If SAIS_Saisie.champs(I).sval <> txtSaisie(I).Text Then modif = True
                    SAIS_Saisie.champs(I).sval = txtSaisie(I).Text
                End If
            Next I
            If modif Then
                SAIS_Saisie.retour = SAIS_RET_MODIF
            Else
                SAIS_Saisie.retour = SAIS_RET_NOMODIF
            End If
        Else
            SAIS_Saisie.retour = Index
        End If
    Case Else
        SAIS_Saisie.retour = Index
    End Select
    
    If SAIS_Saisie.prmfrm.reste_charg�e Then
        Me.Hide
    Else
        Unload Me
    End If
 
End Sub

Private Sub cmd_Click(Index As Integer)

    Call quitter(Index)
    
End Sub

Private Sub ComboCmb_Click(Index As Integer)
    Dim ich As Integer
    
    Me.ComboCmb(Index).tag = " "
    For ich = 0 To UBound(SAIS_Saisie.item_liste())
        If SAIS_Saisie.item_liste(ich).Liste_Num = Index Then
            ' voir si s�lectionn�
            If Me.ComboCmb(Index).ItemData(Me.ComboCmb(Index).ListIndex) = SAIS_Saisie.item_liste(ich).Item_Num Then
                Me.ComboCmb(Index).tag = SAIS_Saisie.item_liste(ich).Item_code_retour
                Exit For
            End If
        End If
    Next ich
End Sub

Private Sub ComboCmd_Click(Index As Integer)
    Dim sql As String, rs As rdoResultset
    Dim ich As Integer, n As Integer, I As Integer
    Dim s As String
    Dim b_sel As Boolean
    Dim op As String
    
    Call CL_Init
    
    Call CL_InitMultiSelect(True, False)
    
    n = 0
    For ich = 0 To UBound(SAIS_Saisie.item_liste())
        If SAIS_Saisie.item_liste(ich).Liste_Num = Index Then
            ' voir si s�lectionn�
            s = SAIS_Saisie.item_liste(ich).Item_LaStr
            Call CL_AddLigne(s, SAIS_Saisie.item_liste(ich).Item_Num, 0, SAIS_Saisie.item_liste(ich).Item_bSel)
            n = n + 1
        End If
    Next ich
    
    If n = 0 Then
        Call MsgBox("La liste est " & SAIS_Saisie.champs(Index).libelle & " Vide", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    
    If n < 20 Then
        Call CL_InitTaille(0, -(n + 1))
    Else
        Call CL_InitTaille(0, -20)
    End If
    
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    ' Quels sont ceux s�lectionn�s ?
    op = ""
    ComboCmd(Index).tag = ""
    ComboCmb(Index).Clear
    For I = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(I).selected Then
            sql = SAIS_Saisie.champs(Index).liste_nomtable & " where " & SAIS_Saisie.champs(Index).liste_chpnum & " = " & CL_liste.lignes(I).num
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                MsgBox "Erreur SQL " & SAIS_Saisie.champs(Index).liste_nomtable
                Exit Sub
            Else
                If rs.EOF Then
                    MsgBox "Erreur SQL vide : " & sql
                Else
                    ComboCmd(Index).tag = ComboCmd(Index).tag & op & rs(SAIS_Saisie.champs(Index).liste_chpretour)
                    op = "|"
                    ' les mettre dans le combo
                    For ich = 0 To UBound(SAIS_Saisie.item_liste())
                        If SAIS_Saisie.item_liste(ich).Liste_Num = Index Then
                            If SAIS_Saisie.champs(Index).liste_multiselect Then
                                ' on ne met que ceux qui sont s�lectionn�s
                                If SAIS_Saisie.item_liste(ich).Item_code_retour = rs(SAIS_Saisie.champs(Index).liste_chpretour) Then
                                    ComboCmb(Index).AddItem SAIS_Saisie.item_liste(ich).Item_LaStr
                                    ComboCmb(Index).ItemData(ComboCmb(Index).ListCount - 1) = SAIS_Saisie.item_liste(ich).Item_Num
                                    Exit For
                                End If
                            End If
                        End If
                    Next ich
                End If
            End If
            rs.Close
        End If
    Next I
    MsgBox ComboCmd(Index).tag
End Sub

Private Sub Form_Activate()
    
    Dim I As Integer
    
    If Not g_form_active Then
        g_form_active = True
        Call FRM_ResizeForm(Me, 0, 0)
        Call initialiser
    Else
        For I = 0 To UBound(SAIS_Saisie.champs())
            txtSaisie(I).Text = SAIS_Saisie.champs(I).sval
        Next I
    End If
    On Error GoTo Fin
    If Len(txtSaisie(g_ind_txt_enabled)) > 0 Then
        txtSaisie(g_ind_txt_enabled).SelLength = Len(txtSaisie(g_ind_txt_enabled).Text)
    End If
    txtSaisie(g_ind_txt_enabled).SetFocus
Fin:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nomchm As String, nomtopic As String
    Dim I As Integer
    
    If Shift = vbAltMask Then
        For I = 0 To g_nbbouton - 1
            If KeyCode = SAIS_Saisie.boutons(I).raccourci_alt Then
                KeyCode = 0
                Call quitter(I)
                Exit Sub
            End If
        Next I
        If KeyCode = vbKeyH Then
            KeyCode = 0
            If SAIS_Saisie.prmfrm.nomhelp <> "" Then
                If STR_GetNbchamp(SAIS_Saisie.prmfrm.nomhelp, ";") = 1 Then
                    nomchm = SAIS_Saisie.prmfrm.nomhelp
                    nomtopic = ""
                Else
                    nomchm = STR_GetChamp(SAIS_Saisie.prmfrm.nomhelp, ";", 0)
                    nomtopic = STR_GetChamp(SAIS_Saisie.prmfrm.nomhelp, ";", 1)
                End If
                Call HtmlHelp(0, nomchm, HH_DISPLAY_TOPIC, nomtopic)
            End If
        End If
    Else
        For I = 0 To g_nbbouton - 1
            If KeyCode = SAIS_Saisie.boutons(I).raccourci_touche Then
                KeyCode = 0
                Call quitter(I)
                Exit Sub
            End If
        Next I
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False
    Call FRM_ResizeForm(Me, 0, 0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(cmd.Count - 1)
    End If
    
End Sub

Private Sub txtSaisie_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If UBound(SAIS_Saisie.champs()) = 0 Or _
            (SAIS_Saisie.champs(Index).validationdirecte And _
            txtSaisie(Index).Text <> "") Then
            Call quitter(0)
        Else
            SendKeys "{TAB}"
        End If
        Exit Sub
    End If
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    Select Case SAIS_Saisie.champs(Index).type
    Case SAIS_TYP_JOUR_SEMAINE
        If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
            If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    Case SAIS_TYP_HEURE
        If KeyAscii = Asc(":") Then Exit Sub
        If KeyAscii = Asc("h") Then Exit Sub
        If KeyAscii = Asc("H") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_DATE
        If KeyAscii = Asc("/") Then Exit Sub
        If KeyAscii = Asc("+") Then Exit Sub
        If KeyAscii = Asc("-") Then Exit Sub
        If KeyAscii = Asc("j") Then Exit Sub
        If KeyAscii = Asc("J") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_ENTIER
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_ENTIER_NEG
        If KeyAscii = Asc("-") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_LETTRE
        If Not STR_EstAlpha(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_LETTRE_PONCT
        If Not STR_EstAlpha(Chr(KeyAscii)) Then
            If Not STR_EstPonctuation(Chr(KeyAscii)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    Case SAIS_TYP_PERIODE
        If KeyAscii = Asc("j") Then Exit Sub
        If KeyAscii = Asc("J") Then Exit Sub
        If KeyAscii = Asc("s") Then Exit Sub
        If KeyAscii = Asc("S") Then Exit Sub
        If KeyAscii = Asc("m") Then Exit Sub
        If KeyAscii = Asc("M") Then Exit Sub
        If KeyAscii = Asc("a") Then Exit Sub
        If KeyAscii = Asc("A") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_CAR_PARTICULIER
        If InStr(SAIS_Saisie.champs(Index).chaine_type, UCase(Chr(KeyAscii))) <= 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End Select
    
    If SAIS_Saisie.champs(Index).conversion = SAIS_CONV_MAJUSCULE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf SAIS_Saisie.champs(Index).conversion = SAIS_CONV_MINUSCULE Then
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtSaisie_LostFocus(Index As Integer)

    Dim I As Integer, n As Integer
    
    If g_nbbouton = 1 Then
        n = 0
    Else
        n = g_nbbouton - 1
    End If
    For I = 0 To n
        If cmd(I).tag = "" Then
            If Not ctrl_saisie_ok(Index, False) Then
                txtSaisie(Index).SetFocus
            End If
        End If
    Next I
    
End Sub


