VERSION 5.00
Begin VB.Form ChoixUtilisateur 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Utilisateur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   3105
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7905
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1350
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher par"
         ForeColor       =   &H00800080&
         Height          =   1035
         Left            =   1920
         TabIndex        =   8
         Top             =   1800
         Width           =   3525
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   2580
            Picture         =   "ChoixUtilisateur.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher dans la structure organisationnelle"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   525
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   5
            Left            =   1500
            Picture         =   "ChoixUtilisateur.frx":0585
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher parmi les fonctions"
            Top             =   420
            Width           =   525
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   360
            Picture         =   "ChoixUtilisateur.frx":09DF
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher parmi les groupes de personnes"
            Top             =   420
            Width           =   525
         End
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   3120
         MaxLength       =   40
         TabIndex        =   0
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ou"
         Height          =   255
         Index           =   2
         Left            =   2130
         TabIndex        =   13
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblDoc 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7485
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   870
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   2940
      Width           =   7905
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   480
         Picture         =   "ChoixUtilisateur.frx":0E7B
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Valider la personne"
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   6930
         Picture         =   "ChoixUtilisateur.frx":12D4
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Quitter sans choisir"
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixUtilisateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entrée : Titre du choix | un seul utilisateur autorisé |
'          Plusieurs utilisateurs autorisés | Dans la liste
' Sortie : U si utilisateur   p_choixlistem contient les utilisateurs selectionnés
'   ou     Fx si fonction
'   ou     Sx si service
'   ou     Rien si Abandon

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_CHOIX_GROUPE = 2
Private Const CMD_CHOIX_FONCTION = 5
Private Const CMD_CHOIX_SERVICE = 4
Private Const CMD_FERMER = 1

Private Const TXT_NOM = 0
Private Const TXT_CODE = 1

Private g_titre As String
Private g_titre_doc As String
Private g_plusieurs_util_autor As Boolean
Private g_choix_dans_liste As Boolean
Private g_actif As Boolean
Private g_fictif As Boolean
Private g_ssite As String
Private g_scr As String

Private g_mode_saisie As Boolean

Private g_txt_avant As String

Private g_form_active As Boolean

Public Function AppelFrm(ByVal v_titre As String, _
                          ByVal v_titre_doc As String, _
                          ByVal v_plusieurs_util_autor As Boolean, _
                          ByVal v_choix_dans_liste As Boolean, _
                          ByVal v_actif As Boolean, _
                          ByVal v_fictif As Boolean) As String

    g_titre = v_titre
    g_titre_doc = v_titre_doc
    g_plusieurs_util_autor = v_plusieurs_util_autor
    g_choix_dans_liste = v_choix_dans_liste
    g_ssite = ""
    g_actif = v_actif
    g_fictif = v_fictif
    
    ChoixUtilisateur.Show 1
    
    AppelFrm = g_scr
    
End Function

Private Sub build_sql_gfsu(ByVal v_sgrp As String, _
                           ByVal v_sfct As String, _
                           ByVal v_sspm As String, _
                           ByRef r_sql As String)

    Dim clause_labo As String, clause As String, sdest As String
    Dim s As String, s2 As String, slst As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    Dim chplabo As String
    
    If mode_Sites Then
        chplabo = "U_Site"
    Else
        chplabo = "U_Labo"
    End If
    clause_labo = ""
    If g_ssite <> "" Then
        clause = " and ("
        n = STR_GetNbchamp(g_ssite, ";")
        For i = 0 To n - 1
            clause_labo = clause_labo & clause & chplabo & " like '%" & STR_GetChamp(g_ssite, ";", i) & ";%'"
            clause = " or "
        Next i
        If clause_labo <> "" Then
            clause_labo = clause_labo + ")"
        End If
    End If
    
    sdest = ""
    If v_sgrp <> "" Then
        n = STR_GetNbchamp(v_sgrp, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(v_sgrp, "|", i)
            If Odbc_Select("select GU_Lst from GroupeUtil where GU_Num=" & Mid$(s, 2), _
                             rs) = P_OK Then
                slst = rs("GU_Lst").Value & ""
                rs.Close
                If slst = "" Then slst = "U1"
                sdest = sdest + slst
            End If
        Next i
    ElseIf v_sfct <> "" Then
        sdest = sdest + v_sfct
    ElseIf v_sspm <> "" Then
        sdest = sdest + v_sspm
    End If
    
    r_sql = "select U_Num, U_Nom, U_Prenom, U_SPM, U_FctTrav" _
            & " from Utilisateur" _
            & " where U_Num=U_Num" _
            & clause_labo
    If g_actif Then
        r_sql = r_sql & " and U_Actif=True"
    End If
    If Not g_fictif Then
        r_sql = r_sql & " and U_Fictif=false"
    End If
    If sdest <> "" Then
        n = STR_GetNbchamp(sdest, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(sdest, "|", i)
            If i = 0 Then
                r_sql = r_sql & " and ("
            Else
                r_sql = r_sql & " or"
            End If
            Select Case left$(s, 1)
            Case "F"
                r_sql = r_sql & " U_FctTrav like '%" & s & ";%'"
            Case "S"
                r_sql = r_sql & " U_SPM like '%" & s & "%'"
            Case "U"
                r_sql = r_sql & " U_Num=" & Mid$(s, 2)
            End Select
        Next i
        r_sql = r_sql & ")"
    End If
    r_sql = r_sql & " order by U_Nom, U_Prenom"

End Sub

Private Sub choisir_dans_la_liste()

    Dim nomutil As String, libsp As String
    Dim i As Integer
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    p_siz_tblu_sel = -1
    
    Call CL_Init
    Call CL_InitTitreHelp(g_titre + " " + g_titre_doc, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    If g_plusieurs_util_autor Then
        Call CL_AddBouton("&Tous", "", 0, 0, 0)
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    For i = 0 To p_siz_tblu
        If recup_sp(p_tblu(i), nomutil, libsp) = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
        Call CL_AddLigne(nomutil & vbTab & libsp, _
                         p_tblu(i), _
                         "", _
                         False)
    Next i
    Call CL_InitTaille(0, -20)
    If g_plusieurs_util_autor Then
        Call CL_InitMultiSelect(True, True)
        ChoixListe.Show 1
        If CL_liste.retour = 2 Then
            Call quitter
            Exit Sub
        End If
        If CL_liste.retour = 1 Then
            p_siz_tblu_sel = p_siz_tblu
            ReDim p_tblu_sel(p_siz_tblu_sel) As Long
            For i = 0 To p_siz_tblu
                p_tblu_sel(i) = p_tblu(i)
            Next i
        Else
            For i = 0 To p_siz_tblu
                If CL_liste.lignes(i).selected = True Then
                    p_siz_tblu_sel = p_siz_tblu_sel + 1
                    ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
                    p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
                End If
            Next i
        End If
    Else
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            Call quitter
            Exit Sub
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
    End If
    
    g_scr = "U"
    Unload Me

End Sub

Private Sub choisir_fonction()

    Dim sql As String, sret As String, sfct As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
    Call CL_Init
    n = 0
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        n = n + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Sub
    End If
    
    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    sfct = ""
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
            sfct = sfct & "F" & CL_liste.lignes(i).num & "|"
        End If
    Next i
    Call choisir_utilisateur_gfsu("", sfct, "")
    
End Sub

Private Sub choisir_groupe()

    Dim sql As String, sret As String, sgrp As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
    Call CL_Init
    n = 0
    sql = "select * from GroupeUtil" _
        & " order by GU_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("GU_Nom").Value, rs("GU_Num").Value, "", False)
        n = n + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Sub
    End If
    
    Call CL_InitTitreHelp("Groupes de personnes", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    sgrp = ""
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
            sgrp = sgrp & "G" & CL_liste.lignes(i).num & "|"
        End If
    Next i
    Call choisir_utilisateur_gfsu(sgrp, "", "")
    
End Sub

Private Sub choisir_service()

    Dim sret As String, ssite As String, s_srv As String, sprm As String
    Dim encore As Boolean
    Dim i As Integer, n As Integer
    Dim numlabo As Long, numutil As Long
    Dim Frm As Form
    
    numlabo = p_numlabo
    
    If g_ssite <> "" Then
        ssite = STR_Supprimer(g_ssite, "L")
    End If
    
    encore = True
    Do
        Call CL_Init
        Set Frm = KS_PrmService
        sret = KS_PrmService.AppelFrm("Choix d'un service", "S", g_plusieurs_util_autor, ssite, "SP", True)
        Set Frm = Nothing
        If sret = "" Then
            Exit Sub
        End If
        If g_plusieurs_util_autor And left$(sret, 1) = "N" Then
            encore = False
        ElseIf Not g_plusieurs_util_autor And left$(sret, 1) = "S" Then
            encore = False
        Else
        End If
    Loop Until encore = False
    
    If g_plusieurs_util_autor Then
        s_srv = ""
        n = CLng(Mid$(sret, 2))
        If n = 0 Then
            Exit Sub
        End If
        For i = 0 To n - 1
            s_srv = s_srv + CL_liste.lignes(i).texte + "|"
        Next i
    Else
        s_srv = sret
    End If
    Call choisir_utilisateur_gfsu("", "", s_srv)
    
End Sub

Private Sub choisir_utilisateur_gfsu(ByVal v_sgrp As String, _
                                     ByVal v_sfct As String, _
                                     ByVal v_sp As String)

    Dim sql As String, libsp As String, libfct As String, sret As String
    Dim s As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
    Call build_sql_gfsu(v_sgrp, v_sfct, v_sp, sql)
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    Call CL_Init
    Call CL_InitTitreHelp(g_titre, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = 0
    While Not rs.EOF
        If P_UtilDansTBL(rs("U_Num").Value) Then
            GoTo lab_suiv1
        End If
        s = rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab
        If P_RecupSPLib(rs("U_SPM").Value, libsp) = P_ERREUR Then
            Exit Sub
        End If
        s = s & libsp & vbTab
        Call CL_AddLigne(s, rs("U_Num").Value, "", False)
        n = n + 1
lab_suiv1:
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        Me.MousePointer = 0
        Exit Sub
    End If
    
    Call CL_InitTaille(0, -20)
    If g_plusieurs_util_autor Then
        Call CL_InitMultiSelect(True, True)
        Call CL_InitGererTousRien(True)
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        For i = 0 To n - 1
            If CL_liste.lignes(i).selected = True Then
                p_siz_tblu_sel = p_siz_tblu_sel + 1
                ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
                p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
            End If
        Next i
    Else
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
    End If
    
    Me.MousePointer = 0
    
    g_scr = "U"
    Unload Me
    
End Sub

Private Sub initialiser()

    p_siz_tblu_sel = -1
    
    g_mode_saisie = False
    
    If g_choix_dans_liste And p_siz_tblu <> -1 Then
        Call choisir_dans_la_liste
        Exit Sub
    End If
    
    Frm.Caption = g_titre
    lblDoc.Caption = g_titre_doc
    
    g_mode_saisie = True
    txt(TXT_NOM).SetFocus

End Sub

Private Sub quitter()

    g_scr = ""
    Unload Me
    
End Sub

Private Function recup_sp(ByVal v_numutil As Long, _
                          ByRef r_nomutil As String, _
                          ByRef r_libsp As String) As Integer

    Dim prenom As String
    Dim s_sp As Variant
    Dim rs As rdoResultset
    
    If Odbc_RecupVal("select U_Nom, U_Prenom, U_SPM from Utilisateur where U_Num=" & v_numutil, _
                     r_nomutil, _
                     prenom, _
                     s_sp) = P_ERREUR Then
        recup_sp = P_ERREUR
        Exit Function
    End If
    
    r_nomutil = r_nomutil + " " + prenom
    
    s_sp = STR_GetChamp(s_sp, "|", 0)
    If P_RecupSPLib(s_sp, r_libsp) = P_ERREUR Then
        recup_sp = P_ERREUR
        Exit Function
    End If
    
    recup_sp = P_OK
    
End Function

Public Function P_RecupSPLib(ByVal v_sp As String, _
                             ByRef r_lib As String) As Integer
    
    Dim sql As String, stype As String, s As String, lib As String, sp As String
    Dim n As Integer
    Dim num As Long

    r_lib = ""
    
    If v_sp <> "" Then
        sp = STR_GetChamp(v_sp, "|", 0)
    Else
        sp = v_sp
    End If
    n = STR_GetNbchamp(sp, ";")
    s = STR_GetChamp(sp, ";", n - 1)
    
    num = Mid$(s, 2)
    If left$(s, 1) = "S" Then
        If P_RecupSrvNom(num, lib) = P_ERREUR Then
            P_RecupSPLib = P_ERREUR
            Exit Function
        End If
    Else
        If P_RecupPosteNom(num, lib) = P_ERREUR Then
            P_RecupSPLib = P_ERREUR
            Exit Function
        End If
    End If
    r_lib = lib
    
    P_RecupSPLib = P_OK

End Function

Private Function verif_util() As Integer

    Dim sql As String, nomutil As String, libsp As String
    Dim codutil As String, mess As String
    Dim est_ok As Boolean
    Dim n As Integer, i As Integer, j As Integer, nbchp_u As Integer, nbchp_p As Integer
    Dim nbtot As Integer
    Dim numutil As Long, numlabo As Long, lnb As Long
    Dim rs As rdoResultset
    Dim chplabo As String
    
    If txt(TXT_NOM).Text = "" And txt(TXT_CODE).Text = "" Then
        verif_util = P_NON
        Exit Function
    End If
    
    If mode_Sites Then
        chplabo = "U_Site"
    Else
        chplabo = "U_Labo"
    End If
    If txt(TXT_NOM).Text <> "" Then
        nomutil = UCase(txt(TXT_NOM).Text)
        If Right$(nomutil, 1) <> "*" Then
            nomutil = nomutil + "*"
        End If
        sql = "select U_Num, U_Nom, U_Prenom, " & chplabo & " from Utilisateur, UtilAppli" _
            & " where U_Num>1" _
            & " and U_SPM<>''" _
            & " and UAPP_UNum=U_Num" _
            & " and UAPP_APPNum=" & p_appli_kalidoc _
            & " and " & Odbc_upper() & "(U_Nom) like " & Odbc_String(nomutil)
        If g_actif Then
            sql = sql & " and U_Actif=True"
        End If
        If Not g_fictif Then
            sql = sql & " and U_Fictif=false"
        End If
        sql = sql & " order by U_Nom, U_Prenom"
    Else
        codutil = UCase(txt(TXT_CODE).Text)
        sql = "select U_Num, U_Nom, U_Prenom, " & chplabo & " from Utilisateur, UtilAppli" _
            & " where U_Num>1" _
            & " and U_SPM<>''" _
            & " and UAPP_UNum=U_Num" _
            & " and UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code=" & Odbc_String(codutil)
        If g_actif Then
            sql = sql & " and U_Actif=True"
        End If
        If Not g_fictif Then
            sql = sql & " and U_Fictif=false"
        End If
    End If
'MsgBox sql
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_util = P_ERREUR
        Exit Function
    End If
    If Not rs.EOF Then
        rs.MoveNext
        If rs.EOF Then
            rs.MovePrevious
            If P_UtilDansTBL(rs("U_Num").Value) Then
                Call MsgBox("'" & rs("U_Prenom").Value & " " & rs("U_Nom").Value & " est déjà dans la liste." & vbCrLf & "Vous ne pouvez pas le rajouter.", vbInformation + vbOKOnly, "")
                rs.Close
                verif_util = P_NON
                Exit Function
            End If
            If txt(TXT_NOM).Text = "" Then
                p_siz_tblu_sel = 0
                ReDim p_tblu_sel(0) As Long
                p_tblu_sel(0) = rs("U_Num").Value
                rs.Close
                GoTo lab_fin
            End If
        Else
            rs.MovePrevious
        End If
        GoTo lab_affiche
    End If
        
lab_affiche:
    Call CL_Init
    Call CL_InitTitreHelp("Personnes ayant le critère recherché", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = 0
    nbtot = 0
    While Not rs.EOF
        nbtot = nbtot + 1
        ' Utilisateur déjà dans la liste ?
        If P_UtilDansTBL(rs("U_Num").Value) Then GoTo lab_suivant
        ' Utilisateur fait partie des labos indiqués ?
        If g_ssite <> "" Then
            est_ok = False
            nbchp_u = STR_GetNbchamp(rs(chplabo).Value, ";")
            nbchp_p = STR_GetNbchamp(g_ssite, ";")
            For i = 0 To nbchp_u - 1
                numlabo = Mid$(STR_GetChamp(rs(chplabo).Value, ";", i), 2)
                For j = 0 To nbchp_p - 1
                    If numlabo = Mid$(STR_GetChamp(g_ssite, ";", j), 2) Then
                        est_ok = True
                        Exit For
                    End If
                Next j
            Next i
            If Not est_ok Then
                GoTo lab_suivant
            End If
        End If
        If recup_sp(rs("U_Num").Value, nomutil, libsp) = P_ERREUR Then
            verif_util = P_ERREUR
            Exit Function
        End If
        Call CL_AddLigne(rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab & libsp, _
                         rs("U_Num").Value, _
                         "", _
                         False)
        n = n + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
        
    If n = 0 Then
        mess = "Aucune personne "
        If g_actif Then
            mess = mess & "ACTIVE "
        End If
        'If p_NbLabo > 1 Then
        '    mess = mess & "faisant partie des sites indiqués "
        'End If
        mess = mess & "n'a été trouvée avec les critères désirés."
        Call MsgBox(mess, vbInformation + vbOKOnly, "")
        verif_util = P_NON
        Exit Function
    End If
        
    Call CL_InitTaille(0, -20)

'    If g_plusieurs_util_autor Then
'        Call CL_InitMultiSelect(True)
'        ChoixListe.Show 1
'        If CL_liste.retour = 1 Then
'            verif_util = P_NON
'            Exit Function
'        End If
'        p_siz_tblu_sel = -1
'        For i = 0 To n - 1
'            If CL_liste.lignes(i).selected Then
'                p_siz_tblu_sel = p_siz_tblu_sel + 1
'                ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
'                p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
'            End If
'        Next i
'    Else
        
        ' Ne pas supprimer : sinon txt_LostFocus reprend la main et ChoixListe
        ' est lancée une 2e fois ...
        g_mode_saisie = False
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            g_mode_saisie = True
            verif_util = P_NON
            Exit Function
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
'    End If
    
lab_fin:
    g_scr = "U"
    Unload Me
    
    verif_util = P_OK

End Function

Public Function P_UtilDansTBL(ByVal v_numutil As Long) As Boolean

    Dim i As Integer
    
    For i = 0 To p_siz_tblu
        If p_tblu(i) = v_numutil Then
            P_UtilDansTBL = True
            Exit Function
        End If
    Next i
    
    P_UtilDansTBL = False
    
End Function

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case CMD_CHOIX_GROUPE
        Call choisir_groupe
    Case CMD_CHOIX_FONCTION
        Call choisir_fonction
    Case CMD_CHOIX_SERVICE
        Call choisir_service
    Case CMD_FERMER
        Call quitter
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_FERMER Then g_mode_saisie = False
    
End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub
    
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If
        
End Sub

Private Sub Form_Load()

    g_mode_saisie = False
    g_form_active = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_avant = txt(TXT_NOM).Text
    
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        If Index = TXT_NOM Then
            Call choisir_utilisateur_gfsu("", "", "")
        End If
    End If
    
End Sub

Private Sub txt_lostfocus(Index As Integer)

    If g_mode_saisie Then
        If g_txt_avant <> txt(Index).Text Then
            If Index = TXT_NOM Or Index = TXT_CODE Then
                If verif_util() <> P_OUI Then
                    txt(TXT_NOM).Text = ""
                    txt(TXT_CODE).Text = ""
                    txt(TXT_NOM).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub
