VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form KS_PrmService 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Services"
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
      Height          =   7515
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8265
      Begin VB.TextBox TxtRecherche 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   5055
      End
      Begin VB.ComboBox CmbNiveau 
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin ComctlLib.TreeView tv 
         Height          =   6645
         Left            =   195
         TabIndex        =   5
         Top             =   555
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11721
         _Version        =   327682
         Indentation     =   2
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "img"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ImageList img 
         Left            =   6120
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   7
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":0852
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":1124
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":19F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":22C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":2B9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KS_PrmService.frx":33EC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDepl 
         BackColor       =   &H000080FF&
         Caption         =   "Cliquez sur le nouveau service ou cliquez ici pour Annuler"
         Height          =   495
         Left            =   2370
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   3585
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   8265
      Begin ComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
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
         Left            =   7560
         Picture         =   "KS_PrmService.frx":3A1E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Quitter"
         Top             =   200
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
         Index           =   0
         Left            =   240
         Picture         =   "KS_PrmService.frx":3FD7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sélectionner"
         Top             =   200
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
         Index           =   2
         Left            =   960
         Picture         =   "KS_PrmService.frx":4430
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimer"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.Label LbldetailSRV 
         Height          =   735
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Menu mnuFct 
      Caption         =   "mnuFct"
      Visible         =   0   'False
      Begin VB.Menu mnuCreerS 
         Caption         =   "&Créer un service"
      End
      Begin VB.Menu mnuSepCrS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModS 
         Caption         =   "&Modifier le service"
      End
      Begin VB.Menu mnuSuppS 
         Caption         =   "&Supprimer le service"
      End
      Begin VB.Menu mnuSepMSS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreerP 
         Caption         =   "C&réer un poste"
      End
      Begin VB.Menu mnuSepCrP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPosteResp 
         Caption         =   "Poste responsable"
      End
      Begin VB.Menu mnuLibPoste 
         Caption         =   "Libellé du poste"
      End
      Begin VB.Menu mnuSuppP 
         Caption         =   "&Supprimer le poste"
      End
      Begin VB.Menu mnuSepSuppP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepl 
         Caption         =   "&Déplacer dans un autre service"
      End
      Begin VB.Menu mnuSepDepl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoirPers 
         Caption         =   "&Voir les personnes"
      End
      Begin VB.Menu mnuSepVoirPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrPers 
         Caption         =   "&Créer une personne"
      End
      Begin VB.Menu mnuSepCrPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModPers 
         Caption         =   "&Modifier les caractéristiques de la personne"
      End
      Begin VB.Menu mnuSepModPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "KS_PrmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OK = 0
Private Const CMD_IMPRIMER = 2
Private Const CMD_QUITTER = 1

Private Const IMG_SRV = 1
Private Const IMG_POSTE = 2
Private Const IMG_SRV_SEL = 3
Private Const IMG_POSTE_SEL = 4
Private Const IMG_SRV_SEL_NOMOD = 5
Private Const IMG_POSTE_SEL_NOMOD = 6
Private Const IMG_UTIL = 7

Private Const MODE_PARAM = 0
Private Const MODE_SELECT = 1
Private Const MODE_PARAM_PERS = 2

Private g_mode_acces As Integer
Private g_smode As String

Private g_plusieurs As Boolean
Private g_ssite As String
Private g_stype As String
Private g_prmpers As Boolean
Private g_numserv As Long
Private g_numsite As Long
Private g_sret As String

Private g_crfct_autor As Boolean
Private g_crutil_autor As Boolean
Private g_modutil_autor As Boolean
Private g_crspm_autor As Boolean
Private g_modspm_autor As Boolean
Private g_supspm_autor As Boolean

Private g_tbl_site() As Long

Private g_lignes() As CL_SLIGNE

Private g_node_crt As Long
Private g_pos_depl As Long

Private g_node As Integer
Private g_expand As Boolean
Private g_button As Integer
Private g_mode_saisie As Boolean
Private g_form_active As Boolean

Public g_ya_niveau As Boolean
Public g_ouvrir As String

'V_smode :      "C" --> Quand on vient de prmClasseur

Public Function AppelFrm(ByVal v_stitre As String, _
                         ByVal v_smode As String, _
                         ByVal v_bplusieurs As Boolean, _
                         ByVal v_ssite As String, _
                         ByVal v_stype As String, _
                         ByVal v_prmpers As Boolean) As String

    If v_smode = "M" Then
        g_smode = v_smode
        g_mode_acces = MODE_PARAM
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "S" Or v_smode = "C" Then
        g_smode = v_smode
        g_mode_acces = MODE_SELECT
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "P" Then
        g_smode = v_smode
        g_mode_acces = MODE_PARAM_PERS
        If left$(v_stype, 1) = "S" Then
            g_numserv = Mid$(v_stype, 2)
            g_numsite = 0
        Else
            g_numserv = 0
            g_numsite = Mid$(v_stype, 2)
        End If
        g_stype = "SP"
    End If
    g_plusieurs = v_bplusieurs
    g_ssite = v_ssite
    g_prmpers = v_prmpers
    
    frm.Caption = v_stitre
    
    Me.Show 1
    
    AppelFrm = g_sret
    
End Function

Private Sub activer_depl()

    g_node_crt = tv.SelectedItem.Index
    g_pos_depl = g_node_crt
    lblDepl.Caption = "Cliquez sur le nouveau dossier de rattachement ou cliquez ici pour ANNULER l'opération"
    lblDepl.BackColor = P_ORANGE
    lblDepl.Visible = True
    
End Sub
Private Function afficher_liste() As Integer

    Dim sql As String, s As String, sfct As String, lib As String
    Dim stag As String, libNiveau As String
    Dim fmodif As Boolean, afficher As Boolean, trouve As Boolean
    Dim img As Integer, i As Integer, nsel As Integer, n As Integer, j As Integer
    Dim numsrv As Long, lnb As Long, num As Long
    Dim rs As rdoResultset
    Dim strSites As String
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    g_mode_saisie = False
    
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "select EJ_Num, EJ_Nom from EntJuridique"
    If Odbc_RecupVal(sql, num, lib) = P_ERREUR Then
        afficher_liste = P_ERREUR
        Exit Function
    End If
    Set ndp = tv.Nodes.Add(, , "L" & num, lib, IMG_SRV, IMG_SRV)
    ndp.Expanded = True
    
    ' Les services
    sql = "select SRV_Num, SRV_NumPere, SRV_Nom, SRV_NivsNum, SRV_Actif from Service " _
        & " where true" _
        & " and SRV_Numpere=0"

    If g_ssite <> "" Then
        sql = sql & " and SRV_Site like '%" & g_ssite & "%'"
    End If
    sql = sql & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste = P_ERREUR
        Exit Function
    End If
    
    If Not rs.EOF Then
        rs.MoveLast
        PgBar.Visible = True
        PgBar.max = rs.RowCount
        PgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        PgBar.Value = PgBar.Value + 1

        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        If Not rs("SRV_Actif").Value Then
            lib = lib & " (inactif)"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            afficher_liste = P_ERREUR
            Exit Function
        End If
        If lnb = 0 And g_smode <> "C" Then
            sql = "select count(*) from Poste " _
                & " where PO_Actif=true" _
                & " and PO_SRVNum=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                afficher_liste = P_ERREUR
                Exit Function
            End If
        End If
        If lnb = 0 Then
            nd.tag = True & "|" & True
        Else
            nd.tag = True & "|" & False
            Set nd = tv.Nodes.Add(nd, _
                               tvwChild, _
                               , _
                               "A charger")
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    PgBar.Visible = False
    
    ' Met en évidence les noeuds 'retenus'
    If g_mode_acces = MODE_SELECT And g_plusieurs Then
        For i = 0 To nsel
            n = STR_GetNbchamp(g_lignes(i).texte, ";")
            s = STR_GetChamp(g_lignes(i).texte, ";", n - 1)
            If s = "0" Then
                s = "L" & p_num_ent_juridique
            End If
            If left$(s, 1) = "S" Or left$(s, 1) = "L" Then
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_SRV_SEL
                Else
                    img = IMG_SRV_SEL_NOMOD
                End If
            Else
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_POSTE_SEL
                Else
                    img = IMG_POSTE_SEL_NOMOD
                End If
            End If
            If TV_NodeExiste(tv, s, nd) = P_NON Then
                Call charger_arbor(s)
            End If
            If TV_NodeExiste(tv, s, nd) = P_OUI Then
                Set nd = tv.Nodes(s)
                nd.image = img
                If g_lignes(i).tag <> "" Then
                    nd.Text = nd.Text & " (" & g_lignes(i).tag & ")"
                End If
                stag = nd.tag
                Call STR_PutChamp(stag, "|", 0, fmodif)
                nd.tag = stag
                nd.SelectedImage = img
                If left$(nd.key, 1) <> "L" Then
                    Set ndp = nd.Parent
                    While left$(ndp.key, 1) <> "L"
                        ndp.Expanded = True
                        Set ndp = ndp.Parent
                    Wend
                End If
            End If
        Next i
    End If
    
    Call ouvrir_serv_poste
    
    tv.SetFocus
    g_mode_saisie = True
    
    Set nd = Nothing
    Set ndp = Nothing
    Set ndp_sav = Nothing
    
    afficher_liste = P_OK
    
End Function

Private Sub ouvrir_serv_poste()

    Dim sql As String
    Dim encore As Boolean
    Dim numposte As Long
    Dim nd As Node
    Dim rs As rdoResultset
    
    If g_ouvrir <> "" Then
        numposte = 0
        If left$(g_ouvrir, 1) = "P" Then
            numposte = Mid$(g_ouvrir, 2)
        ElseIf left$(g_ouvrir, 1) = "U" Then
            sql = "select U_Po_Princ from Utilisateur" _
                & " where U_Num=" & Mid$(g_ouvrir, 2)
            If Odbc_SelectV(sql, rs) = P_OK Then
                If Not rs.EOF Then
                    numposte = rs("U_Po_Princ").Value
                End If
                rs.Close
            End If
        End If
        If numposte > 0 Then
            If TV_NodeExiste(tv, "P" & numposte, nd) = P_NON Then
                Call charger_arbor("P" & numposte)
            End If
            Set nd = tv.Nodes("P" & numposte)
            encore = True
            While encore
                If nd.Index = nd.Root.Index Then
                    encore = False
                Else
                    Set nd = nd.Parent
                    nd.Expanded = True
                End If
            Wend
            tv.SetFocus
            Set tv.SelectedItem = tv.Nodes("P" & numposte)
            SendKeys "{DOWN}"
            SendKeys "{UP}"
            DoEvents
            Set tv.SelectedItem = tv.Nodes("P" & numposte)
            If left$(g_ouvrir, 1) = "U" Then
                Call ajouter_pers_tv
            End If
        End If
    Else
        Set tv.SelectedItem = tv.Nodes(1).Root
    End If
    
End Sub

Private Function charger_arbor(ByVal v_ssrv As String) As Integer

    Dim sql As String, s_srv As String, s As String
    Dim i As Integer, n As Integer
    Dim numsrv As Long, numposte As Long
    Dim nd As Node, ndp As Node
    
    If left$(v_ssrv, 1) = "L" Then
        Exit Function
    ElseIf left$(v_ssrv, 1) = "P" Then
        numposte = Mid$(v_ssrv, 2)
        sql = "select PO_SRVNum from Poste where PO_Num=" & numposte
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
    Else
        numsrv = Mid$(v_ssrv, 2)
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
    End If
    s_srv = numsrv & ";"
    While numsrv > 0
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
        If numsrv > 0 Then
            s_srv = numsrv & ";" & s_srv
        End If
    Wend
    
    n = STR_GetNbchamp(s_srv, ";")
    For i = 0 To n - 1
        numsrv = STR_GetChamp(s_srv, ";", i)
        If TV_NodeExiste(tv, "S" & numsrv, nd) = P_OUI Then
            If STR_GetChamp(nd.tag, "|", 1) = False Then
                tv.Nodes.Remove (nd.Child.Index)
                If charger_service(numsrv) = P_ERREUR Then
                    charger_arbor = P_ERREUR
                    Exit Function
                End If
            End If
        End If
    Next i
    
    charger_arbor = P_OK

End Function

Private Function charger_service(ByVal v_numsrv As Long) As Integer

    Dim sql As String, sfct As String, lib As String, stag As String
    Dim libNiveau As String
    Dim img As Integer, i As Integer
    Dim lnb As Long
    Dim strSites As String
    Dim rs As rdoResultset, rsU As rdoResultset
    Dim nd As Node, ndp As Node, ndu As Node
    Dim strRemplace As String
    
    If v_numsrv = 0 Then
        Set ndp = tv.Nodes(1)
    Else
        Set ndp = tv.Nodes("S" & v_numsrv)
        ndp.tag = True & "|" & False
    End If

    ' Les postes
    If g_smode <> "C" Then
        sql = "select PO_Num, PO_Libelle, FT_Libelle, FT_NivRemplace" _
            & " from Poste, FctTrav" _
            & " where FT_Num=PO_FTNum" _
            & " and PO_Actif=true" _
            & " and PO_SRVNum=" & v_numsrv _
            & " order by PO_Libelle"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        
        If Not rs.EOF Then
            rs.MoveLast
            PgBar.Visible = True
            PgBar.max = rs.RowCount
            PgBar.Value = 0
            rs.MoveFirst
        End If
        
        While Not rs.EOF
            PgBar.Value = PgBar.Value + 1
            sfct = rs("FT_Libelle").Value
            If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                sfct = sfct & " *"
            End If
            ' Vérification du niveau de remplacement
            If FctNivRemplace(rs("FT_NivRemplace"), v_numsrv, strRemplace) < 0 Then
                MsgBox strRemplace
                sfct = sfct & " (" & strRemplace & ")"
            End If
            
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "P" & rs("PO_Num").Value, _
                                   sfct, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
            nd.tag = True
            If g_mode_acces = MODE_PARAM_PERS Then
                ' Les personnes associées
                sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
                    & " where U_SPM like '%P" & rs("PO_Num").Value & ";%'" _
                    & " and U_Actif=true"
                If Odbc_SelectV(sql, rsU) = P_ERREUR Then
                    charger_service = P_ERREUR
                    Exit Function
                End If
                While Not rsU.EOF
                    Set ndu = tv.Nodes.Add(nd, _
                                           tvwChild, _
                                           "", _
                                           rsU("U_Nom").Value & " " & rsU("U_Prenom").Value, _
                                           IMG_UTIL, _
                                           IMG_UTIL)
                    ndu.tag = "U" & rsU("U_Num").Value
                    rsU.MoveNext
                Wend
                rsU.Close
            End If
            rs.MoveNext
        Wend
        rs.Close
        PgBar.Visible = False
    End If
    
    If TxtRecherche.Text <> "" Then
        GoTo lab_fin
    End If
    
    ' Les services
    sql = "select SRV_Num, SRV_NumPere, SRV_Nom, SRV_NivsNum, SRV_Actif from Service " _
        & " where true" _
        & " and SRV_Numpere=" & v_numsrv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_service = P_ERREUR
        Exit Function
    End If
    
    If Not rs.EOF Then
        rs.MoveLast
        PgBar.Visible = True
        PgBar.max = rs.RowCount
        PgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        PgBar.Value = PgBar.Value + 1

        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        If Not rs("SRV_Actif").Value Then
            lib = lib & " (inactif)"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        If lnb = 0 And g_smode <> "C" Then
            sql = "select count(*) from Poste " _
                & " where PO_Actif=true" _
                & " and PO_SRVNum=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                charger_service = P_ERREUR
                Exit Function
            End If
        End If
        If lnb = 0 Then
            nd.tag = True & "|" & True
        Else
            nd.tag = True & "|" & False
            Set nd = tv.Nodes.Add(nd, _
                               tvwChild, _
                               , _
                               "A charger")
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    PgBar.Visible = False
    
lab_fin:
    
    stag = ndp.tag
    Call STR_PutChamp(stag, "|", 1, True)
    ndp.tag = stag
    
    charger_service = P_OK
    
End Function

Private Function FctNivRemplace(ByVal v_FT_NivRemplace, ByVal v_numsrv, ByRef r_strRemplace As String) As Integer
    ' Voir si pour ce service il a un père du niveau de remplacement indiqué
    Dim sql As String, rs As rdoResultset
    Dim encore As Boolean
    Dim ilya As Boolean
    Dim srvnum As Long
    Dim srvnom_prem As String
    Dim strNiveau As String
    
    If v_FT_NivRemplace = 0 Then
        FctNivRemplace = 0
    Else
        srvnum = v_numsrv
        encore = True
        While encore
            sql = "select SRV_Num, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service where SRV_Num=" & srvnum
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                r_strRemplace = "Erreur " & sql
                FctNivRemplace = P_ERREUR
                Exit Function
            ElseIf rs.EOF Then
                Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", strNiveau)
                r_strRemplace = "Il n'y a pas de niveau de remplacement (" & strNiveau & ")" '    & " pour " & srvnom_prem
                FctNivRemplace = -1
                rs.Close
                Exit Function
            ElseIf rs("SRV_NivsNum") = v_FT_NivRemplace Then
                Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", strNiveau)
                r_strRemplace = "Niveau de remplacement : " & strNiveau & " => " & rs("SRV_Nom")
                FctNivRemplace = 1
                rs.Close
                Exit Function
            Else    ' voir son pere
                srvnum = rs("SRV_NumPere")
                If srvnom_prem = "" Then srvnom_prem = rs("SRV_Nom")
            End If
        Wend
        rs.Close
    End If
End Function


Private Function recup_lib_niveau(ByVal v_nivsnum As Long) As String
    
    Dim sql As String, nNiv As Long
    Dim libNiveau As String
    
    If Not g_ya_niveau Then
        recup_lib_niveau = ""
        Exit Function
    End If
    
    sql = "select count(*) from niveau_structure where Nivs_Num=" & v_nivsnum
    If Odbc_Count(sql, nNiv) = P_ERREUR Then
        recup_lib_niveau = ""
        Exit Function
    Else
        If v_nivsnum = 0 Then
            recup_lib_niveau = ""
        Else
            sql = "select Nivs_Nom from niveau_structure where Nivs_Num=" & v_nivsnum
            If Odbc_RecupVal(sql, libNiveau) = P_ERREUR Then
                recup_lib_niveau = ""
                Exit Function
            Else
                recup_lib_niveau = libNiveau
            End If
        End If
    End If
    
End Function

Private Function afficher_liste_OLD() As Integer

    Dim sql As String, s As String, sfct As String
    Dim fmodif As Boolean, afficher As Boolean, trouve As Boolean
    Dim img As Integer, i As Integer, nsel As Integer, n As Integer, j As Integer
    Dim mode As Integer
    Dim numsrv As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    g_mode_saisie = False
    
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "select L_Num, L_Code from Laboratoire order by L_Code"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        afficher_liste_OLD = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Set nd = tv.Nodes.Add(, , "L" & rs("L_Num").Value, rs("L_Code").Value)
        If rs("L_Num").Value = p_numlabo Then
            nd.selected = True
            nd.Expanded = True
        End If
        nd.Expanded = True
'        nd.Sorted = True
        rs.MoveNext
    Wend
    rs.Close
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
        & " where SRV_Actif=true" _
        & " order by SRV_LNum, SRV_NumPere, SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste_OLD = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        If rs("SRV_NumPere").Value = 0 Then
            Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
        Else
            If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
                GoTo lab_suivant
            End If
            If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                Call ajouter_service(rs("SRV_NumPere").Value)
            End If
            Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
        End If
        If ndp.Children > 0 Then
            n = ndp.Children
            Set ndp_sav = ndp
            Set ndp = ndp.Child
            trouve = False
            For i = 1 To n
                If ndp.Text > rs("SRV_Nom").Value Then
                    mode = tvwPrevious
                    trouve = True
                    Exit For
                End If
                Set ndp = ndp.Next
            Next i
            If Not trouve Then
                Set ndp = ndp_sav
                mode = tvwChild
            End If
        Else
            mode = tvwChild
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               mode, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
'        nd.Sorted = True
        nd.tag = True
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    
    ' Les postes
    If g_smode <> "C" Then
        'Si l'on souhaite afficher les postes
        sql = "select PO_Num, PO_SRVNum, PO_Libelle, FT_Libelle" _
            & " from Poste, FctTrav" _
            & " where FT_Num=PO_FTNum" _
            & " and PO_Actif=true"
    '    If g_mode_acces = MODE_PARAM Then
    '        sql = sql & " and PO_LNum=" & p_NumLabo
    '    ElseIf g_tbl_site(0) > 0 Then
        If g_mode_acces <> MODE_PARAM Then
            If g_tbl_site(0) > 0 Then
                For i = 0 To UBound(g_tbl_site())
                    If i = 0 Then
                        sql = sql & " and ("
                    Else
                        sql = sql & " or"
                    End If
                    sql = sql & " PO_LNum=" & g_tbl_site(i)
                Next i
                sql = sql + ")"
            End If
        End If
        sql = sql & " order by PO_SRVNum, PO_Libelle"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            afficher_liste_OLD = P_ERREUR
            Exit Function
        End If
        numsrv = 0
        While Not rs.EOF
            afficher = True
            If rs("PO_SRVNum").Value <> numsrv Then
                numsrv = rs("PO_SRVNum").Value
                If TV_NodeExiste(tv, "S" & rs("PO_SRVNum").Value, ndp) = P_OUI Then
                    If ndp.Children > 0 Then
                        Set ndp = ndp.Child
                        mode = tvwPrevious
                    Else
                        mode = tvwChild
                    End If
                Else
                    afficher = False
                End If
            Else
                Set ndp = nd
                mode = tvwNext
            End If
            sfct = rs("FT_Libelle").Value
            If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                sfct = sfct & " *"
            End If
            If afficher Then
                Set nd = tv.Nodes.Add(ndp, _
                                       mode, _
                                       "P" & rs("PO_Num").Value, _
                                       sfct, _
                                       IMG_POSTE, _
                                       IMG_POSTE)
                nd.tag = True
            End If
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    ' Met en évidence les noeuds 'retenus'
    If g_mode_acces = MODE_SELECT And g_plusieurs Then
        For i = 0 To nsel
            n = STR_GetNbchamp(g_lignes(i).texte, ";")
            s = STR_GetChamp(g_lignes(i).texte, ";", n - 1)
            If s = "0" Then
                s = "S0"
            End If
            If left$(s, 1) = "S" Then
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_SRV_SEL
                Else
                    img = IMG_SRV_SEL_NOMOD
                End If
            Else
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_POSTE_SEL
                Else
                    img = IMG_POSTE_SEL_NOMOD
                End If
            End If
            If s = "S0" Then
                Set nd = tv.Nodes(1)
            Else
                Set nd = tv.Nodes(s)
            End If
            nd.image = img
            nd.tag = fmodif
            nd.SelectedImage = img
            If s <> "S0" Then
                Set ndp = nd.Parent
                While left$(ndp.key, 1) <> "L"
                    ndp.Expanded = True
                    Set ndp = ndp.Parent
                Wend
            End If
        Next i
    End If
    
    tv.SetFocus
    g_mode_saisie = True
    
    Set nd = Nothing
    Set ndp = Nothing
    Set ndp_sav = Nothing
    
    afficher_liste_OLD = P_OK
    
End Function

Private Function afficher_liste2() As Integer

    Dim sql As String, s As String, codsite As String, nomsrv As String, sfct As String
    Dim trouve As Boolean, faff As Boolean, afficher As Boolean
    Dim img As Integer, i As Integer, nsel As Integer, n As Integer, j As Integer
    Dim mode As Integer
    Dim numsrv As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    g_mode_saisie = False
    
    tv.Nodes.Clear
    
    If g_numserv = 0 Then
        sql = "select L_Code from Laboratoire" _
            & " where L_Num=" & g_numsite _
            & " order by L_Code"
        If Odbc_RecupVal(sql, codsite) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set nd = tv.Nodes.Add(, , "L" & g_numsite, codsite)
        nd.Expanded = True
        nd.Sorted = True
        ' Les services
        sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
            & " order by SRV_LNum, SRV_NumPere, SRV_Nom"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        While Not rs.EOF
            If rs("SRV_NumPere").Value = 0 Then
                Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
            Else
                If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
                    GoTo lab_suivant
                End If
                If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                    Call ajouter_service(rs("SRV_NumPere").Value)
                End If
                Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
            End If
            If ndp.Children > 0 Then
                n = ndp.Children
                Set ndp_sav = ndp
                Set ndp = ndp.Child
                trouve = False
                For i = 1 To n
                    If ndp.Text > rs("SRV_Nom").Value Then
                        mode = tvwPrevious
                        trouve = True
                        Exit For
                    End If
                    Set ndp = ndp.Next
                Next i
                If Not trouve Then
                    Set ndp = ndp_sav
                    mode = tvwChild
                End If
            Else
                mode = tvwChild
            End If
            Set nd = tv.Nodes.Add(ndp, _
                                   mode, _
                                   "S" & rs("SRV_Num").Value, _
                                   rs("SRV_Nom").Value, _
                                   IMG_SRV, _
                                   IMG_SRV)
'            nd.Sorted = True
            nd.tag = True
lab_suivant:
            rs.MoveNext
        Wend
        rs.Close
    Else
        ' Les services
        sql = "select SRV_Nom from Service" _
            & " where SRV_Num=" & g_numserv _
            & " order by SRV_Nom"
        If Odbc_RecupVal(sql, nomsrv) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set nd = tv.Nodes.Add(, _
                               tvwChild, _
                               "S" & g_numserv, _
                               nomsrv, _
                               IMG_SRV, _
                               IMG_SRV)
        nd.Expanded = True
 '       nd.Sorted = True
        nd.tag = True
        Call ajouter_fils(g_numserv)
    End If
    
    ' Les postes
    sql = "select PO_Num, PO_SRVNum, PO_Libelle, FT_Libelle" _
        & " from Poste, FctTrav" _
        & " where FT_Num=PO_FTNum" _
        & " and PO_Actif=true" _
        & " order by PO_SRVNum, PO_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste2 = P_ERREUR
        Exit Function
    End If
    numsrv = 0
    While Not rs.EOF
        afficher = True
        If rs("PO_SRVNum").Value <> numsrv Then
            If TV_NodeExiste(tv, "S" & rs("PO_SRVNum").Value, ndp) = P_OUI Then
                numsrv = rs("PO_SRVNum").Value
                If ndp.Children > 0 Then
                    Set ndp = ndp.Child
                    mode = tvwPrevious
                Else
                    mode = tvwChild
                End If
            Else
                afficher = False
            End If
        Else
            Set ndp = nd
            mode = tvwNext
        End If
        If afficher Then
            sfct = rs("FT_Libelle").Value
            If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                sfct = sfct & " *"
            End If
            Set nd = tv.Nodes.Add(ndp, _
                                   mode, _
                                   "P" & rs("PO_Num").Value, _
                                   sfct, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
            nd.tag = True
            nd.Expanded = True
            nd.Sorted = True
        End If
        rs.MoveNext
    Wend
    rs.Close

    ' Les personnes associées
    For i = 1 To tv.Nodes.Count
        Set ndp = tv.Nodes(i)
        If left$(ndp.key, 1) = "P" Then
            sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
                & " where U_SPM like '%" & ndp.key & ";%' and U_Actif=true"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                afficher_liste2 = P_ERREUR
                Exit Function
            End If
            While Not rs.EOF
                Set nd = tv.Nodes.Add(ndp, _
                                       tvwChild, _
                                       "", _
                                       rs("U_Nom").Value & " " & rs("U_Prenom").Value, _
                                       IMG_UTIL, _
                                       IMG_UTIL)
                nd.tag = "U" & rs("U_Num").Value
                rs.MoveNext
            Wend
            rs.Close
        End If
    Next i
    
    tv.SetFocus
    g_mode_saisie = True
    
    afficher_liste2 = P_OK
    
End Function

Private Function afficher_liste3(v_rech As String) As Integer

    Dim sql As String, s As String, sfct As String, lib As String
    Dim libNiveau As String, condRech As String, op As String
    Dim smot As String, stag As String
    Dim fmodif As Boolean, afficher As Boolean, trouve As Boolean
    Dim img As Integer, i As Integer, nsel As Integer, n As Integer, j As Integer
    Dim mode As Integer, imot As Integer
    Dim numsrv As Long, lnb As Long, num As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    Dim tb()
    
    g_mode_saisie = False
    
    ' Les services
    sql = "select SRV_Num, SRV_Site, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true [CONDITION_NIVEAU] [CONDITION_RECHERCHE]" _
        & " order by SRV_Nom"
    
    If Me.CmbNiveau.ListIndex <= 0 Then
        sql = Replace(sql, "[CONDITION_NIVEAU]", "")
    Else
        If Odbc_SelectV("select Nivs_num from niveau_structure" _
                        & " Where Nivs_Num=" & Me.CmbNiveau.ItemData(Me.CmbNiveau.ListIndex), rs2) = P_ERREUR Then
            Call quitter
            Exit Function
        Else
            If rs2.EOF Then
                Call quitter
                Exit Function
            Else
                sql = Replace(sql, "[CONDITION_NIVEAU]", " and SRV_NivsNum=" & rs2("Nivs_Num") & " ")
            End If
        End If
    End If
    
    op = ""
    condRech = ""
    op = ""
    For imot = 0 To STR_GetNbchamp(Me.TxtRecherche.Text, " ")
        smot = Trim(STR_GetChamp(Me.TxtRecherche.Text, " ", imot))
        smot = LCase(STR_Phonet(smot))
        If smot <> "" Then
            condRech = condRech & op & " ( translate(lower(SRV_Nom),'éèàçù', 'eeacu') like '%" & smot & "%' or translate(lower(SRV_code),'éèàçù', 'eeacu') like '%" & smot & "%' )"
            op = " And "
        End If
        'Debug.Print condRech
    Next imot
    sql = Replace(sql, "[CONDITION_RECHERCHE]", " and (" & condRech & ")")
    
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        g_mode_saisie = True
        afficher_liste3 = P_ERREUR
        Exit Function
    End If
    
    If rs.EOF Then
        If Me.CmbNiveau.ListIndex <= 0 Then
            MsgBox "Aucun  trouvé"
        Else
            MsgBox "Aucun '" & Me.CmbNiveau.Text & "' trouvé"
        End If
        Me.TxtRecherche.Text = ""
        g_mode_saisie = True
        Exit Function
    End If
            
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "select EJ_Num, EJ_Nom from EntJuridique"
    If Odbc_RecupVal(sql, num, lib) = P_ERREUR Then
        afficher_liste3 = P_ERREUR
        Exit Function
    End If
    p_NumSite = p_num_ent_juridique
    Set ndp = tv.Nodes.Add(, , "L" & p_NumSite, lib)
    ndp.Expanded = True
    
    If Not rs.EOF Then
        rs.MoveLast
        PgBar.Visible = True
        PgBar.max = rs.RowCount
        PgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        PgBar.Value = PgBar.Value + 1
        If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
            GoTo lab_suivant
        End If
        If rs("SRV_NumPere").Value = 0 Then
            'Set ndp = tv.Nodes("L" & ps_num_first_site(rs("SRV_Site").Value))
            Set ndp = tv.Nodes("L" & p_num_ent_juridique)
        Else
            If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                Call ajouter_service(rs("SRV_NumPere").Value)
            End If
            Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
        End If
        
        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        nd.tag = True & "|" & True
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    PgBar.Visible = False
    
    ' on les ouvre tous
    For n = 1 To tv.Nodes.Count
        Set ndp = tv.Nodes(n)
        ' Les postes
        If g_smode <> "C" Then
            If left$(ndp.key, 1) = "S" Then
                If ndp.Children > 0 Then
                    sql = "select PO_Num, PO_Libelle, FT_Libelle" _
                        & " from Poste, FctTrav" _
                        & " where FT_Num=PO_FTNum" _
                        & " and PO_Actif=true" _
                        & " and PO_SRVNum=" & Mid$(ndp.key, 2) _
                        & " order by PO_Libelle desc"
                    If Odbc_SelectV(sql, rs) = P_ERREUR Then
                        afficher_liste3 = P_ERREUR
                        Exit Function
                    End If
                    While Not rs.EOF
                        sfct = rs("FT_Libelle").Value
                        If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                            sfct = sfct & " *"
                        End If
                        Set nd = tv.Nodes.Add(ndp.Child, _
                                               tvwPrevious, _
                                               "P" & rs("PO_Num").Value, _
                                               sfct, _
                                               IMG_POSTE, _
                                               IMG_POSTE)
                        nd.tag = True
                        rs.MoveNext
                    Wend
                    ndp.Expanded = True
                    rs.Close
                Else
                    sql = "select count(*) from Poste" _
                        & " where PO_Actif=true" _
                        & " and PO_SRVNum=" & Mid$(ndp.key, 2)
                    If Odbc_Count(sql, lnb) = P_ERREUR Then
                        afficher_liste3 = P_ERREUR
                        Exit Function
                    End If
                    If lnb > 0 Then
                        Set nd = tv.Nodes.Add(ndp, _
                                           tvwChild, _
                                           , _
                                           "A charger")
                        stag = ndp.tag
                        Call STR_PutChamp(stag, "|", 1, False)
                        ndp.tag = stag
                    End If
                End If
            End If
        End If
    Next n
    
    ' Met en évidence les noeuds 'retenus'
    If g_mode_acces = MODE_SELECT And g_plusieurs Then
        For i = 0 To nsel
            n = STR_GetNbchamp(g_lignes(i).texte, ";")
            s = STR_GetChamp(g_lignes(i).texte, ";", n - 1)
            If left$(s, 1) = "S" Then
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_SRV_SEL
                Else
                    img = IMG_SRV_SEL_NOMOD
                End If
            Else
                fmodif = g_lignes(i).fmodif
                If fmodif Then
                    img = IMG_POSTE_SEL
                Else
                    img = IMG_POSTE_SEL_NOMOD
                End If
            End If
            If TV_NodeExiste(tv, s, nd) = P_OUI Then
                nd.image = img
                nd.tag = fmodif
                nd.SelectedImage = img
                If left$(nd.key, 1) <> "L" Then
                    Set ndp = nd.Parent
                    While left$(ndp.key, 1) <> "L"
                        ndp.Expanded = True
                        Set ndp = ndp.Parent
                    Wend
                End If
            End If
        Next i
    End If
    
    Call ouvrir_serv_poste
    
    tv.SetFocus
    g_mode_saisie = True
    
    Set nd = Nothing
    Set ndp = Nothing
    Set ndp_sav = Nothing
    
    PgBar.Visible = False
    
    afficher_liste3 = P_OK

End Function


Private Sub afficher_menu(ByVal v_bclavier As Boolean)

    Dim key As String, tag As String, libresp As String, sql As String, libposte As String
    Dim numposte As Long, numresp As Long
    
    key = tv.SelectedItem.key
    Select Case left$(key, 1)
    Case "L"
        mnuCreerS.Visible = g_crspm_autor
        mnuSepCrS.Visible = g_crspm_autor
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "S"
        mnuCreerS.Visible = g_crspm_autor
        mnuSepCrS.Visible = g_crspm_autor
        mnuModS.Visible = g_modspm_autor
        mnuSuppS.Visible = g_supspm_autor
        mnuSepMSS.Visible = IIf(g_modspm_autor Or g_supspm_autor, True, False)
        mnuCreerP.Visible = g_crspm_autor
        mnuSepCrP.Visible = g_crspm_autor
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        If tv.SelectedItem.Index = tv.SelectedItem.Root.Index Then
            mnuDepl.Visible = False
            mnuSepDepl.Visible = False
        Else
            mnuDepl.Visible = g_modspm_autor
            mnuSepDepl.Visible = g_modspm_autor
        End If
        mnuVoirPers.Visible = True
        mnuSepVoirPers.Visible = True
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "P"
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = True
        numposte = Mid$(key, 2)
        sql = "select PO_NumResp, PO_Libelle from Poste" _
            & " where PO_Num=" & numposte
        If Odbc_RecupVal(sql, numposte, libposte) = P_ERREUR Then
            libresp = "???"
        ElseIf numposte > 0 Then
            sql = "select FT_Libelle from Poste, FctTrav" _
                & " where PO_Num=" & numposte _
                & " and FT_Num=PO_FTNum"
            If Odbc_RecupVal(sql, libresp) = P_ERREUR Then
                libresp = "???"
            End If
        ElseIf numposte = 0 Then
            libresp = "Est responsable"
        Else
            libresp = "NON RENSEIGNE"
        End If
        libresp = "--- Poste responsable : " & libresp & " ---"
        mnuPosteResp.Caption = libresp
        mnuLibPoste.Visible = True
        mnuLibPoste.Caption = "Poste : " & libposte
        mnuSuppP.Visible = g_supspm_autor
        mnuSepSuppP.Visible = True
        mnuDepl.Visible = g_modspm_autor
        mnuSepDepl.Visible = g_modspm_autor
        If g_mode_acces = MODE_PARAM_PERS Then
            mnuVoirPers.Visible = False
            mnuSepVoirPers.Visible = False
        Else
            If tv.SelectedItem.Children > 0 Then
                mnuVoirPers.Visible = False
                mnuSepVoirPers.Visible = False
            Else
                mnuVoirPers.Visible = True
                mnuSepVoirPers.Visible = True
            End If
        End If
        If g_mode_acces = MODE_PARAM_PERS Or g_prmpers Then
            mnuCrPers.Visible = g_crutil_autor
            mnuSepCrPers.Visible = g_crutil_autor
        Else
            mnuCrPers.Visible = False
            mnuSepCrPers.Visible = False
        End If
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case Else
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        If g_mode_acces = MODE_PARAM_PERS Or g_prmpers Then
            mnuModPers.Visible = g_modutil_autor
            mnuSepModPers.Visible = g_modutil_autor
        Else
            mnuModPers.Visible = False
            mnuSepModPers.Visible = False
        End If
    End Select
    
    If v_bclavier Then
        Call PopupMenu(mnuFct, , tv.left, tv.Top)
    Else
        Call PopupMenu(mnuFct)
    End If
    
End Sub

Private Function ajouter_fils(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    sql = "select SRV_Num, SRV_Nom from Service" _
        & " where SRV_NumPere=" & v_numsrv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_fils = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Set ndp = tv.Nodes("S" & v_numsrv)
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
'        nd.Sorted = True
        nd.tag = True
        Call ajouter_fils(rs("SRV_Num").Value)
        rs.MoveNext
    Wend
    rs.Close
    
    ajouter_fils = P_OK
    
End Function

Private Sub ajouter_pers_tv()

    Dim sql As String, sposte As String, s As String
    Dim i As Integer, n   As Integer
    Dim spm As Variant
    Dim nd As Node, ndp As Node
    Dim rs As rdoResultset
    
    sql = "select U_Num, U_Nom, U_Prenom, U_SPM from Utilisateur" _
        & " where U_SPM like '%" & tv.SelectedItem.key & ";%' and U_Actif=true" _
        & " order by U_Nom, U_Prenom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
            spm = rs("U_SPM").Value
            n = STR_GetNbchamp(spm, "|")
            For i = 0 To n - 1
                s = STR_GetChamp(spm, "|", i)
                If InStr(s, tv.SelectedItem.key + ";") > 0 Then
                    sposte = STR_GetChamp(s, ";", STR_GetNbchamp(s, ";") - 1)
                    Set ndp = tv.Nodes(sposte)
                    Set nd = tv.Nodes.Add(ndp, _
                                           tvwChild, _
                                           "", _
                                           rs("U_Nom").Value & " " & rs("U_Prenom").Value, _
                                           IMG_UTIL, _
                                           IMG_UTIL)
                    nd.tag = "U" & rs("U_Num").Value
                    ndp.Expanded = True
                    Set ndp = ndp.Parent
                    While left$(ndp.key, 1) <> "L"
                        ndp.Expanded = True
                        Set ndp = ndp.Parent
                    Wend
                End If
            Next i
            rs.MoveNext
        Wend
    End If
    rs.Close

End Sub

Private Function ajouter_service(ByVal v_numsrv As Long) As Integer

    Dim sql As String, lib As String
    Dim trouve As Boolean
    Dim mode As Integer, i As Integer, n As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    Dim libNiveau As String
    
    If v_numsrv = 0 Then
        ajouter_service = P_OK
        Exit Function
    End If
    
    If TV_NodeExiste(tv, "S" & v_numsrv, nd) = P_OUI Then
        ajouter_service = P_OK
        Exit Function
    End If
    
    sql = "select SRV_NumPere, SRV_Nom, SRV_NivsNum from Service" _
        & " where SRV_Num=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_service = P_ERREUR
        Exit Function
    End If
    If rs("SRV_NumPere").Value > 0 Then
        If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
            Call ajouter_service(rs("SRV_NumPere").Value)
        End If
        Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
    Else
        Set ndp = tv.Nodes(1).Root
    End If
    lib = rs("SRV_Nom").Value
    libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
    If libNiveau <> "" Then
        lib = lib & " (" & libNiveau & ")"
    End If
    Set nd = tv.Nodes.Add(ndp, _
                           tvwChild, _
                           "S" & v_numsrv, _
                           lib, _
                           IMG_SRV, _
                           IMG_SRV)
    nd.tag = True & "|" & True
    
    ajouter_service = P_OK
    
End Function

Private Function ajouter_service_OLD(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim trouve As Boolean
    Dim mode As Integer, i As Integer, n As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
        & " where SRV_Num=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_service_OLD = P_ERREUR
        Exit Function
    End If
    If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
        ajouter_service_OLD = P_OK
        Exit Function
    End If
    If rs("SRV_NumPere").Value = 0 Then
        Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
    Else
        If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
            Call ajouter_service(rs("SRV_NumPere").Value)
        End If
        Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
    End If
    If ndp.Children > 0 Then
        n = ndp.Children
        Set ndp_sav = ndp
        Set ndp = ndp.Child
        trouve = False
        For i = 1 To n
            If ndp.Text > rs("SRV_Nom").Value Then
                mode = tvwPrevious
                trouve = True
                Exit For
            End If
            Set ndp = ndp.Next
        Next i
        If Not trouve Then
            Set ndp = ndp_sav
            mode = tvwChild
        End If
    Else
        mode = tvwChild
    End If
    Set nd = tv.Nodes.Add(ndp, _
                           mode, _
                           "S" & rs("SRV_Num").Value, _
                           rs("SRV_Nom").Value, _
                           IMG_SRV, _
                           IMG_SRV)
'    nd.Sorted = True
    nd.tag = True
    
    ajouter_service_OLD = P_OK
    
End Function

Private Sub basculer_selection()

    Dim img As Long
    
    If tv.SelectedItem.tag = False Then
        Exit Sub
    End If
    
    Select Case tv.SelectedItem.SelectedImage
    Case IMG_SRV
        If left$(tv.SelectedItem.key, 1) = "L" Then
        ElseIf InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
            Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_SRV_SEL
    Case IMG_POSTE
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
            Call MsgBox("Vous ne pouvez pas sélectionner un poste.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_POSTE_SEL
    Case IMG_SRV_SEL
        img = IMG_SRV
    Case IMG_POSTE_SEL
        img = IMG_POSTE
    End Select
    
    tv.SelectedItem.SelectedImage = img
    tv.SelectedItem.image = img
    
End Sub

Private Sub creer_fonction(ByRef r_Num As Long, _
                           ByRef r_lib As String)

    Dim sret As String
    Dim frm As Form
    
    Set frm = KS_PrmFonction
    sret = KS_PrmFonction.AppelFrm(True)
    Set frm = Nothing
    If sret = "" Then
        r_Num = 0
        Exit Sub
    End If
    r_Num = STR_GetChamp(sret, "|", 0)
    r_lib = STR_GetChamp(sret, "|", 1)
    
End Sub

Private Function creer_poste() As Integer
    
    Dim sql As String, lib As String
    Dim trouve As Boolean
    Dim i As Integer, nbenf As Integer, n As Integer, btn_sortie As Integer
    Dim nfct As Integer, ie As Integer
    Dim numsrv As Long, num As Long, tbl_fct() As Long, numlabo As Long, lnb As Long
    Dim numfct
    Dim rs As rdoResultset
    Dim nd As Node, nde As Node, ndp As Node
    
    Set nd = tv.SelectedItem
    numsrv = Mid$(nd.key, 2)
    nbenf = nd.Children
    nfct = -1
    Set nde = nd.Child
    For i = 1 To nbenf
        If left$(nde.key, 1) = "P" Then
            num = Mid$(nde.key, 2)
            sql = "select PO_FTNum from Poste" _
                & " where PO_Num=" & num
            If Odbc_RecupVal(sql, numfct) = P_ERREUR Then
                creer_poste = P_ERREUR
                Exit Function
            End If
            nfct = nfct + 1
            ReDim Preserve tbl_fct(nfct) As Long
            tbl_fct(nfct) = numfct
        End If
        Set nde = nde.Next
    Next i
    
    Call CL_Init
    
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        creer_poste = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        trouve = False
        For i = 0 To nfct
            If tbl_fct(i) = rs("FT_Num").Value Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
            n = n + 1
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 And Not g_crfct_autor Then
        Call MsgBox("Aucune fonction ne peut être ajoutée à ce service.", vbInformation + vbOKOnly, "")
        creer_poste = P_OK
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    If g_crfct_autor Then
        Call CL_AddBouton("&Créer une fonction", "", 0, 0, 1800)
        btn_sortie = 2
    Else
        btn_sortie = 1
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    Call CL_AffiSelFirst
    Call CL_InitResteCachée(True)
lab_choix:
    ChoixListe.Show 1
    ' Sortie
    If CL_liste.retour = btn_sortie Then
        GoTo lab_fin
    End If
    
    ' Création
    If CL_liste.retour = 1 Then
        Call creer_fonction(num, lib)
        If num > 0 Then
            Call CL_AddLigne(lib, num, "", True)
            n = n + 1
        End If
        GoTo lab_choix
    End If
    
    ' Ajout des sélectionnés
'    Call TV_FirstParent(nd, ndp)
'    numlabo = Mid$(ndp.key, 2)
    ' Ca ne marchait pas si on avait sélectionné un service et non tout le site
    ' -> numlabo se retrouvait avec le numserv
    numlabo = 1
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
            Call Odbc_AddNew("Poste", _
                             "PO_Num", _
                             "po_seq", _
                             True, _
                             num, _
                             "PO_SRVNum", numsrv, _
                             "PO_FTNum", CL_liste.lignes(i).num, _
                             "PO_Libelle", CL_liste.lignes(i).texte, _
                             "PO_NumResp", -1, _
                             "PO_LNum", numlabo, _
                             "PO_Actif", True)
            Set nde = tv.Nodes.Add(nd, _
                                   tvwChild, _
                                   "P" & num, _
                                   CL_liste.lignes(i).texte, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
        End If
    Next i
    
lab_fin:
    Unload ChoixListe
    creer_poste = P_OK

End Function

Private Function creer_service() As Integer

    Dim code As String, lib As String, libcourt As String
    Dim first_chp As Integer
    Dim num As Long, numpere As Long, numlabo As Long, lnb As Long
    Dim nd As Node, ndp As Node
    
    Set nd = tv.SelectedItem
    
    code = ""
    lib = ""
    libcourt = ""
    
lab_saisie:
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Service", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    If left$(nd.key, 1) = "S" Then
        Call SAIS_AddChamp("Rattaché à", -50, 50, 0, True, nd.Text)
        first_chp = 1
        numpere = Mid$(nd.key, 2)
    Else
        first_chp = 0
        numpere = 0
    End If
    Call SAIS_AddChamp("Code", 8, 8, SAIS_TYP_TOUT_CAR, True, code)
    Call SAIS_AddChamp("Nom", 50, 50, SAIS_TYP_TOUT_CAR, False, lib)
    Call SAIS_AddChamp("Nom court", 30, 30, SAIS_TYP_TOUT_CAR, True, libcourt)
    Saisie.Show 1
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        creer_service = P_NON
        Exit Function
    End If
    code = SAIS_Saisie.champs(first_chp).sval
    lib = SAIS_Saisie.champs(first_chp + 1).sval
    libcourt = SAIS_Saisie.champs(first_chp + 2).sval
    If code <> "" Then
        Call Odbc_Count("select count(*) from service where srv_code=" & Odbc_String(SAIS_Saisie.champs(first_chp).sval), lnb)
        If lnb > 0 Then
            Call MsgBox("Le code '" & code & "' est déjà attribué." & vbCrLf & vbCrLf & "Veuillez choisir un autre code.", vbInformation + vbOKOnly, "")
            GoTo lab_saisie
        End If
    End If
    
    Call TV_FirstParent(nd, ndp)
    numlabo = Mid$(ndp.key, 2)
    Call Odbc_AddNew("Service", _
                     "SRV_Num", _
                     "srv_seq", _
                     True, _
                     num, _
                     "SRV_Code", code, _
                     "SRV_Nom", lib, _
                     "SRV_libcourt", libcourt, _
                     "SRV_NumPere", numpere, _
                     "SRV_LNum", numlabo)
    
'    nd.Sorted = True
    nd.Expanded = True
    Set nd = tv.Nodes.Add(nd, tvwChild, "S" & num, lib, IMG_SRV, IMG_SRV)
    Set tv.SelectedItem = nd
    SendKeys "{DOWN}"
    SendKeys "{UP}"
    DoEvents
    Set tv.SelectedItem = nd
    
    creer_service = P_OUI
    
End Function

Private Function deplacer_sp() As Integer

    Dim key As String, sql As String, stype_src As String, stype_dest As String
    Dim s_sp_src As String, s_sp_dest As String, s_sp As String, s As String
    Dim key_depl As String
    Dim encore As Boolean
    Dim i As Integer, nbch As Integer
    Dim numsrv_src As Long, numsrv_dest As Long, lnb As Long
    Dim nd_src As Node, nd_dest As Node, nd As Node, ndp As Node
    Dim rs As rdoResultset
    
    lblDepl.Visible = False
    
    Set nd_dest = tv.SelectedItem
    Set nd_src = tv.Nodes(g_pos_depl)
    g_pos_depl = 0
    
    If nd_src.key = nd_dest.key Then
        Call MsgBox("Vous ne pouvez pas rattacher l'objet à lui-même !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    
    stype_src = left$(nd_src.key, 1)
    stype_dest = left$(nd_dest.key, 1)
    If stype_dest = "P" Then
        Call MsgBox("Vous ne pouvez pas sélectionner un poste !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    If stype_src = "P" And stype_dest = "L" Then
        Call MsgBox("Vous ne pouvez déplacer un poste que dans un service !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    
    ' On regarde si nd_dest n'est pas enfant de nd_src !
    If left$(nd_dest.key, 1) <> "L" Then
        Set nd = nd_dest.Parent
        While left$(nd.key, 1) <> "L"
            If nd.Index = nd_src.Index Then
                Call MsgBox("Vous ne pouvez déplacer le service dans un des ses fils !", vbExclamation + vbOKOnly, "")
                deplacer_sp = P_OK
                Exit Function
            End If
            Set nd = nd.Parent
        Wend
        numsrv_dest = CLng(Mid$(nd_dest.key, 2))
    Else
        numsrv_dest = 0
    End If
    numsrv_src = CLng(Mid$(nd_src.key, 2))
    
    key_depl = nd_src.key
    s_sp_src = ""
    Set nd = nd_src
    While left$(nd.key, 1) <> "L"
        s_sp_src = nd.key & ";" & s_sp_src
        Set nd = nd.Parent
    Wend
        
    s_sp_dest = nd_src.key & ";"
    Set nd = nd_dest
    While left$(nd.key, 1) <> "L"
        s_sp_dest = nd.key & ";" & s_sp_dest
        Set nd = nd.Parent
    Wend
    
    If Odbc_BeginTrans() = P_ERREUR Then
        deplacer_sp = P_ERREUR
        Exit Function
    End If
    ' Mise à jour Service : SRV_NumPere
    If stype_src = "P" Then
        If Odbc_Update("Poste", _
                        "PO_Num", _
                        "where PO_Num=" & numsrv_src, _
                        "PO_SRVNum", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    Else
        If Odbc_Update("Service", _
                        "SRV_Num", _
                        "where SRV_Num=" & numsrv_src, _
                        "SRV_NumPere", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    End If
    ' Mise à jour Documentation
    sql = "select DO_Num, DO_Dest from Documentation" _
        & " where DO_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("DO_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Documentation", _
                        "DO_Num", _
                        "where DO_Num=" & rs("DO_Num").Value, _
                        "DO_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Dossier
    sql = "select DS_Num, DS_Dest from Dossier" _
        & " where DS_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("DS_Dest").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("Dossier", _
                        "DS_Num", _
                        "where DS_Num=" & rs("DS_Num").Value, _
                        "DS_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Document
    sql = "select D_Num, D_Dest from Document" _
        & " where D_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("D_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Document", _
                        "D_Num", _
                        "where D_Num=" & rs("D_Num").Value, _
                        "D_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour GroupeUtil
    sql = "select GU_Num, GU_Lst from GroupeUtil" _
        & " where GU_Lst like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("GU_Lst").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("GroupeUtil", _
                        "GU_Num", _
                        "where GU_Num=" & rs("GU_Num").Value, _
                        "GU_Lst", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Utilisateur
    sql = "select U_Num, U_SPM from Utilisateur" _
        & " where U_Spm like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        nbch = STR_GetNbchamp(rs("U_SPM").Value, "|")
        s_sp = ""
        For i = 1 To nbch
            s = STR_GetChamp(rs("U_SPM").Value, "|", i - 1)
            s = Replace(s, s_sp_src, s_sp_dest)
            s_sp = s_sp + s + "|"
        Next i
        If Odbc_Update("Utilisateur", _
                        "U_Num", _
                        "where U_Num=" & rs("U_Num").Value, _
                        "U_SPM", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close

    Call Odbc_CommitTrans
    
' AVANT
'    If g_mode_acces <> MODE_PARAM_PERS Then
'        Call afficher_liste
'    Else
'        Call afficher_liste2
'    End If

' Remplacé par
    Set nd_src.Parent = nd_dest
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.Nodes(key_depl).EnsureVisible
    deplacer_sp = P_OK
    Exit Function
' ***

' AVANT
    ' Se repositionne sur le "SP" déplacé
    For i = 1 To tv.Nodes.Count
        If tv.Nodes(i).key = key_depl Then
            Set ndp = tv.Nodes(i).Parent
            While left$(ndp.key, 1) <> "L"
                ndp.Expanded = True
                Set ndp = ndp.Parent
            Wend
        End If
    Next i
    
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.SetFocus
    SendKeys "{DOWN}"
    SendKeys "{UP}"
' ****

    deplacer_sp = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    deplacer_sp = P_ERREUR
    
End Function

Private Sub imprimer()

    
End Sub

Private Sub initialiser()

    Dim i As Integer, nbchp As Integer, nsel As Integer
    Dim sql As String, rs As rdoResultset
    Dim lnb As Long
    
    g_crfct_autor = False
    g_crutil_autor = False
    g_modutil_autor = False
    g_crspm_autor = False
    g_modspm_autor = False
    g_supspm_autor = False
    
    g_node_crt = 1
    
    Dim numL As Integer, codeL As String
    
    If g_mode_acces = MODE_SELECT Then
        'Call Odbc_RecupVal("select L_Num, L_Code from Laboratoire order by L_Code", numL, codeL)
        'p_numlabo = numL
        nsel = -1
        On Error Resume Next
        nsel = UBound(CL_liste.lignes)
        On Error GoTo 0
        If nsel >= 0 Then
            ReDim g_lignes(nsel) As CL_SLIGNE
            For i = 0 To nsel
                g_lignes(i) = CL_liste.lignes(i)
            Next i
        Else
            Erase g_lignes()
        End If
        If g_ssite = "" Then
            ReDim g_tbl_site(0) As Long
            g_tbl_site(0) = 0
        Else
            nbchp = STR_GetNbchamp(g_ssite, ";")
            ReDim g_tbl_site(nbchp - 1) As Long
            For i = 0 To nbchp - 1
                g_tbl_site(i) = STR_GetChamp(g_ssite, ";", i)
            Next i
        End If
        cmd(CMD_IMPRIMER).Visible = False
    Else
        p_numlabo = 1
        cmd(CMD_OK).Visible = False
        cmd(CMD_IMPRIMER).left = cmd(CMD_OK).left
    End If
    
    ' charger le combo des niveaux
    g_ya_niveau = False
    sql = "select count(*) from niveau_structure"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
    End If
    
    If lnb > 0 Then
        g_ya_niveau = True
        Me.CmbNiveau.AddItem "Tous"
        Me.CmbNiveau.ItemData(Me.CmbNiveau.ListCount - 1) = 0
        sql = "select Nivs_Nom, Nivs_Num from niveau_structure Order by Nivs_Num"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter
            Exit Sub
        Else
            While Not rs.EOF
                Me.CmbNiveau.AddItem rs("Nivs_Nom")
                Me.CmbNiveau.ItemData(Me.CmbNiveau.ListCount - 1) = rs("Nivs_Num")
                rs.MoveNext
            Wend
        End If
        Me.CmbNiveau.ListIndex = 0
    Else
        Me.CmbNiveau.Visible = False
    End If
    If g_mode_acces <> MODE_PARAM_PERS Then
        If afficher_liste() = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
    Else
        If afficher_liste2() = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
    End If
    
    Set tv.SelectedItem = tv.Nodes(1)
    tv.SetFocus
    SendKeys "{PGDN}"
    SendKeys "{HOME}"
    DoEvents
    
End Sub

Private Function modifier_poste() As Integer

    Dim libposte As String, libfct As String, sql As String
    Dim nch As Integer
    Dim num As Long
    Dim nd As Node
    
    Set nd = tv.SelectedItem
    num = CLng(Mid$(nd.key, 2))
    sql = "select PO_Libelle, FT_Libelle from Poste, FctTrav" _
        & " where PO_Num=" & num _
        & " and FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, libposte, libfct) = P_ERREUR Then
        modifier_poste = P_ERREUR
        Exit Function
    End If
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Poste", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call SAIS_AddChamp("Fonction", -80, 80, 0, True, libfct)
    If left$(nd.Parent.key, 1) <> "L" Then
        Call SAIS_AddChamp("Service", -50, 50, 0, True, nd.Parent.Text)
        nch = 2
    Else
        nch = 1
    End If
    Call SAIS_AddChamp("Nom", 80, 80, SAIS_TYP_TOUT_CAR, False, libposte)
    Saisie.Show 1
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        modifier_poste = P_NON
        Exit Function
    End If
    libposte = SAIS_Saisie.champs(nch).sval
    
    Call Odbc_Update("Poste", _
                     "PO_Num", _
                     "where PO_Num=" & num, _
                     "PO_Libelle", libposte)
    If libposte <> libfct Then
        libfct = libfct & " *"
    End If
    nd.Text = libfct
    
    modifier_poste = P_OK
    
End Function

Private Function modifier_service() As Integer

    Dim code As String, lib As String, libcourt As String
    Dim first_chp As Integer
    Dim num As Long, lnb As Long
    Dim nd As Node
    
    Set nd = tv.SelectedItem
    num = CLng(Mid$(nd.key, 2))
    Call Odbc_RecupVal("select srv_code, srv_nom, srv_libcourt from service where srv_num=" & num, _
                        code, lib, libcourt)
    
lab_saisie:
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Service", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    If left$(nd.Parent.key, 1) <> "L" Then
        Call SAIS_AddChamp("Rattaché à", -50, 50, 0, True, nd.Parent.Text)
        first_chp = 1
    Else
        first_chp = 0
    End If
    Call SAIS_AddChamp("Code", 8, 8, SAIS_TYP_TOUT_CAR, True, code)
    Call SAIS_AddChamp("Nom", 50, 50, SAIS_TYP_TOUT_CAR, False, lib)
    Call SAIS_AddChamp("Nom court", 30, 30, SAIS_TYP_TOUT_CAR, True, libcourt)
    Saisie.Show 1
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        modifier_service = P_NON
        Exit Function
    End If
    code = SAIS_Saisie.champs(first_chp).sval
    lib = SAIS_Saisie.champs(first_chp + 1).sval
    libcourt = SAIS_Saisie.champs(first_chp + 2).sval
    
    If code <> "" Then
        Call Odbc_Count("select count(*) from service where srv_code=" & Odbc_String(SAIS_Saisie.champs(first_chp).sval) & " and srv_num<>" & num, lnb)
        If lnb > 0 Then
            Call MsgBox("Le code '" & code & "' est déjà attribué." & vbCrLf & vbCrLf & "Veuillez choisir un autre code.", vbInformation + vbOKOnly, "")
            GoTo lab_saisie
        End If
    End If
    
    Call Odbc_Update("Service", _
                     "SRV_Num", _
                     "where SRV_Num=" & num, _
                     "SRV_code", code, _
                     "SRV_Nom", lib, _
                     "SRV_libcourt", libcourt)
    nd.Text = lib
    
    modifier_service = P_OK
    
End Function

Private Function po_dans_histordoc(ByVal v_num As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from DocEtapeVersion" _
        & " where DEV_PONum=" & v_num
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        po_dans_histordoc = True
        Exit Function
    End If
    If lnb = 0 Then
        po_dans_histordoc = False
    Else
        po_dans_histordoc = True
    End If

End Function

Private Sub quitter()
    
    g_sret = ""
    
    Unload Me
    
End Sub

Private Sub selectionner()

    Dim sp As String, sm As String, s As String
    Dim encore As Boolean
    Dim n As Integer, i As Integer, j As Integer, nbch As Integer, img As Integer
    Dim nd As Node, ndp As Node
    
    If g_plusieurs Then
        n = 0
        For i = 2 To tv.Nodes.Count
            img = tv.Nodes(i).image
            If img = IMG_POSTE_SEL Or img = IMG_SRV_SEL Or img = IMG_POSTE_SEL_NOMOD Or img = IMG_SRV_SEL_NOMOD Then
                sp = tv.Nodes(i).key & ";"
                Set nd = tv.Nodes(i)
                encore = True
                Do
                    Set ndp = nd.Parent
                    s = ndp.key
                    If left$(s, 1) = "L" Then
                        encore = False
                    Else
                        sp = sp + s + ";"
                        Set nd = ndp
                    End If
                Loop Until Not encore
                s = ""
                nbch = STR_GetNbchamp(sp, ";")
                For j = nbch To 1 Step -1
                    s = s + STR_GetChamp(sp, ";", j - 1) + ";"
                Next j
                ReDim Preserve CL_liste.lignes(n)
                CL_liste.lignes(n).texte = s
                If img = IMG_POSTE_SEL Or img = IMG_SRV_SEL Then
                    CL_liste.lignes(n).tag = True
                Else
                    CL_liste.lignes(n).tag = False
                End If
                n = n + 1
            End If
        Next i
        
        ' Cas ou l'on souhaite pouvoir selectionner tous les services
        If g_smode = "C" Then
            If n = 0 Then
                If tv.Nodes.Item(1) = tv.SelectedItem Then
                    ReDim Preserve CL_liste.lignes(n)
                    CL_liste.lignes(n).texte = tv.SelectedItem
                    CL_liste.lignes(n).tag = True
                End If
            End If
        End If
        
        g_sret = "N" & n
    Else
        Set nd = tv.SelectedItem
        If InStr(g_stype, left$(nd.key, 1)) = 0 Or nd.key = "" Then
            If nd.key = "" Then
                Call MsgBox("Vous ne pouvez pas sélectionner une personne.", vbInformation + vbOKOnly, "")
            ElseIf left$(nd.key, 1) = "S" Then
                Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            ElseIf nd.key = "L1" Then   ' on a tout sélectionné
                g_sret = "0"
                Unload Me
            Else
                Call MsgBox("Vous ne pouvez pas sélectionner un poste.", vbInformation + vbOKOnly, "")
            End If
            Exit Sub
        End If
        sp = nd.key & ";"
        If left$(sp, 1) = "L" Then
            encore = False
        Else
            encore = True
        End If
        While encore
            Set ndp = nd.Parent
            s = ndp.key
            If left$(s, 1) = "L" Then
                encore = False
            Else
                sp = sp + s + ";"
                Set nd = ndp
            End If
        Wend
        s = ""
        nbch = STR_GetNbchamp(sp, ";")
        For j = nbch To 1 Step -1
            s = s + STR_GetChamp(sp, ";", j - 1) + ";"
        Next j
        g_sret = s
    End If
    
    Unload Me
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Function srvpo_dans_do(ByVal v_stype As String, _
                               ByVal v_num As Long, _
                               ByRef r_cas As Integer) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    r_cas = 0
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_do = True
        Exit Function
    End If

    r_cas = 1
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_do = True
        Exit Function
    End If

    r_cas = 2
    sql = "select count(*) from Document" _
        & " where D_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb = 0 Then
        srvpo_dans_do = False
    Else
        srvpo_dans_do = True
    End If

End Function

Private Function srvpo_dans_util(ByVal v_stype As String, _
                                 ByVal v_num As Long, _
                                 ByRef r_yaactif As Boolean) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Utilisateur" _
        & " where U_SPM like '%" & v_stype & v_num & ";%'" _
        & " and U_Actif=true"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    If lnb > 0 Then
        r_yaactif = True
        srvpo_dans_util = True
        Exit Function
    End If
    
    sql = "select count(*) from Utilisateur" _
        & " where U_SPM like '%" & v_stype & v_num & ";%'" _
        & " and U_Actif=false"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    If lnb > 0 Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    
    srvpo_dans_util = False

End Function

Private Function supprimer() As Integer

    Dim lib As String, sql As String, sobj As String, stype As String
    Dim ya_actif As Boolean
    Dim reponse As Integer, cas As Integer
    Dim num As Long, lnb As Long, numsrv As Long
    Dim nd As Node, ndr As Node
    
    num = CLng(Mid$(tv.SelectedItem.key, 2))
    stype = left$(tv.SelectedItem.key, 1)
    
    If stype = "S" Then
        sobj = "ce service"
    Else
        sobj = "ce poste"
    End If
    
    If p_appli_kalidoc > 0 Then
        If srvpo_dans_do(stype, num, cas) Then
            Select Case cas
            Case 0
                Call MsgBox("Une (ou plusieurs) documentation est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            Case 1
                Call MsgBox("Un (ou plusieurs) dossier est associé à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            Case 2
                Call MsgBox("Un (ou plusieurs) document est associé à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            End Select
            supprimer = P_OK
            Exit Function
        End If
    End If
    
    If srvpo_dans_util(stype, num, ya_actif) Then
        If ya_actif Then
            Call MsgBox("Une (ou plusieurs) personne est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
        Else
            Call MsgBox("Une (ou plusieurs) personne INACTIVE est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
        End If
        supprimer = P_OK
        Exit Function
    End If
    
    Select Case stype
    Case "S"
        If P_RecupSrvNom(num, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = "du service " & lib
    Case "P"
        sql = "select PO_SRVNum, FT_Libelle from Poste, FctTrav" _
            & " where PO_Num=" & num _
            & " and FT_Num=PO_FTNum"
        If Odbc_RecupVal(sql, numsrv, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = "du poste " & lib
        If P_RecupSrvNom(numsrv, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = sobj & " dans le service " & lib
    End Select
    
    reponse = MsgBox("Confirmez-vous la suppression " & sobj & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        supprimer = P_OK
        Exit Function
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    Set ndr = tv.SelectedItem
    Set nd = ndr
    Do
        num = Mid$(nd.key, 2)
        If left$(nd.key, 1) = "S" Then
            If Odbc_Delete("Service", "SRV_Num", "where SRV_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        Else
            ' supprimer les postes (dans synchro_kb) de ce service
            ' ou supprimer ce poste de la table synchro_kb
            If Odbc_Delete("Synchro_kd", "Synk_num", "WHERE Synk_pokd=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
            ' Poste dans l'historique des versions -> on l'inhibe
            If po_dans_histordoc(num) Then
                If Odbc_Update("Poste", "PO_Num", "where PO_Num=" & num, _
                               "PO_Actif", False) = P_ERREUR Then
                    GoTo err_enreg
                End If
            Else
                If Odbc_Delete("Poste", "PO_Num", "where PO_Num=" & num, lnb) = P_ERREUR Then
                    GoTo err_enreg
                End If
            End If
' Le poste doit être supprimé de toutes les références : formetape.fore_dest
        End If
    Loop Until Not TV_ChildNextParent(nd, ndr)
    
    If Odbc_CommitTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    tv.Nodes.Remove (ndr.Index)
    tv.Refresh
    
    supprimer = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    supprimer = P_ERREUR
    
End Function

Private Sub CmbNiveau_Click()
    If g_mode_saisie Then Call afficher_liste3(Me.TxtRecherche.Text)
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call selectionner
    Case CMD_IMPRIMER
        Call imprimer
    Case CMD_QUITTER
        Call quitter
    End Select
    
End Sub

Private Sub Form_Activate()

    If g_form_active Then
        Exit Sub
    End If
    
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyO And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If g_mode_acces = MODE_SELECT Then Call selectionner
    ElseIf (KeyCode = vbKeyI And Shift = vbAltMask) Or KeyCode = vbKeyF3 Then
        KeyCode = 0
        If g_mode_acces = MODE_PARAM Then Call imprimer
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_spm.htm")
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
        Exit Sub
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
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

Private Sub lblDepl_Click()

    g_pos_depl = 0
    lblDepl.Visible = False
    
End Sub

Private Sub mnuCreerP_Click()

    Call creer_poste
    
End Sub

Private Sub mnucreerS_Click()

    Call creer_service
    
End Sub

Private Sub mnuCrPers_Click()

    g_sret = "0|" + tv.SelectedItem.key
    Unload Me

End Sub

Private Sub mnuDepl_Click()

    Call activer_depl
    
End Sub

Private Sub mnuLibPoste_Click()

    Call modifier_poste
    
End Sub

Private Sub mnuModPers_Click()

    g_sret = Mid$(tv.SelectedItem.tag, 2)
    Unload Me
    
End Sub

Private Sub mnuModS_Click()

    Call modifier_service
    
End Sub

Private Sub mnuSuppP_Click()

    Call supprimer
    
End Sub

Private Sub mnuSuppS_Click()

    Call supprimer
    
End Sub

Private Sub mnuVoirPers_Click()

    Call ajouter_pers_tv
    
End Sub

Private Sub tv_Click()

    If g_node = tv.SelectedItem.Index And g_expand <> tv.SelectedItem.Expanded Then
        Exit Sub
    End If
    
    If g_button = vbRightButton Then
        Call afficher_menu(False)
    ElseIf g_button = vbLeftButton Then
        If g_pos_depl <> 0 Then
            If deplacer_sp() = P_ERREUR Then
                Call quitter
                Exit Sub
            End If
        ElseIf g_plusieurs Then
            Call basculer_selection
        End If
        Call majLibDetailSRV(tv.SelectedItem.key, tv.SelectedItem.tag)
    End If
        
End Sub

Private Sub tv_Expand(ByVal Node As ComctlLib.Node)

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    g_button = -1
    Me.LbldetailSRV.Visible = False
    If left$(Node.key, 1) = "S" Then
        If STR_GetChamp(Node.tag, "|", 1) = False Then
            tv.Nodes.Remove (Node.Child.Index)
            charger_service (Mid$(Node.key, 2))
            Call majLibDetailSRV(Node.key, Node.tag)
        End If
    End If
    
End Sub

Private Sub majLibDetailSRV(ByVal v_tv_key As String, ByVal v_tv_tag As String)
    Dim sql As String
    Dim srv_code As String, srv_stru_import As String
    Dim SRV_Num As Long
    Dim u_matricule As String
    Dim strSites As String
        
    LbldetailSRV.Visible = False
    If Mid(v_tv_key, 1, 1) = "S" Then
        sql = "select srv_code,srv_stru_import,srv_num from service where srv_num=" & Replace(v_tv_key, "S", "")
        Call Odbc_RecupVal(sql, srv_code, srv_stru_import, SRV_Num)
        LbldetailSRV.Visible = True
        LbldetailSRV.Caption = "Num=" & SRV_Num
        LbldetailSRV.Caption = LbldetailSRV.Caption & "   " & IIf(srv_code <> "", "Code " & srv_code, "")
        LbldetailSRV.Caption = LbldetailSRV.Caption & "   " & IIf(srv_stru_import <> "", "Import " & srv_stru_import, "")
        If p_MultiSite Then
            'strSites = ps_get_service_site("CODE", Replace(v_tv_key, "S", ""))
            LbldetailSRV.Caption = LbldetailSRV.Caption & " (Sites : " & strSites & ")"
        End If
        LbldetailSRV.Caption = Trim(LbldetailSRV.Caption)
    ElseIf Mid(v_tv_key, 1, 1) = "P" Then
        If p_MultiSite Then
            'strSites = ps_get_poste_site("LIB", Replace(v_tv_key, "P", ""))
            LbldetailSRV.Caption = "   (Sites : " & strSites & ")"
            LbldetailSRV.Visible = IIf(LbldetailSRV.Caption <> "", True, False)
        End If
    ElseIf Mid(v_tv_key, 1, 1) = "U" Then
        sql = "select u_matricule from utilisateur where u_num=" & Replace(v_tv_key, "U", "")
        Call Odbc_RecupVal(sql, u_matricule)
        LbldetailSRV.Visible = True
        LbldetailSRV.Caption = IIf(u_matricule <> "", "Matricule " & u_matricule, "")
        'strSites = ps_get_utilisateur_site("LIB", Replace(v_tv_key, "U", ""))
        LbldetailSRV.Caption = LbldetailSRV.Caption & IIf(p_MultiSite, "   (Sites : " & strSites & ")", "")
        LbldetailSRV.Visible = IIf(Trim(LbldetailSRV.Caption) <> "", True, False)
    End If
End Sub

Private Sub tv_Expand_OLD(ByVal Node As ComctlLib.Node)

    g_button = -1
'    g_expand = True
    
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call afficher_menu(True)
    End If
    
End Sub

Private Sub tv_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        If g_plusieurs Then
            Call basculer_selection
        End If
    End If
    
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    g_button = Button
    g_node = tv.SelectedItem.Index
    g_expand = tv.SelectedItem.Expanded
    
End Sub

Private Sub txtRecherche_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If TxtRecherche.Text <> "" Then
            Call afficher_liste3(TxtRecherche.Text)
        Else
            If g_mode_acces <> MODE_PARAM_PERS Then
                If afficher_liste() = P_ERREUR Then
                    Call quitter
                    Exit Sub
                End If
            Else
                If afficher_liste2() = P_ERREUR Then
                    Call quitter
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub
