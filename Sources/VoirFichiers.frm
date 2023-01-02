VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VoirFichiers 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Résultats"
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
      TabIndex        =   3
      Top             =   0
      Width           =   10905
      Begin ComctlLib.TreeView tv 
         Height          =   6285
         Left            =   195
         TabIndex        =   4
         Top             =   915
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   11086
         _Version        =   327682
         Indentation     =   2
         LabelEdit       =   1
         LineStyle       =   1
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
      Begin VB.Label lblDiff 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   10575
      End
      Begin ComctlLib.ImageList img 
         Left            =   6480
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VoirFichiers.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VoirFichiers.frx":047A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VoirFichiers.frx":0840
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   800
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   10905
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
         Left            =   9720
         Picture         =   "VoirFichiers.frx":0E8A
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   600
         Picture         =   "VoirFichiers.frx":1443
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sélectionner"
         Top             =   200
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   550
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAfficheExcel 
         Caption         =   "Ouvrir le document &Excel"
      End
      Begin VB.Menu mnuDeposeExcel 
         Caption         =   "Re-déposer le document Excel"
      End
      Begin VB.Menu mnuAfficheHTML 
         Caption         =   "Ouvrir le document &HTML"
      End
      Begin VB.Menu mnuDiffuser 
         Caption         =   "&Diffuser"
      End
      Begin VB.Menu mnuRenommer 
         Caption         =   "&Renommer"
      End
      Begin VB.Menu mnuSupprimer 
         Caption         =   "&Supprimer"
      End
   End
End
Attribute VB_Name = "VoirFichiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1

Private Const IMG_EXCEL = 2
Private Const IMG_HTML = 3
Private Const IMG_DOSSIER = 1

Private g_numModele As Long
Private g_mode As String
Private g_rp As String
Private g_Ext_XLS As String

Private g_form_active As Boolean
Private g_button As Integer

Public Sub AppelFrm(ByVal v_nummodele As Long, ByVal v_mode As String, ByRef v_rp As String)

    g_numModele = v_nummodele
    g_mode = v_mode
    
    Me.Show 1
    If v_mode = "SEL" Then
        v_rp = g_rp
    End If
End Sub

Private Sub afficher_liste()
    Dim sext As String
    Dim sql As String
    Dim num_modele As Long, numdoc As Long, numfich As Long, img As Long
    Dim nd As Node, ndd As Node, ndr As Node
    Dim rsr As rdoResultset, rsd As rdoResultset, rsf As rdoResultset
    Dim LeXlsExiste As Boolean, LeTxtExiste As Boolean
    Dim nomfich_serv As String
    Dim laS As String
    
    tv.Nodes.Clear
    
    ' Les rapports
    sql = "select * from rapport_type"
    If p_CodeUtil = "ROOT" Then
        sql = sql & " where true"
    Else
        sql = sql & " where rp_user_admin like '%U" & p_NumUtil & "=%'"
    End If
    If g_numModele > 0 Then
        sql = sql & " and rp_num=" & g_numModele
    End If
    If Odbc_SelectV(sql, rsr) = P_ERREUR Then
        Exit Sub
    End If
    While Not rsr.EOF
        num_modele = rsr("rp_num").Value
        'Debug.Print rsr("rp_titre_modele") & " " & rsr("rp_user_admin")
        Set ndr = tv.Nodes.Add(, tvwChild, "R" & num_modele, rsr("rp_titre_modele").Value, IMG_DOSSIER)
        ndr.tag = "RP_" & num_modele
        If g_numModele > 0 Then
            ndr.Expanded = True
        End If
        sql = "select * from rp_document" _
              & " where rpd_rpnum=" & num_modele
        If Odbc_SelectV(sql, rsd) = P_ERREUR Then
            Exit Sub
        End If
        ' Les documents du rapport
        If rsd.EOF Then
            tv.Nodes.Remove (ndr.Index)
            GoTo nextRSR
        End If
        While Not rsd.EOF
            numdoc = rsd("rpd_num").Value
            'Debug.Print sql
            'Debug.Print "   rpd_num=" & rsd("rpd_num").Value & " document " & rsd("rpd_titre").Value
            If TV_NodeExiste(tv, num_modele & "_" & numdoc, ndd) = P_NON Then
                Set ndd = tv.Nodes.Add(ndr, tvwChild, "D" & numdoc, rsd("rpd_titre").Value, IMG_DOSSIER)
                ndd.tag = "RP_" & num_modele & "_" & numdoc
            End If
            sql = "select * from rp_fichier" _
                  & " where rpf_rpdnum=" & numdoc
            'Debug.Print sql
            If Odbc_SelectV(sql, rsf) = P_ERREUR Then
                Exit Sub
            End If
            ' Les fichiers du document
            If rsf.EOF Then
                tv.Nodes.Remove (ndd.Index)
                GoTo nextRSD
            End If
            While Not rsf.EOF
                Debug.Print "      fichier " & rsf(0).Value & " " & rsf(1).Value & " " & rsf(2).Value & " " & rsf(3).Value & " " & rsf(4).Value; ""
                If rsf("rpf_diff_faite").Value Then
                    img = IMG_HTML
                Else
                    img = IMG_EXCEL
                End If
                ' voir s'ils existent
                LeXlsExiste = False
                LeTxtExiste = True
                numfich = rsf("rpf_num").Value
                numdoc = rsf(2).Value
                nomfich_serv = p_Chemin_Résultats & "/RP_" & rsf("rpf_rpnum") & "/Doc_" & numdoc & "/" & numfich '  & ".xls"
                Debug.Print nomfich_serv
                sext = Positionne_Extension(nomfich_serv)
                Set nd = tv.Nodes.Add(ndd, tvwChild, "F" & numfich & "_D" & numdoc, rsf("rpf_titre").Value & IIf(sext = "", " (!!!)", ""), img)
                nd.tag = "RP_" & num_modele & "_" & numdoc & "_" & numfich
                If Not rsf("rpf_diff_faite").Value Then
                    ndd.Expanded = True
                End If
                rsf.MoveNext
            Wend
            rsf.Close
nextRSD:
            rsd.MoveNext
        Wend
        rsd.Close
nextRSR:
        rsr.MoveNext
    Wend
    rsr.Close
    tv.SetFocus
    
End Sub

Private Sub afficher_menu()

    Dim diff_faite As Boolean
    Dim sql As String
    Dim numfich As Long
    Dim numdoc As Long
    
    If tv.SelectedItem.image = IMG_EXCEL Or tv.SelectedItem.image = IMG_HTML Then
        mnuAfficheExcel.Visible = True
        mnuDeposeExcel.Visible = False
        If mnuDeposeExcel.tag <> "" Then
            mnuDeposeExcel.Visible = True
        End If
        mnuAfficheHTML.Visible = True
        numfich = Mid$(STR_GetChamp(tv.SelectedItem.key, "_", 0), 2)
        numdoc = Mid$(STR_GetChamp(tv.SelectedItem.key, "_", 1), 2)
        sql = "select rpf_diff_faite from rp_fichier" _
            & " where rpf_num=" & numfich & " and rpf_rpdnum=" & numdoc
        Call Odbc_RecupVal(sql, diff_faite)
        If diff_faite Then
            mnuDiffuser.Visible = False
        Else
            mnuDiffuser.Visible = True
        End If
        mnuSupprimer.Visible = True
        mnuRenommer.Visible = True
    Else
        mnuAfficheExcel.Visible = False
        mnuDeposeExcel.Visible = False
        mnuAfficheHTML.Visible = False
        mnuDiffuser.Visible = False
        mnuSupprimer.Visible = True
        If STR_GetNbchamp(tv.SelectedItem.tag, "_") = 3 Then
            mnuRenommer.Visible = True
        Else
            mnuRenommer.Visible = False
        End If
    End If
    PopupMenu mnuMenu

End Sub

Private Function build_arbor_serv(ByVal v_ssp As String) As String

    Dim sql As String, s As String
    Dim numsrv As Long
    
    If left$(v_ssp, 1) = "S" Then
        numsrv = Mid$(v_ssp, 2)
    Else
        sql = "select po_srvnum from poste where po_num = " & Mid$(v_ssp, 2)
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            build_arbor_serv = v_ssp
            Exit Function
        End If
    End If
    
    s = v_ssp & ";"
    Do
        sql = "select srv_numpere from service where srv_num = " & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            build_arbor_serv = v_ssp
            Exit Function
        End If
        If numsrv > 0 Then
            s = "S" & numsrv & ";" & s
        End If
    Loop Until numsrv = 0
    
    build_arbor_serv = s
    
End Function

Private Function CreerDoc(ByVal v_numdoc As Long, ByVal d_titre As String) As Long

    Dim fok As Boolean, docpublic As Boolean
    Dim sql As String, modele As Variant, refdoc As String, titre_doc As String
    Dim s As String, chemin_doc_serv As String, nomfich_doc_serv As String
    Dim nomfich_resu_serv  As String, nomfich_loc_xls As String, liberr As String
    Dim chemin_loc_html  As String, nomfich_loc_html As String
    Dim chemin_docpub_serv As String, s_sp As String, s2 As String
    Dim n As Integer, I As Integer
    Dim numDos As Long, numnat As Long, lnb As Long, numdocs As Long
    Dim ordre As Long, numdoc As Long, lbid As Long
    Dim lstdest_rapp As Variant, lstresp_dos As Variant, lstresp As Variant
    Dim lstdest_doc As Variant
    Dim champModele As String
    
    sql = "select rpd_titre, rpd_dsnum, rpd_ndnum, rpd_modele, rpd_public, rpd_lstdest from rp_document" _
        & " where rpd_num=" & v_numdoc
    Call Odbc_RecupVal(sql, titre_doc, numDos, numnat, modele, docpublic, lstdest_rapp)
    
    ' Reformatage de lstdest
    lstdest_doc = ""
    n = STR_GetNbchamp(lstdest_rapp, ";")
    For I = 0 To n - 1
        s2 = STR_GetChamp(lstdest_rapp, ";", I)
        If left$(s2, 1) <> "P" Then
            If left$(s2, 1) = "S" Then
                lstdest_doc = lstdest_doc & s2 & ";|"
            Else
                lstdest_doc = lstdest_doc & s2 & "|"
            End If
        Else
            s_sp = build_arbor_serv(s2)
            If left$(s2, 1) = "S" Then
                lstdest_doc = lstdest_doc & ";"
            End If
            lstdest_doc = lstdest_doc & s_sp & "|"
        End If
    Next I
    
    refdoc = RecupDocReference(numDos, numnat)
    
    fok = False
    While Not fok
        fok = True
        If refdoc = "" Then
            refdoc = InputBox("Référence du document" & vbCrLf & "(20 caractères maximum)", "Envoi vers KaliDoc", refdoc)
        End If
        If Len(refdoc) > 20 Then
            MsgBox "Référence trop longue (" & Len(refdoc) & "car. pour 20 car. maximum)"
            refdoc = ""
            fok = False
        Else
            If refdoc = "" Then
                CreerDoc = P_ERREUR
                Exit Function
            End If
            sql = "select count(*) from document where D_Ident = '" & refdoc & "'"
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                CreerDoc = P_ERREUR
                Exit Function
            End If
            If lnb > 0 Then
                MsgBox "Cette référence existe déjà"
                refdoc = ""
                fok = False
            End If
        End If
    Wend
    
    ' Document
    sql = "select ds_donum, ds_LstResp from Dossier where ds_num=" & numDos
    If Odbc_RecupVal(sql, numdocs, lstresp_dos) = P_ERREUR Then
        CreerDoc = P_ERREUR
        Exit Function
    End If
    
    lstresp = "U" & p_NumUtil & ";0;0;1;|"
    n = STR_GetNbchamp(lstresp_dos, "|")
    For I = 0 To n - 1
        s = STR_GetChamp(lstresp_dos, "|", I)
        If InStr(s, "U" & p_NumUtil & ";") = 0 Then
            lstresp = lstresp & STR_GetChamp(s, ";", 0) & ";" & STR_GetChamp(s, ";", 4) & ";" & STR_GetChamp(s, ";", 5) & ";" & STR_GetChamp(s, ";", 2) & ";|"
        End If
    Next I
    
    ' Détermine l'ordre du document
    sql = "select max(D_Ordre) from Document" _
        & " where D_DSNum=" & numDos
    If Odbc_MinMax(sql, ordre) = P_ERREUR Then
        CreerDoc = P_ERREUR
        Exit Function
    End If
    ordre = ordre + 1

    If Odbc_BeginTrans() = P_ERREUR Then
        Exit Function
    End If
    
    ' Gestion du modele => d_modcnum ou d_modele (selon P_MODCNUM_ou_MODELE = donm_modcnum ou donm_modele)
    If P_MODCNUM_ou_MODELE = "donm_modcnum" Then
        champModele = "d_modcnum"
    Else
        champModele = "d_modele"
    End If
    ' Le document aura le cycle d'ordre 0
    If Odbc_AddNew("Document", "D_Num", "d_seq", True, numdoc, _
                    "D_DONum", numdocs, "D_DSNum", numDos, _
                    "D_Ordre", ordre, "D_NDNum", numnat, _
                    "D_NumVers", 1, "D_LibVers", "1", _
                    "D_DateVers", Date, "D_CYOrdre", 0, _
                    "D_Ident", refdoc, "D_Titre", titre_doc & " (" & d_titre & ")", _
                    "D_Descr", "", "D_Referentiel", "", _
                    "D_Theme", "", "d_datecreation", Format(Date, "YYYY-MM-DD"), _
                    "D_UNumResp", p_NumUtil, "D_LstResp", lstresp, _
                    champModele, modele, "D_NomFichier", "", _
                    "D_Public", docpublic, "D_Convertir", True, _
                    "D_AjouterPJ", True, "D_Telecharger", True, _
                    "D_VisuIntranet", True, "D_GererEval", False, _
                    "D_Site", "L1;", "D_Dest", lstdest_doc, _
                    "D_CycleAnnul", 0, "D_CycleRevision", "", _
                    "D_MotPasse", "", "D_DiffListe", True, _
                    "D_DiffExTempo", False, "D_SRVNum_emet", 0, _
                    "D_GarderStyle", True, "D_ArchDuree", "", _
                    "D_ArchLieu", "", "D_ArchResp", "") = P_ERREUR Then
        GoTo err_enreg
    End If
        
    sql = "select PG_CheminDoc, PG_Chemindocpubliweb from PrmGen_HTTP"
    If Odbc_RecupVal(sql, chemin_doc_serv, chemin_docpub_serv) = P_ERREUR Then
        GoTo err_enreg
    End If
    nomfich_doc_serv = chemin_doc_serv & "/" & numdoc & p_PointExtensionXls
    If Odbc_Update("Document", _
                    "D_Num", _
                    "where D_Num=" & numdoc, _
                    "D_NomFichier", nomfich_doc_serv) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    lblDiff.Caption = "Copier le fichier vers le serveur"
    ' copier le xls
    s = "RP_" & STR_GetChamp(tv.SelectedItem.tag, "_", 1) & "/Doc_" & STR_GetChamp(tv.SelectedItem.tag, "_", 2) & "/" & STR_GetChamp(tv.SelectedItem.tag, "_", 3)
    nomfich_resu_serv = p_Chemin_Résultats & "/" & s & p_PointExtensionXls
    If KF_CopierFichier(nomfich_resu_serv, nomfich_doc_serv) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    ' Refaire le html
    lblDiff.Caption = "Transformer en HTML"
    nomfich_loc_xls = p_chemin_appli & "\tmp\RP_" & Format(Time, "hhmmss") & p_PointExtensionXls
    If KF_GetFichier(nomfich_resu_serv, nomfich_loc_xls) = P_ERREUR Then
        GoTo lab_diffusion
    End If
    
    Call Public_VerifOuvrir(nomfich_loc_xls, False, False, p_tbl_FichExcelPublier)

    ' transformer en HTML
    chemin_loc_html = p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil
    If Not FICH_EstRepertoire(chemin_loc_html, False) Then
        Call FICH_CreerRepComp(chemin_loc_html, False, False)
    End If
    nomfich_loc_html = chemin_loc_html & "\" & numdoc & "-1.html"
    Exc_wrk.SaveAs FileName:=nomfich_loc_html, _
        FileFormat:=44, ReadOnlyRecommended:=False, CreateBackup:=False
    Call Exc_wrk.Close
    ' Transfère le .html sur le serveur
    lblDiff.Caption = "Transférer le HTML vers le serveur"
    If HTTP_Appel_PutDos(p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil, _
                         chemin_docpub_serv & "/", False, False, liberr) <> HTTP_OK Then
        MsgBox liberr
    End If
    ' Vider le dossier local de transfert
    Call FICH_EffacerRep(chemin_loc_html)
    Call FICH_EffacerFichier(nomfich_loc_html, False)
    Call FICH_EffacerFichier(nomfich_loc_xls, False)
        
lab_diffusion:
    lblDiff.Caption = "Diffusion aux destinataires"
    If ajouter_tbl_prmdiffusion(numdoc, True, lstdest_rapp) = P_ERREUR Then
        GoTo err_enreg
    End If
        
    ' Ajout dans DocUtil : responsables, acteurs, destinataires
    If ajouter_tbl_docutil(numdoc, numDos, numdocs, numnat) = P_ERREUR Then
        GoTo err_enreg
    End If
        
    If Odbc_AddNew("DocVersion", "DV_Num", "dev_seq", False, lbid, _
                    "DV_DNum", numdoc, _
                    "DV_NumVers", 1, _
                    "DV_NumVersW", 1, _
                    "DV_LibVers", "1", _
                    "DV_Ext", p_PointExtensionXls, _
                    "DV_CYOrdre", 0, _
                    "DV_DocExisteQual", False, _
                    "DV_DateVers", Format(Date, "YYYY-MM-DD"), _
                    "DV_MotifCourt", "Création", _
                    "DV_MotifLong", "Création par KaliRP", _
                    "DV_TypConvHTML", 1, _
                    "DV_SizeScr", 1) = P_ERREUR Then
        GoTo err_enreg
    End If
        
    If Odbc_AddNew("DocEtapeVersion", "DEV_Num", "dev_seq", False, lbid, _
                    "DEV_DNum", numdoc, _
                    "DEV_NumVers", 1, _
                    "DEV_CYOrdre", 0, _
                    "DEV_UNum", p_NumUtil, _
                    "DEV_UNumFait", p_NumUtil, _
                    "DEV_PONum", 0, _
                    "DEV_Date", Format(Date, "YYYY-MM-DD"), _
                    "DEV_Refus", False, _
                    "DEV_Commentaire", "Création par KaliRP", _
                    "DEV_Intitule", "", _
                    "DEV_IntituleRemplace", False) = P_ERREUR Then
        GoTo err_enreg
    End If
        
    Call Odbc_CommitTrans
    CreerDoc = P_OK
    Exit Function
        
err_enreg:
    Call Odbc_RollbackTrans
    CreerDoc = P_ERREUR
    
End Function

Private Function ajouter_tbl_docutil(ByVal v_numdoc As Long, _
                                     ByVal v_numdos As Long, _
                                     ByVal v_numdocs As Long, _
                                     ByVal v_numnat As Long) As Integer

    Dim sannul As String, sql As String, sannul_doc As String
    Dim ya_acteur As Boolean, ya_un_acteur As Boolean
    Dim prem_cycle As Integer, last_cycle As Integer, iadd As Integer, nbenf As Integer, siztbl As Integer
    Dim prem_numutil As Long, numposte As Long, lbid As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    
    sql = "select don_cycleannul from docsnature" _
            & " where don_ndnum=" & v_numnat _
            & " and don_donum=" & v_numdocs
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_tbl_docutil = P_ERREUR
        Exit Function
    Else
        If rs.EOF Then
            MsgBox "Aucune nature de document n'est attachée au dossier"
            ajouter_tbl_docutil = P_ERREUR
            Exit Function
        End If
    End If
    
    sannul_doc = ""
    ya_un_acteur = False
    prem_cycle = 0
    last_cycle = 0
    sql = "select * from cycle where cy_ordre>0" _
        & " and cy_acteur <> ''"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_tbl_docutil = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        If InStr(sannul, "C" & rs("cy_ordre").Value & ";") = 0 Then
            ya_acteur = False
            sql = "select dsu_unum, dsu_niveau, dsu_unumr from dosutil where dsu_dsnum=" & v_numdos
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                ajouter_tbl_docutil = P_ERREUR
                Exit Function
            End If
            While Not rs2.EOF
                ya_acteur = True
                ya_un_acteur = True
                If prem_cycle = 0 Then
                    prem_cycle = rs("cy_ordre").Value
                    prem_numutil = rs2("dsu_unum").Value
                End If
                numposte = recup_poste(rs2("dsu_unum").Value)
                If numposte = 0 Then
                    ajouter_tbl_docutil = P_ERREUR
                    Exit Function
                End If
                If AjouterDocUtil_Act(v_numdoc, _
                                        rs2("dsu_unum").Value, _
                                        rs("cy_ordre").Value, _
                                        rs2("dsu_niveau").Value, _
                                        rs2("dsu_unumr").Value, _
                                        numposte) = P_ERREUR Then
                    ajouter_tbl_docutil = P_ERREUR
                    Exit Function
                End If
                ' Fait exprès pour l'instant ...
                rs2.MoveLast
                rs2.MoveNext
            Wend
            rs2.Close
            If Not ya_acteur Then
                sannul_doc = sannul_doc & "C" & rs("cy_ordre").Value & ";"
                last_cycle = rs("cy_ordre").Value
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    If Not ya_un_acteur Then
        If sannul_doc = "" Or last_cycle = 0 Then
            MsgBox "Problème au niveau des acteurs ..."
            ajouter_tbl_docutil = P_ERREUR
            Exit Function
        Else
            sannul_doc = Replace(sannul_doc, "C" & last_cycle & ";", "")
            numposte = recup_poste(p_NumUtil)
            If numposte = 0 Then
                ajouter_tbl_docutil = P_ERREUR
                Exit Function
            End If
            If AjouterDocUtil_Act(v_numdoc, _
                                    p_NumUtil, _
                                    last_cycle, _
                                    1, _
                                    0, _
                                    numposte) = P_ERREUR Then
                ajouter_tbl_docutil = P_ERREUR
                Exit Function
            End If
            prem_cycle = last_cycle
            prem_numutil = p_NumUtil
        End If
    End If
    If sannul_doc <> "" Then
        Call Odbc_Update("Document", "D_Num", "where D_num=" & v_numdoc, _
                         "D_CycleAnnul", sannul_doc)
    End If
    
    If Odbc_AddNew("DocAction", _
                    "DAC_Num", _
                    "dac_seq", _
                    False, _
                    lbid, _
                    "DAC_DNum", v_numdoc, _
                    "DAC_CYOrdre", prem_cycle, _
                    "DAC_Niveau", 1, _
                    "DAC_UNum", prem_numutil, _
                    "DAC_PONum", numposte, _
                    "DAC_Date", Date, _
                    "DAC_DatePrevue", Null, _
                    "DAC_CYOrdreRefus", 0, _
                    "DAC_UNumRefus", 0, _
                    "DAC_Commentaire", "", _
                    "DAC_UNumModif", 0, _
                    "DAC_ActionVu", False, _
                    "DAC_YaModif", False, _
                    "DAC_Relecture", False, _
                    "DAC_Intitule", "", _
                    "DAC_IntituleRemplace", False) = P_ERREUR Then
        ajouter_tbl_docutil = P_ERREUR
        Exit Function
    End If
    
    p_NumDocs = v_numdocs
    If P_MajUtilADIM(prem_numutil, "A", 1) = P_ERREUR Then
        ajouter_tbl_docutil = P_ERREUR
        Exit Function
    End If

    ajouter_tbl_docutil = P_OK
    
End Function

Private Function RecupDocReference(ByVal v_numdos As Long, _
                                    ByVal v_numNature As String) As String

    Dim smsq As String, sident As String, sref As String
    Dim sql As String, sref_rech As String, smax As String, sC As String
    Dim schp As String, sval As String, s As String, ssrv As String
    Dim sref1 As String, smsq_init As String, sval1 As String
    Dim ignorer As Boolean, rechercher As Boolean
    Dim I As Integer, pos1 As Integer, pos2 As Integer, pos As Integer, lg As Integer
    Dim lg_smsq As Integer, nb As Integer
    Dim max As Long, lnb As Long
    Dim rs As rdoResultset
    
    If Odbc_RecupVal("select DS_Ident, DS_Masque from Dossier where DS_Num=" & v_numdos, _
                     sident, _
                     smsq_init) = P_ERREUR Then
        RecupDocReference = ""
        Exit Function
    End If
    'smsq_init = "RP_<NAT>_<NNNN>"
    ' Pas de format
    If smsq_init = "" Then
        RecupDocReference = ""
        Exit Function
    End If
    smsq_init = STR_Remplacer(smsq_init, "<R>", sident)
    
    ' Code Nature
    sql = "select * from NatureDoc where ND_num = " & v_numNature
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        RecupDocReference = ""
        Exit Function
    End If
    If rs.EOF Then
        RecupDocReference = ""
        Exit Function
    End If
        
    smsq_init = STR_Remplacer(smsq_init, "<NAT>", rs("ND_Code"))
    
    sref_rech = Odbc_StringJoker(smsq_init)
    sref_rech = Replace(sref_rech, "<NNNNN>", "_____")
    sref_rech = Replace(sref_rech, "<NNNN>", "____")
    sref_rech = Replace(sref_rech, "<NNN>", "___")
    sref_rech = Replace(sref_rech, "<NN>", "__")
    sref_rech = Replace(sref_rech, "<N>", "_")
    If sref_rech = smsq_init Then
        RecupDocReference = smsq_init
        Exit Function
    End If
    ' Zone manuelle
    s = ""
    ignorer = False
    For I = 1 To Len(sref_rech)
        If Mid$(sref_rech, I, 1) = "<" Then
            ignorer = True
            s = s & "%"
        ElseIf Mid$(sref_rech, I, 1) = ">" Then
            ignorer = False
        ElseIf Not ignorer Then
            s = s & Mid$(sref_rech, I, 1)
        End If
    Next I
    sref_rech = s
lab_recherche:
    smsq = smsq_init
    If InStr(sref_rech, "?") > 0 Then
        sref_rech = Replace(sref_rech, "?", "_")
        sql = "select D_Ident from Document" _
            & " where D_Ident like " & sref_rech
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            RecupDocReference = ""
            Exit Function
        End If
        If Not rs.EOF Then
            pos = InStr(smsq, "<N")
            If pos = 0 Then
                rs.Close
                RecupDocReference = ""
                Exit Function
            End If
            pos2 = InStr(pos + 1, smsq, ">")
            If pos2 = 0 Then
                rs.Close
                RecupDocReference = ""
                Exit Function
            End If
            lg = pos2 - pos - 1
            max = 0
            smax = ""
            While Not rs.EOF
                s = Mid$(rs("D_Ident").Value, pos, lg)
                If IsNumeric(s) Then
                    If CLng(s) > max Then
                        max = CLng(s)
                        smax = rs("D_Ident").Value
                    End If
                End If
                rs.MoveNext
            Wend
        End If
        rs.Close
    Else
        sql = "select max(D_Ident) from Document" _
            & " where D_Ident like " & sref_rech
        If Odbc_Select(sql, rs) = P_ERREUR Then
            RecupDocReference = ""
            Exit Function
        End If
        smax = rs(0).Value & ""
        rs.Close
    End If
    
    sref = smsq
    pos = 0
    pos1 = InStr(pos + 1, smsq, "<N")
    If pos1 > 0 Then
        pos2 = InStr(pos1 + 1, smsq, ">")
        nb = 0
        For I = 1 To pos1 - 1
            If Mid$(smsq, I, 1) = "<" Then
                nb = nb + 1
            ElseIf Mid$(smsq, I, 1) = ">" Then
                nb = nb + 1
            End If
        Next I
        pos1 = pos1 - nb
        nb = 0
        For I = 1 To pos2 - 1
            If Mid$(smsq, I, 1) = "<" Then
                nb = nb + 1
            ElseIf Mid$(smsq, I, 1) = ">" Then
                nb = nb + 1
            End If
        Next I
        pos2 = pos2 - nb
        smsq = Replace(smsq, "<", "")
        smsq = Replace(smsq, ">", "")
        If pos2 > 0 Then
            schp = Mid$(smsq, pos1, pos2 - pos1)
            Select Case schp
            Case "N", "NN", "NNN", "NNNN", "NNNNN"
                sval = Mid$(smax, pos1, Len(schp))
                If STR_EstEntierPos(sval) Then
                    sval1 = sval
                    sval = STR_Incrementer(sval)
                    If Len(sval) < Len(schp) Then
                        sval = String$(Len(schp) - Len(sval), "0") + sval
                    ElseIf Len(sval) > Len(schp) Then
                        If CInt(sval) = CInt(sval1) + 1 Then
                            sref = Replace(sref, "<" & schp & ">", sval)
                            Call Odbc_Count("select count(*) from document where d_ident=" & Odbc_String(sref), lnb)
                            If lnb > 0 Then
                                rechercher = True
                            Else
                                rechercher = False
                            End If
                        Else
                            rechercher = True
                        End If
                        If rechercher Then
                            sref_rech = Replace(sref_rech, String$(Len(schp), "_"), String$(Len(schp) + 1, "_"))
                            smsq_init = Replace(smsq_init, "<" & String$(Len(schp), "N"), "<" & String$(Len(schp) + 1, "N"))
                            GoTo lab_recherche
                        End If
                    End If
                    sref = Replace(sref, "<" & schp & ">", sval)
                Else
                    sval = String$(Len(schp) - 1, "0") + "1"
                    sref1 = STR_Remplacer(sref, "<" & schp & ">", sval)
                    Do
                        Call Odbc_Count("select count(*) from document where d_ident=" & Odbc_String(sref1), lnb)
                        If lnb > 0 Then
                            sval = STR_Incrementer(sval)
                            sref1 = STR_Remplacer(sref, "<" & schp & ">", sval)
                        End If
                    Loop Until lnb = 0
                    sref = sref1
                End If
            End Select
        End If
    End If

    RecupDocReference = sref

End Function

Private Function recup_poste(ByVal v_numutil As Long) As Long

    Dim sql As String
    Dim numposte As Long
    
    sql = "select u_po_princ from utilisateur where u_num=" & v_numutil
    If Odbc_RecupVal(sql, numposte) = P_ERREUR Then
        numposte = 0
    End If
    recup_poste = numposte
    
End Function

Private Function AjouterDocUtil_Act(ByVal v_numd As Long, _
                                    ByVal v_numutil As Long, _
                                    ByVal v_cyordre As Long, _
                                    ByVal v_niveau As Long, _
                                    ByVal v_numutilr As Long, _
                                    ByVal v_numposte As Long) As Integer

    Dim lbid As Long
    
    If Odbc_AddNew("DocUtil", _
                    "DU_Num", _
                    "du_seq", _
                    False, _
                    lbid, _
                    "DU_DNum", v_numd, _
                    "DU_UNum", v_numutil, _
                    "DU_CYOrdre", v_cyordre, _
                    "DU_Niveau", v_niveau, _
                    "DU_UNumR", v_numutilr, _
                    "DU_PONum", v_numposte, _
                    "DU_Intitule", "", _
                    "DU_IntituleRemplace", False) = P_ERREUR Then
        AjouterDocUtil_Act = P_ERREUR
        Exit Function
    End If
    
    AjouterDocUtil_Act = P_OK
    
End Function

Private Function ajouter_tbl_prmdiffusion(ByVal v_numdoc As Long, _
                                          ByVal v_modecreation As Boolean, _
                                          ByVal v_sdest As Variant) As Integer

    Dim sql As String, s As String, slstgrp As String
    Dim sG As String, sg1 As String
    Dim I As Integer, n As Integer, j As Integer, ng As Integer
    Dim sdest As Variant
    Dim rs As rdoResultset
    
    n = STR_GetNbchamp(v_sdest, ";")
    sdest = ""
    For I = 0 To n - 1
        s = STR_GetChamp(v_sdest, ";", I)
        Select Case left$(s, 1)
        Case "G"
            If Odbc_Select("select GU_Lst from GroupeUtil where GU_Num=" & Mid$(s, 2), _
                             rs) = P_OK Then
                sG = rs("GU_Lst").Value & ""
                ng = STR_GetNbchamp(sG, "|")
                For j = 0 To ng - 1
                    sg1 = STR_GetNbchamp(sG, "|")
                    sdest = sdest + sg1 + ";"
                Next j
                rs.Close
            End If
        Case "F", "S", "U"
            sdest = sdest & s & ";"
        End Select
    Next I
    If sdest = "" Then
        MsgBox "Document créé mais aucun destinataire n'est spécifié"
    Else
        n = STR_GetNbchamp(sdest, ";")
        For I = 0 To n - 1
            s = STR_GetChamp(sdest, ";", I)
            Select Case left$(s, 1)
            Case "F"
                sql = "select u_num from utilisateur where U_FctTrav like '%" & s & ";%'"
            Case "S"
                sql = "select u_num from utilisateur where U_SPM like '%" & s & "%'"
            Case "U"
                sql = "select u_num from utilisateur where U_Num=" & Mid$(s, 2)
            End Select
        Next I
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            ajouter_tbl_prmdiffusion = P_ERREUR
            Exit Function
        End If
        
        If v_modecreation Then
            While Not rs.EOF
                sql = "insert into docprmdiffusion" _
                    & " (dpd_dnum, dpd_unum, dpd_estresp, dpd_typediff, dpd_exemplaire)" _
                    & " values" _
                    & " (" & v_numdoc & "," & rs("U_num").Value _
                    & ", false, 2" _
                    & ", '')"
                On Error GoTo err_execute
                Call Odbc_Cnx.Execute(sql)
                On Error GoTo 0
                rs.MoveNext
            Wend
        Else
            MsgBox "Modification de la diffusion : à faire"
        End If
        rs.Close
    End If
    
    ajouter_tbl_prmdiffusion = P_OK
    Exit Function
    
err_execute:
    MsgBox "Erreur Execute pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    ajouter_tbl_prmdiffusion = P_ERREUR
    
End Function

Private Sub Diffuser()
    
    Dim trouve As Boolean, diffuser_kd As Boolean
    Dim sql As String, s_undest As String, titre_rp As String
    Dim titre As String
    Dim nbu As Integer, idest As Integer, I As Integer
    Dim numdoc As Long, numutil As Long
    Dim lstdest As Variant
    Dim tbl_util() As Long, numfich As Long
    Dim mess As Variant
    Dim rs As rdoResultset
    
    lblDiff.Caption = "Diffusion en cours ..."
    
    numdoc = Mid$(tv.SelectedItem.Parent.key, 2)
    sql = "select rpd_publier_kd, rpd_lstdest from rp_document" _
        & " where rpd_num=" & numdoc
    Call Odbc_RecupVal(sql, diffuser_kd, lstdest)
    
    titre = tv.SelectedItem.Text
    If diffuser_kd Then
        If CreerDoc(numdoc, titre) = P_ERREUR Then
            lblDiff.Caption = ""
            Call MsgBox("Erreur lors de la diffusion")
            Exit Sub
        End If
    End If
    
    ' Envoi mail
    If Not diffuser_kd Then
        lblDiff.Caption = "Diffusion aux destinataires"
        nbu = -1
        For idest = 0 To STR_GetNbchamp(lstdest, ";")
            s_undest = STR_GetChamp(lstdest, ";", idest)
            If left$(s_undest, 1) = "G" Then
                sql = "select GU_Lst from GroupeUtil where GU_Num=" & Mid$(s_undest, 2)
                If Odbc_SelectV(sql, rs) = P_OK Then
                    If Not rs.EOF Then
                        'lstdest = lstdest & rs("gu_lst").Value
                    End If
                    rs.Close
                End If
            ElseIf left$(s_undest, 1) = "U" Then
                numutil = Mid$(s_undest, 2)
                sql = "select U_Num from Utilisateur where U_Num=" & numutil _
                    & " and u_actif=true"
                If Odbc_SelectV(sql, rs) = P_OK Then
                    If Not rs.EOF Then
                        trouve = False
                        For I = 0 To nbu
                            If tbl_util(I) = rs("u_num").Value Then
                                trouve = True
                                Exit For
                            End If
                        Next I
                        If Not trouve Then
                            nbu = nbu + 1
                            ReDim Preserve tbl_util(nbu) As Long
                            tbl_util(nbu) = rs("u_num").Value
                        End If
                    End If
                    rs.Close
                End If
            ElseIf left$(s_undest, 1) = "S" Then
                sql = "select U_Num from Utilisateur where U_spm like '%" & s_undest & ";%'" _
                    & " and u_actif=true"
                If Odbc_SelectV(sql, rs) = P_OK Then
                    While Not rs.EOF
                        trouve = False
                        For I = 0 To nbu
                            If tbl_util(I) = rs("u_num").Value Then
                                trouve = True
                                Exit For
                            End If
                        Next I
                        If Not trouve Then
                            nbu = nbu + 1
                            ReDim Preserve tbl_util(nbu) As Long
                            tbl_util(nbu) = rs("u_num").Value
                        End If
                        rs.MoveNext
                    Wend
                    rs.Close
                End If
            ElseIf left$(s_undest, 1) = "F" Then
                sql = "select U_Num from Utilisateur where U_fcttrav like '%" & s_undest & ";%'" _
                    & " and u_actif=true"
                If Odbc_SelectV(sql, rs) = P_OK Then
                    While Not rs.EOF
                        trouve = False
                        For I = 0 To nbu
                            If tbl_util(I) = rs("u_num").Value Then
                                trouve = True
                                Exit For
                            End If
                        Next I
                        If Not trouve Then
                            nbu = nbu + 1
                            ReDim Preserve tbl_util(nbu) As Long
                            tbl_util(nbu) = rs("u_num").Value
                        End If
                        rs.MoveNext
                    Wend
                    rs.Close
                End If
            End If
        Next idest
        
        ' Envoie un mail à chaque destinataire
        If nbu >= 0 Then
            titre_rp = tv.SelectedItem.Parent.Parent.Text
            titre = tv.SelectedItem.Text _
                    & " (" & tv.SelectedItem.Parent.Text & ")"
            Call envoyer_mail_aux_dest(tbl_util(), 1, titre_rp, titre)
        End If
    End If
    
    titre = tv.SelectedItem.Parent.Text & " (" & tv.SelectedItem.Text & ")"
    numfich = Mid$(STR_GetChamp(tv.SelectedItem.key, "_", 0), 2)
    numdoc = Mid$(STR_GetChamp(tv.SelectedItem.key, "_", 1), 2)
    sql = "update rp_fichier set rpf_diff_faite='t'" _
        & " where rpf_num=" & numfich & " and rpf_rpdnum=" & numdoc
    Call Odbc_Cnx.Execute(sql)
    tv.SelectedItem.image = IMG_HTML
    
    lblDiff.Caption = ""
    Call MsgBox("Diffusion de " & titre & " terminée")

End Sub

Private Sub envoyer_mail_aux_dest(ByRef v_tblutil() As Long, _
                                  ByVal v_typediff As Integer, _
                                  ByVal v_titre_rp As String, _
                                  ByVal v_titre As String)

    Dim doit_auth As Boolean, est_externe As Boolean
    Dim chemin_php As String, sutil As String, nomfich_serv As String
    Dim nomfich_loc As String
    Dim laS As String
    Dim chemin_http As String
    Dim util As String, cnd_sversconf As String
    Dim nbu As Integer, iu As Integer
    Dim numutil As Long
    Dim nd As String
    Dim url As String
    Dim mess As Variant, mess_intro As Variant, mess_fin As Variant
    Dim rpnum As Long, docnum As Long, fichnum As Long
    
    Call Odbc_RecupVal("select pg_cheminphp from prmgen_http", chemin_php)
' chemin_php = ""
    
    mess_intro = "Bonjour," & vbCrLf & vbCrLf _
                 & "Un rapport concernant '" & v_titre_rp & "' est diffusé." & vbCrLf & vbCrLf
    mess_fin = "Merci et bonne journée." & vbCrLf & vbCrLf _
               & "---- KaliDoc Gestion documentaire orientée qualité ----" & vbCrLf
    
    nbu = UBound(v_tblutil())
    For iu = 0 To nbu
        numutil = v_tblutil(iu)
        If Odbc_RecupVal("select U_KW_MailAuth, U_Externe from Utilisateur where U_Num=" & numutil, _
                         doit_auth, _
                         est_externe) = P_ERREUR Then
            Exit Sub
        End If
' est_externe = True
        mess = ""
        If chemin_php <> "" And Not est_externe Then
            If Not doit_auth Then
                sutil = "&V_util=" & STR_CrypterNombre(Format(numutil, "#0000000"))
            Else
                sutil = ""
            End If
            util = STR_CrypterNombre(Format(numutil, "#0000000"))
            If p_S_Vers_Conf <> "" Then
                cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
            End If
            If v_typediff = 2 Then
            '    mess = mess & "Vous pouvez accuser réception en cliquant sur " _
            '            & chemin_php & "/pident.php?in=ar&dc=" & v_numdoc & "-" & v_numvers & sutil
            Else
                mess = mess & "Vous pouvez le consulter à partir de votre tableau de bord ou en cliquant sur " & vbCrLf
            '            & chemin_php & "/pident.php?in=accesdoc&V_doc=" & v_numdoc & "-" & v_numvers & sutil
                laS = Replace(p_cheminKW, "publiweb", "")
                laS = Replace(p_Chemin_Résultats, laS, chemin_php & "/")
                laS = Replace(laS, "publiweb/", "")
                nd = tv.SelectedItem.tag
                rpnum = val(STR_GetChamp(nd, "_", 1))
                docnum = val(STR_GetChamp(nd, "_", 2))
                fichnum = val(STR_GetChamp(nd, "_", 3))
                laS = laS & "/RP_" & rpnum & "/Doc_" & docnum & "/" & fichnum & ".html"
                url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & laS
                mess = mess & url
            End If
            mess = mess & vbCrLf & vbCrLf
        End If
        
        mess = mess_intro & mess & mess_fin
        nomfich_serv = p_Chemin_Résultats & "/RP_" & rpnum & "/Doc_" & docnum & "/" & fichnum & g_Ext_XLS    'p_PointExtensionXls
        nomfich_loc = p_chemin_appli & "\tmp\RP" & Format(Time, "hhmmss") & g_Ext_XLS   ' p_PointExtensionXls
        If KF_GetFichier(nomfich_serv, nomfich_loc) = P_OK Then
            If UtilEnvoyerMail(p_NumUtil, _
                                   numutil, _
                                   "Nouveau rapport : " & v_titre_rp, _
                                   mess, _
                                   False, _
                                   nomfich_loc) <> P_OUI Then
                Call UtilEnvoyerKaliMail(p_NumUtil, _
                                       numutil, _
                                       "Nouveau rapport : " & v_titre_rp, _
                                       mess, _
                                       nomfich_loc)
            End If
            Call FICH_EffacerFichier(nomfich_loc, False)
        End If
GoTo lab_suiv

        If est_externe Then
            mess = mess_intro & mess & mess_fin
            nomfich_serv = p_Chemin_Résultats & "/RP_" & rpnum & "/Doc_" & docnum & "/" & fichnum & p_PointExtensionXls
            nomfich_loc = p_chemin_appli & "\tmp\RP" & Format(Time, "hhmmss") & p_PointExtensionXls
            If KF_GetFichier(nomfich_serv, nomfich_loc) = P_OK Then
                Call UtilEnvoyerMail(p_NumUtil, _
                                       numutil, _
                                       "Nouveau rapport : " & v_titre_rp, _
                                       mess, _
                                       nomfich_loc)
            End If
            Call FICH_EffacerFichier(nomfich_loc, False)
        Else
            If v_typediff = 2 Then
                'Call P_UtilEnvoyerMail(p_NumUtil, _
                '                       numutil, _
                '                       "AR KaliDoc : Rapport " & titre_rp, _
                '                       mess, _
                '                       True)
            Else
                'Call P_UtilEnvoyerMessage(v_numdoc, v_numvers, v_libvers, _
                '                          p_NumUtil, v_numutil, _
                '                          "Nouveau document KaliDoc : " & titre_doc, _
                '                          mess, _
                '                          False, _
                '                          "", _
                '                          titre_doc, _
                '                          "Le document '" & titre_doc & "' (Réf.:" & refdoc & " / Version " & v_libvers & ") est diffusé par KaliDoc." & vbCrLf & vbCrLf, _
                '                          True)
            End If
        End If
lab_suiv:
    Next iu
    
End Sub

Private Sub UtilEnvoyerKaliMail(ByVal v_numutile As Long, _
                                ByVal v_numutild As Long, _
                                ByVal v_stitre_m As Variant, _
                                ByVal v_mess_m As Variant, _
                                ByVal v_nomfich As String)

    Dim sext As String, nompj As String, nompj_serv As String
    Dim chemin_pj As String
    Dim numkm As Long, numkmd As Long, lbid As Long
    
    If Odbc_RecupVal("select pg_cheminpj from prmgen_http", _
                     chemin_pj) = P_ERREUR Then
        Exit Sub
    End If
    
    'v_mess_m = convertir_kalimail(v_mess_m)
    If Len(v_stitre_m) > 100 Then
        v_stitre_m = left$(v_stitre_m, 100)
    End If
    If Len(v_mess_m) > 1024 Then
        v_mess_m = left$(v_mess_m, 1024)
    End If
    Call Odbc_BeginTrans
    ' Enregistrer le sujet et le corps du KaliMail
    If Odbc_AddNew("kalimail", "km_num", "km_seq", True, numkm, _
                    "km_sujet", v_stitre_m, _
                    "km_corps", v_mess_m, _
                    "km_typelien", 0, _
                    "km_liblien", "", _
                    "km_urllien", "") = P_ERREUR Then
        Call Odbc_RollbackTrans
        Exit Sub
    End If
    ' Enregistrer le destinataire
    If Odbc_AddNew("kalimaildetail", "kmd_num", "kmd_seq", True, numkmd, _
                   "kmd_kmnum", numkm, _
                   "kmd_expnum", v_numutile, _
                   "kmd_destnum", v_numutild, _
                   "kmd_dateenvoi", Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss"), _
                   "kmd_prioritaire", False, _
                   "kmd_niveau", 1, _
                   "kmd_suppexp", 0, _
                   "kmd_suppdest", 0, _
                   "kmd_numpere", 0, _
                   "kmd_accusereception", False, _
                   "kmd_kmdsnum", 0) = P_ERREUR Then
        Call Odbc_RollbackTrans
        Exit Sub
    End If
    ' Enregistrer les PJ
    sext = Mid$(v_nomfich, InStrRev(v_nomfich, "."))
    nompj = "PJ_K_" & numkm & "_" & v_numutile & "_0" & sext
    nompj_serv = chemin_pj & "/" & nompj
    If KF_PutFichier(nompj_serv, v_nomfich) = P_ERREUR Then
        Call Odbc_RollbackTrans
        Exit Sub
    End If
    If Odbc_AddNew("piecejointe", "pj_num", "pj_seq", False, lbid, _
                    "pj_unum", v_numutile, _
                    "pj_chemin", nompj, _
                    "pj_titre", "rapport", _
                    "pj_type_pere", "K", _
                    "pj_num_pere", numkm, _
                    "pj_session", "", _
                    "pj_numf_f", 0, _
                    "pj_numf_e", 0, _
                    "pj_numf_c", 0) = P_ERREUR Then
        Call Odbc_RollbackTrans
        Exit Sub
    End If
    
    Call Odbc_CommitTrans
    
End Sub

Private Sub initialiser()

    mnuDeposeExcel.tag = ""
    Call afficher_liste
    
    If tv.Nodes.Count = 0 Then
        MsgBox "Aucun fichier de Résultats"
        Unload Me
        Exit Sub
    Else
        Set tv.SelectedItem = tv.Nodes(1)
    End If
    tv.SetFocus

End Sub

Private Sub quitter()
    
    Unload Me
    
End Sub

Private Sub supprimer()

    Dim nomrep As String, nomfich As String, sql As String
    Dim numfich As Long, numdoc As Long
    Dim nd As Node
    Dim Num As Long
    Dim rpnum As Long, docnum As Long, fichnum As Long
    
    Set nd = tv.SelectedItem
    If nd.image = IMG_EXCEL Or nd.image = IMG_HTML Then
        rpnum = val(STR_GetChamp(nd.tag, "_", 1))
        docnum = val(STR_GetChamp(nd.tag, "_", 2))
        fichnum = val(STR_GetChamp(nd.tag, "_", 3))
        Call supprimerFichier(fichnum, docnum, rpnum)
    Else
        If STR_GetNbchamp(nd.tag, "_") = 2 Then
            ' si pas de slash => niveau RP : supprimer tous les documents et les fichiers
            rpnum = val(STR_GetChamp(nd.tag, "_", 1))
            Call supprimerRP(rpnum)
        ElseIf STR_GetNbchamp(nd.tag, "_") = 3 Then
            ' si un slash => niveau Document : supprimer ce document et ses fichiers
            rpnum = val(STR_GetChamp(nd.tag, "_", 1))
            docnum = val(STR_GetChamp(nd.tag, "_", 2))
            Call supprimerDocument(docnum, rpnum)
        End If
    End If
        
    Call initialiser
    
End Sub

Private Function supprimerRP(v_rpnum As Long)
    Dim nomrep As String
    Dim sql As String, rs As rdoResultset
        
    sql = "select * from rp_document where rpd_rpnum=" & v_rpnum
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        supprimerRP = ""
        Exit Function
    End If
    While Not rs.EOF
        Call supprimerDocument(rs("rpd_num"), v_rpnum)
        rs.MoveNext
    Wend
    nomrep = p_Chemin_Résultats & "/RP_" & v_rpnum
    If KF_EstRepertoire(nomrep, False) Then
        Call KF_EffacerRepertoire(nomrep)
    End If

End Function

Private Function supprimerDocument(v_docnum As Long, v_rpnum As Long)
    Dim nomrep As String
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from rp_fichier where rpf_rpnum=" & v_rpnum & " and rpf_rpdnum=" & v_docnum
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        supprimerDocument = ""
        Exit Function
    End If
    While Not rs.EOF
        Call supprimerFichier(rs("rpf_num"), v_docnum, v_rpnum)
        rs.MoveNext
    Wend
    nomrep = p_Chemin_Résultats & "/RP_" & v_rpnum & "/Doc_" & v_docnum
    If KF_EstRepertoire(nomrep, False) Then
        Call KF_EffacerRepertoire(nomrep)
    End If

End Function

Private Function supprimerFichier(v_fichnum As Long, v_docnum As Long, v_rpnum As Long)
    Dim nomrep As String
    Dim nomfich As String
    Dim sext As String
    Dim sql As String, rs As rdoResultset
    
    nomfich = p_Chemin_Résultats & "/RP_" & v_rpnum & "/Doc_" & v_docnum & "/" & v_fichnum  ' & p_PointExtensionXls
    sext = Positionne_Extension(nomfich)
    If sext = "" Then
        Call MsgBox("Fichier " & nomfich & " non trouvé sur le serveur ni en xls ni en xlsx")
    Else
        Call KF_EffacerFichier(nomfich & sext, False)
        Call KF_EffacerFichier(nomfich & ".txt", False)
        nomfich = p_Chemin_Résultats & "/RP_" & v_rpnum & "/Doc_" & v_docnum & "/" & v_fichnum & ".html"
        Call KF_EffacerFichier(nomfich, False)
        
        nomrep = p_Chemin_Résultats & "/RP_" & v_rpnum & "/Doc_" & v_docnum & "/" & v_fichnum & "_fichiers"
        Call KF_EffacerRepertoire(nomrep)
    End If
    sql = "delete from rp_fichier where rpf_num=" & v_fichnum & " and rpf_rpdnum=" & v_docnum
    Call Odbc_Cnx.Execute(sql)
End Function

Private Sub renommer()

    Dim nomrep As String, nomfich As String, sql As String
    Dim fichnum As Long, docnum As Long
    Dim nd As Node
    Dim s As String
    
    Set nd = tv.SelectedItem
    If nd.image = IMG_EXCEL Or nd.image = IMG_HTML Then
        fichnum = val(STR_GetChamp(nd.tag, "_", 3))
        docnum = val(STR_GetChamp(nd.tag, "_", 2))
        s = InputBox("Renommer", "Renommer", nd.Text)
        If s <> "" Then
            sql = "update rp_fichier set rpf_titre=" & Odbc_String(s) & " where rpf_num=" & fichnum & " and rpf_rpdnum=" & docnum
            Call Odbc_Cnx.Execute(sql)
            nd.Text = s
        End If
    Else
        If STR_GetNbchamp(nd.tag, "_") = 3 Then
            ' si un slash => niveau Document : supprimer ce document et ses fichiers
            docnum = val(STR_GetChamp(nd.tag, "_", 2))
            s = InputBox("Renommer", "Renommer", nd.Text)
            If s <> "" Then
                sql = "update rp_document set rpd_titre=" & Odbc_String(s) & " where rpd_num=" & docnum
                Call Odbc_Cnx.Execute(sql)
                nd.Text = s
            End If
        End If
    End If
        
End Sub

Private Function RecupUtilPpointNom(ByVal v_num As Long, _
                                     ByRef r_nom As String) As Integer

    Dim sql As String, nom As String, prenom As String
    
    If v_num = p_SuperUser Then
        r_nom = "Administrateur"
        RecupUtilPpointNom = P_OK
        Exit Function
    End If
    
    sql = "select U_Nom, U_Prenom" _
        & " from Utilisateur" _
        & " where U_Num=" & v_num
    If Odbc_RecupVal(sql, nom, prenom) = P_ERREUR Then
        RecupUtilPpointNom = P_ERREUR
        Exit Function
    End If
    If prenom <> "" Then
        r_nom = left$(prenom, 1) + ". "
    Else
        r_nom = ""
    End If
    r_nom = r_nom + nom
    
    RecupUtilPpointNom = P_OK

End Function

Private Function UtilEnvoyerMail(ByVal v_numutile As Long, _
                                  ByVal v_numutild As Long, _
                                  ByVal v_sujet As String, _
                                  ByVal v_message As Variant, _
                                  ByVal v_ajoutermess As Boolean, _
                                  Optional v_nomfich As Variant) As Integer

    Dim nomsrc As String, adrsrc As String, nomdest As String, adrdest As String, nomfich As String, sql As String
    Dim smtp_webmaster As String, smtp_user As String, smtp_password As String
    Dim cr As Integer, reponse As Integer
    Dim numzone As Long
    Dim mess As Variant
    Dim rs As rdoResultset
    Dim Frm As Form
    
    If Odbc_RecupVal("select pg_adrsmtp, pg_adrwebmaster, pg_smtpuser, pg_smtppassword from prmgen_http", _
                     p_smtp_adrsrv, smtp_webmaster, smtp_user, smtp_password) = P_ERREUR Then
        UtilEnvoyerMail = P_ERREUR
        Exit Function
    End If
    
    ' N° de la zone Adrmail
    If Odbc_RecupVal("select ZU_Num from ZoneUtil where ZU_Code='ADRMAIL'", _
                      numzone) = P_ERREUR Then
        UtilEnvoyerMail = P_ERREUR
        Exit Function
    End If
    
    ' Recherche l'adrmail du destinataire
    sql = "select UC_Valeur from UtilCoordonnee" _
        & " where UC_Type='U'" _
        & " and UC_TypeNum=" & v_numutild _
        & " and UC_ZUNum=" & numzone
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        UtilEnvoyerMail = P_ERREUR
        Exit Function
    End If
    ' Il n'en a pas
    If rs.EOF Then
        rs.Close
        UtilEnvoyerMail = P_NON
        Exit Function
    End If
    adrdest = rs("UC_Valeur").Value
    rs.Close
    Call RecupUtilPpointNom(v_numutild, nomdest)
    lblDiff.Caption = "Diffusion à " & nomdest & " (" & adrdest & ")"
    
    ' Emetteur < 0 = resp du document
    If v_numutile <> 0 Then
        Call RecupUtilPpointNom(Abs(v_numutile), nomsrc)
        sql = "select UC_Valeur from UtilCoordonnee" _
            & " where UC_Type='U'" _
            & " and UC_TypeNum=" & Abs(v_numutile) _
            & " and UC_ZUNum=" & numzone
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            UtilEnvoyerMail = P_ERREUR
            Exit Function
        End If
        If rs.EOF Then
            nomsrc = "KaliDoc"
            adrsrc = smtp_webmaster
        Else
            adrsrc = rs("UC_Valeur").Value
        End If
        rs.Close
        If v_numutile < 0 Then
            nomsrc = "KaliDoc"
            If adrsrc = "" Then
                adrsrc = smtp_webmaster
            End If
        End If
    Else
        nomsrc = "KaliDoc"
        adrsrc = smtp_webmaster
    End If
    
    ' Envoi du message
    If adrsrc <> "" And adrdest <> "" Then
        If v_ajoutermess Then
            ' Pas d'envoi
            If p_message_choisir = -1 Then
                UtilEnvoyerMail = P_OUI
                Exit Function
            End If
            ' Choix de l'envoi
            'If p_message_choisir = 1 Then
            '    Set frm = SaisirComplMail
            '    cr = SaisirComplMail.AppelFrm("Complément d'information pour " & nomdest, _
            '                                  v_message, _
            '                                  p_message_un_par_un, _
            '                                  p_message)
            '    Set frm = Nothing
            '    ' Annulation de l'envoi du mail
            '    If cr = 3 Then
            '        P_UtilEnvoyerMail = P_OUI
            '        Exit Function
            '    ' Annulation de l'envoi du mail pour tous
            '    ElseIf cr = 4 Then
            '        p_message_choisir = -1
            '        P_UtilEnvoyerMail = P_OUI
            '        Exit Function
            '    ' Envoi du mail pour tous
            '    ElseIf cr = 2 Then
            '        p_message_choisir = 0
            '    End If
            'End If
            'If p_message <> "" Then
            '    v_message = v_message & vbCrLf & vbCrLf & p_message
            'End If
        End If
        'Pour que les messages arrivent à Kalitech ...
        If InStr(smtp_webmaster, "kalitech@") > 0 Then
            nomsrc = "Utilisateur" & v_numutile
            adrsrc = "kalitech@kalitech.fr"
            nomdest = "Utilisateur" & v_numutild
            adrdest = "kalitech@kalitech.fr"
        End If
        If IsMissing(v_nomfich) Then
            nomfich = ""
        Else
            nomfich = v_nomfich
        End If
        Set Frm = FMail_SMTP
        cr = FMail_SMTP.EnvoiMessage(nomsrc, adrsrc, nomdest, adrdest, v_sujet, v_message, nomfich, smtp_user, smtp_password)
        Set Frm = Nothing
        If cr < 0 Then
            UtilEnvoyerMail = P_ERREUR
        End If
    Else
        UtilEnvoyerMail = P_NON
    End If

End Function

Private Sub depose_fichier(ByVal v_nd As Node, _
                         ByVal v_ext As String)

    Dim nomfich_serv As String, nomfich_loc As String, url As String
    Dim chemin_doc As String
    Dim rpnum As Long, docnum As Long, fichnum As Long
    Dim liberr As String
    Dim nomrep_serv As String
    Dim sext As String
    Dim NumFichier As String, FicOut As String, sChemin As String, FicOutHTML As String, FicTmp As String
    
    rpnum = val(STR_GetChamp(v_nd.tag, "_", 1))
    docnum = val(STR_GetChamp(v_nd.tag, "_", 2))
    fichnum = val(STR_GetChamp(v_nd.tag, "_", 3))
    nomfich_serv = p_Chemin_Résultats & "/RP_" & rpnum & "/Doc_" & docnum & "/" & fichnum ''v_ext
    sext = Positionne_Extension(nomfich_serv)
    nomfich_loc = Me.mnuDeposeExcel.tag

    nomrep_serv = p_Chemin_Résultats & "/RP_" & rpnum & "/Doc_" & docnum
    If Not KF_EstRepertoire(nomrep_serv, False) Then
        MsgBox nomrep_serv & " n'existe pas"
    End If
    ' l'ouvrir
    Call Public_VerifOuvrir(nomfich_loc, False, True, p_tbl_FichExcelPublier)
    ' transformer en HTML
    sChemin = p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil
    If Not FICH_EstRepertoire(sChemin, False) Then
        Call FICH_CreerRepComp(sChemin, False, False)
    End If
    FicOutHTML = sChemin & "\" & fichnum & ".html"
    Exc_wrk.SaveAs FileName:=FicOutHTML, _
        FileFormat:=44, ReadOnlyRecommended:=False, CreateBackup:=False
    Call Exc_wrk.Close
    
    ' Transfère le .xls sur le serveur
    If KF_PutFichier(nomfich_serv & sext, nomfich_loc) = P_ERREUR Then
        MsgBox "Erreur de dépot pour " & FicOut
    End If
    
    ' Transfère le .html sur le serveur
    If HTTP_Appel_PutDos(p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil, _
                         nomrep_serv & "/", False, False, liberr) <> HTTP_OK Then
        MsgBox liberr
    Else
        ' Vider le dossier local de transfert
        Call FICH_EffacerRep(sChemin)
        Me.mnuDeposeExcel.tag = ""
    End If

End Sub

Private Function Positionne_Extension(v_nomfichier)
    Dim sext As String
    
    ' Version Excel
    Call Excel_Init
    Exc_obj_Version = Exc_obj.Version
    If val(Exc_obj.Version) < 12 Then
        p_VersionExcel = "2003"
        p_PointExtensionXls = ".xls"
    Else
        p_VersionExcel = "2007"
        p_PointExtensionXls = ".xlsx"
    End If

    sext = ""
    If KF_FichierExiste(v_nomfichier & ".xls") Then
        sext = ".xls"
    Else
        If KF_FichierExiste(v_nomfichier & ".xlsx") Then
            sext = ".xlsx"
        End If
    End If
    Positionne_Extension = sext
End Function
Private Sub voir_fichier(ByVal v_nd As Node, _
                         ByVal v_ext As String)

    Dim nomfich_serv As String, nomfich_loc As String, url As String
    Dim lsS As String
    Dim sext As String
    Dim chemin_doc As String
    Dim rpnum As Long, docnum As Long, fichnum As Long
    
    ' Version Excel
    Call Excel_Init
    Exc_obj_Version = Exc_obj.Version
    If val(Exc_obj.Version) < 12 Then
        p_VersionExcel = "2003"
        p_PointExtensionXls = ".xls"
    Else
        p_VersionExcel = "2007"
        p_PointExtensionXls = ".xlsx"
    End If

    If v_ext = ".html" Then ' on veut ouvrir le html
        rpnum = val(STR_GetChamp(v_nd.tag, "_", 1))
        docnum = val(STR_GetChamp(v_nd.tag, "_", 2))
        fichnum = val(STR_GetChamp(v_nd.tag, "_", 3))
        If Odbc_RecupVal("select pg_cheminphp from prmgen_http", chemin_doc) = P_ERREUR Then
            Exit Sub
        End If
        url = p_HTTP_Résultats & "/RP_" & rpnum _
            & "/Doc_" & docnum _
            & "/" & fichnum & v_ext
        Call SYS_ExecShell("C:\Program Files\Internet Explorer\iexplore.exe " & url, True, True)
    Else    ' on veut le Excel
        rpnum = val(STR_GetChamp(v_nd.tag, "_", 1))
        docnum = val(STR_GetChamp(v_nd.tag, "_", 2))
        fichnum = val(STR_GetChamp(v_nd.tag, "_", 3))
        nomfich_serv = p_Chemin_Résultats & "/RP_" & rpnum & "/Doc_" & docnum & "/" & fichnum   '& ".xlsx"
        sext = Positionne_Extension(nomfich_serv)
        g_Ext_XLS = sext
        If sext = "" Then
            Call MsgBox("Fichier " & nomfich_serv & " non trouvé sur le serveur ni en xls ni en xlsx")
        Else
            nomfich_loc = p_chemin_appli & "\tmp\RP" & Format(Time, "hhmmss") & sext
            Call FICH_EffacerFichier(nomfich_loc, False)
            If KF_GetFichier(nomfich_serv & sext, nomfich_loc) = P_OK Then
                Call Excel_AfficherDoc(nomfich_loc, "", True, True)
                Me.mnuDeposeExcel.tag = nomfich_loc
            End If
        End If
    End If

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
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

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
        Exit Sub
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
End Sub

Private Sub Form_Load()

    g_form_active = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub

Private Sub mnuAfficheExcel_Click()
    
    Call voir_fichier(tv.SelectedItem, p_PointExtensionXls)
    
End Sub

Private Sub mnuAfficheHTML_Click()
    
    Call voir_fichier(tv.SelectedItem, ".html")
    
End Sub

Private Sub mnuDeposeExcel_Click()
    Call depose_fichier(tv.SelectedItem, p_PointExtensionXls)
End Sub

Private Sub mnuDiffuser_Click()

    Call Diffuser
    
End Sub

Private Sub mnuRenommer_Click()

     Call renommer
     
End Sub

Private Sub mnuSupprimer_Click()

    Call supprimer
    
End Sub

Private Sub tv_Click()
    
    Dim chemin As String
    
    If g_mode = "RES" And g_button = vbRightButton Then
        Call afficher_menu
    End If
    If g_mode = "SEL" And g_button = vbLeftButton Then
        If tv.SelectedItem.image = IMG_EXCEL Or tv.SelectedItem.image = IMG_HTML Then
            g_rp = tv.SelectedItem.tag
            Call quitter
        End If
    End If
End Sub


Private Sub tv_Expand(ByVal Node As ComctlLib.Node)

    g_button = -1
    
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
    End If
    
End Sub

Private Sub tv_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    g_button = Button

End Sub



