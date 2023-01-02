VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RapportType 
   BackColor       =   &H00808000&
   Caption         =   "Tableaux de Bord"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmHTTPD 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   6120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgbarHTTPDTaille 
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblHTTPDTaille 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblHTTPDTemps 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblMaj 
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
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Index           =   1
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "RapportType.frx":0000
      Top             =   1920
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Intégration de données de KaliForm vers Excel"
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
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton CmdBidon 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7080
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFC0C0&
         Height          =   1935
         Index           =   2
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "RapportType.frx":00E6
         Top             =   1920
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFC0C0&
         Height          =   1935
         Index           =   0
         Left            =   6480
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "RapportType.frx":01EE
         Top             =   1920
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   0
         Left            =   6480
         Picture         =   "RapportType.frx":0361
         ScaleHeight     =   975
         ScaleWidth      =   3135
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1485
         Index           =   1
         Left            =   9120
         Picture         =   "RapportType.frx":AD3B
         ScaleHeight     =   1485
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton ComGénérer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Générer et publier un tableau de Bord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   360
         MaskColor       =   &H00FFC0C0&
         Picture         =   "RapportType.frx":13A41
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1800
         Width           =   5715
      End
      Begin VB.CommandButton ComOuvrirModele 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paramétrage d'un modèle de tableau de bord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   360
         MaskColor       =   &H00FFC0C0&
         Picture         =   "RapportType.frx":13E0F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2880
         Width           =   5715
      End
      Begin VB.CommandButton ComResultats 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Accès aux fichiers résultats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   360
         MaskColor       =   &H00FFC0C0&
         Picture         =   "RapportType.frx":14216
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   5715
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   8880
         TabIndex        =   16
         Top             =   3840
         Width           =   2175
      End
   End
   Begin VB.Menu MnuQuitter 
      Caption         =   "Quitter"
   End
   Begin VB.Menu MnuAide 
      Caption         =   "Aide"
   End
   Begin VB.Menu MnuAPropos 
      Caption         =   "?"
      Begin VB.Menu MnuBaseServeur 
         Caption         =   "Base Serveur = "
      End
      Begin VB.Menu MnuBaseKD 
         Caption         =   "Base Locale ODBC = "
      End
      Begin VB.Menu MnuCheminAppli 
         Caption         =   "Chemin application"
      End
      Begin VB.Menu MnuCheminMod_Serveur 
         Caption         =   "Chemin des Modèles (Serveur)"
      End
      Begin VB.Menu MnuCheminMod_Local 
         Caption         =   "Chemin des Modèles (Local)"
      End
      Begin VB.Menu MnuCheminRes 
         Caption         =   "Chemin des Résultats"
      End
      Begin VB.Menu MnuSVersConf 
         Caption         =   "S_Vers_Conf"
      End
      Begin VB.Menu MnuFichierIni 
         Caption         =   "Fichier Ini ="
      End
      Begin VB.Menu MnuScmd 
         Caption         =   "Commande"
      End
      Begin VB.Menu MnuHTTP 
         Caption         =   "HTTP"
         Begin VB.Menu mnuHTTPDConfig1 
            Caption         =   "1"
         End
         Begin VB.Menu mnuHTTPDConfig2 
            Caption         =   "2"
         End
      End
      Begin VB.Menu MnuTrace 
         Caption         =   "Trace des Erreurs"
         Begin VB.Menu MnuTraceActive 
            Caption         =   "Activer"
         End
         Begin VB.Menu MnuTraceFichier 
            Caption         =   "Fichier des erreurs"
         End
         Begin VB.Menu MnuTraceVider 
            Caption         =   "Vider le Fichier"
         End
      End
   End
End
Attribute VB_Name = "RapportType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private g_form_active As Boolean
Private g_strdest As String
Private g_bcr As Boolean
Private g_CheminModele As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Function PrepareDossierResultat(v_path)
   Dim Chemin_Resultats As String
Test_Path:
   If FICH_EstRepertoire(v_path, False) Then
      Chemin_Resultats = v_path
      If Not FICH_EstRepertoire(Chemin_Resultats, False) Then
         MkDir (Chemin_Resultats)
         GoTo Test_Path
      Else
         Chemin_Resultats = v_path & "\"
         If Not FICH_EstRepertoire(Chemin_Resultats, False) Then
            MkDir (Chemin_Resultats)
            GoTo Test_Path
         End If
      End If
   Else
      MkDir (v_path)
      GoTo Test_Path
   End If
   PrepareDossierResultat = True
End Function

Private Function ControleFichierExterne(TabFichier(), CheminFichier, StrTitre, StrFeuille)
   Dim i As Integer
   Dim laDim As Integer
   Dim frm As Form
   Dim nomfich As String
   Dim NomXLS As String
   Dim ret As Integer
   
   laDim = 0
   ControleFichierExterne = -1
   On Error GoTo Err_ControleFichierExterne
   For i = 1 To UBound(TabFichier(), 2)
      If StrTitre = TabFichier(1, i) Then
         ControleFichierExterne = i
         Exit For
      End If
   Next i
   If ControleFichierExterne = -1 Then
      laDim = UBound(TabFichier(), 2)
      GoTo TestFichier
   End If
   GoTo Fin_Err_ControleFichierExterne:
Err_ControleFichierExterne:
   If Err = 9 Then
      'Resume Next
TestFichier:
      ret = MsgBox("Votre paramétrage fait référence à un fichier dénommé : " & StrTitre & Chr(13) & Chr(10) & "Voulez vous le choisir", vbQuestion + vbYesNo)
      If ret = vbYes Then
         Set frm = Com_ChoixFichier
         nomfich = Com_ChoixFichier.AppelFrm("Choix du Modèle Excel", "c:", CheminFichier, "*.xls", False)
         Set frm = Nothing
         If nomfich = "" Then
            ControleFichierExterne = -2
            Exit Function
         End If
      Else
         ControleFichierExterne = -2
         Exit Function
      End If
      CheminFichier = nomfich
      NomXLS = Excel_OuvrirDoc(CheminFichier, "", Exc_wrk, False)
      
      If FICH_FichierExiste(CheminFichier) Then
         laDim = laDim + 1
         ReDim Preserve TabFichier(4, laDim)
         TabFichier(1, laDim) = Trim(StrTitre)
         TabFichier(2, laDim) = Trim(CheminFichier)
         TabFichier(3, laDim) = ""
         TabFichier(4, laDim) = NomXLS
         ControleFichierExterne = laDim
      End If
   Else
      MsgBox "ControleFichierExterne : " & Err & " " & Error$
   End If
Fin_Err_ControleFichierExterne:
   On Error GoTo 0
End Function

Private Sub ComGénérer_Click()
    Dim frm As Form
    Dim bcr As Boolean
    
    p_ModePublication = "Publier"
        
    FctTrace ("======================================================")
    FctTrace ("RapportType Avant appel de PiloteExcelBis Pour Publier")
    FctTrace ("======================================================")
    
    Set frm = PiloteExcelBis
    bcr = PiloteExcelBis.AppelFrm(0, 0, "G")
    Set frm = Nothing

    If p_boolRetournerAuParam Then
        p_boolRetournerAuParam = False
        Call ComOuvrirModele_Click
    End If

End Sub

Private Sub ComGénérer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.Text(1).Visible = True
    Me.Text(0).Visible = False
    Me.Text(2).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.ComGénérer.MaskColor
    Me.ComGénérer.SetFocus
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub ComOuvrirModele_Click()
    Dim frm As Form
    Dim bcr As Boolean
    
    p_ModePublication = "Param"
    
    FctTrace ("=========================================================")
    FctTrace ("RapportType Avant appel de PiloteExcelBis pour Paramétrer")
    FctTrace ("=========================================================")

    Set frm = PiloteExcelBis
Faire:
    bcr = PiloteExcelBis.AppelFrm(0, 0, "P")
    Set frm = Nothing
    
    FctTrace ("RapportType Après appel de PiloteExcelBis")
    
    Call VerifSiVide
    
    FctTrace ("RapportType Après VerifSiVide")
    
End Sub

Private Function init_param_exe(ByVal v_scmd As String, _
                                ByRef r_numfor As Integer, _
                                ByRef r_numutil As Integer, _
                                ByRef r_nummodele As Integer, _
                                ByRef r_direct As Boolean) As Integer
    Dim numfor As String
    Dim nom_bdd As String
    Dim NumUtil As String
    Dim NumForm As String
    Dim NumModele As String
    Dim sql As String, rs As rdoResultset
    Dim nbprm As Integer, n As Integer, i As Integer
    
    'nbprm = STR_GetNbchamp(v_scmd, ";")
    'If nbprm < 3 Then
    '    init_param_exe = P_ERREUR
    '    GoTo ErrParametres
    'End If
    
    ' 1- Chemin export : p_CheminRapportType
    p_CheminRapportType = STR_GetChamp(v_scmd, ";", p_SCMD_CHEMIN_APPLI)

    If p_bool_ModeDebug Then MsgBox "p_CheminRapportType = " & p_CheminRapportType
    
    If p_CheminRapportType = "" Then
        init_param_exe = P_ERREUR
        GoTo ErrParametres
    Else
        Me.MnuCheminAppli.Caption = "Chemin Application = " & p_CheminRapportType
    End If
    
    ' 2- Nom Fichier Ini : p_CheminRapportType_Ini
    p_CheminRapportType_Ini = STR_GetChamp(v_scmd, ";", p_SCMD_CHEMIN_INI)
    If p_bool_ModeDebug Then MsgBox "p_CheminRapportType_Ini = " & p_CheminRapportType_Ini
    If p_CheminRapportType_Ini = "" Then
        init_param_exe = P_ERREUR
        GoTo ErrParametres
    Else
        p_CheminRapportType_Ini = p_CheminRapportType & "\" & p_CheminRapportType_Ini
    End If
    
    ' 3- Nom de la base
    nom_bdd = STR_GetChamp(v_scmd, ";", 2)
    If p_bool_ModeDebug Then MsgBox "nom_bdd = " & nom_bdd
    If nom_bdd = "VIDE" Or nom_bdd = "" Then
    Else
        Me.MnuBaseKD.Caption = "Base KaliDoc = " & nom_bdd
        p_nomBDD_ODBC = nom_bdd
    End If
    
    ' 4- Chemin application : p_chemin_appli
    p_chemin_appli = STR_GetChamp(v_scmd, ";", 3)
    RapportType.MnuCheminAppli.Caption = p_chemin_appli
    If p_bool_ModeDebug Then MsgBox "p_chemin_appli = " & p_chemin_appli
    If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
    End If
    
    ' Connexion à la base
    If nom_bdd <> "" And nom_bdd <> "VIDE" Then
        If Odbc_Init("PG", nom_bdd) = P_ERREUR Then
            init_param_exe = P_ERREUR
            Exit Function
        End If
    End If
    
    ' 5- Numéro de Formulaire
    numfor = STR_GetChamp(v_scmd, ";", 4)
    If numfor <> "" Then
        r_numfor = val(numfor)
    End If
    
    ' 6- Numéro d'utilisateur
    NumUtil = STR_GetChamp(v_scmd, ";", 5)
    If NumUtil <> "" Then
        r_numutil = val(NumUtil)
    End If

    ' 7- Numéro de modèle
    NumModele = STR_GetChamp(v_scmd, ";", 6)
    If NumModele <> "" Then
        r_nummodele = val(NumModele)
    End If

lab_fin:
    init_param_exe = P_OK
    Exit Function
            
ErrParametres:
    Call MsgBox("Usage : RapportType <Chemin Export>;<Nom fichier Ini>;<Nom BDD>;<Chemin Appli KaliDoc>;<NumFormulaire>;<NumUtil>;<NumModèle>" & vbCr & vbLf _
            & "cmd:" & v_scmd & vbCr & vbLf _
            & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
    init_param_exe = P_ERREUR
End Function

Private Sub ComOuvrirModele_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.Text(0).Visible = True
    Me.Text(1).Visible = False
    Me.Text(2).Visible = False
    Me.ComOuvrirModele.BackColor = Me.ComOuvrirModele.MaskColor
    Me.ComOuvrirModele.SetFocus
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub ComResultats_Click()
    Dim frm As Form
    Dim bcr As String
    
    Set frm = VoirFichiers
    bcr = VoirFichiers.AppelFrm("")
    Set frm = Nothing

End Sub

Private Sub ComResultats_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.Text(2).Visible = True
    Me.Text(0).Visible = False
    Me.Text(1).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.ComResultats.MaskColor
    Me.ComResultats.SetFocus
End Sub

Private Sub VerifSiVide()
    Dim sql As String, rs As rdoResultset
    Dim Chemin_Résultats As String
    Dim fso As FileSystemObject, fd As Integer
    Dim nbDossiers As Integer
    Dim Dossier As Variant
    Dim fileItem As Variant
    
    sql = "select * from rapport_type where rp_user_admin like '%U" & p_NumUtil & ";%'"
    sql = sql & " or rp_user_admin like '%U" & p_NumUtil & "=%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        Me.ComGénérer.Visible = False
        Me.ComResultats.Visible = False
    Else
        
        Me.ComGénérer.Visible = True
        
        If p_Drive_Résultats <> "" Then
            Chemin_Résultats = p_Drive_Résultats & p_Chemin_Résultats
        Else
            Chemin_Résultats = p_Chemin_Résultats
        End If
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        nbDossiers = 0
        ' Lire les sous répertoires
        If FICH_EstRepertoire(Chemin_Résultats, False) Then
            For Each Dossier In fso.GetFolder(Chemin_Résultats).SubFolders
                Set fileItem = fso.GetFolder(Dossier)
                nbDossiers = 1
                Exit For
            Next
        End If
        If nbDossiers > 0 Then
            Me.ComResultats.Visible = True
        Else
            Me.ComResultats.Visible = False
        End If
    End If
End Sub

Private Sub TesterResolutionEcran()
    Dim Anc_ScreenResolution As String
    
    Anc_ScreenResolution = ScreenResolution(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    
    Me.Caption = Me.Caption & "  (" & Anc_ScreenResolution & ")"
    If pAnc_Largeur_Ecran < pNew_Largeur_Ecran Or pAnc_Hauteur_Ecran < pNew_Hauteur_Ecran Then
        MsgBox "La résolution de votre écran va être adaptée"
    
        Call ResolutionEcran(pNew_Largeur_Ecran, pNew_Hauteur_Ecran)
        
        p_Bool_Modif_Resolution = True
    End If
End Sub


Public Function AppelFrm(ByRef v_bcr As Boolean, ByRef v_strdest As String, v_CheminModele As String) As String

    g_strdest = v_strdest
    g_bcr = v_bcr
    g_CheminModele = v_CheminModele
        
    Me.Show 1
    
    v_bcr = g_bcr
    
End Function

Private Function AppelDirect(ByVal v_unum As Integer) As String
    Dim sql As String, rs As rdoResultset
    Dim bad_util As Boolean
    
    sql = "select UAPP_Code, UAPP_MotPasse from Utilisateur, UtilAppli" _
        & " where UAPP_APPNum=" & p_appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_Num=UAPP_UNum" _
        & " and U_Num=" & v_unum
    'MsgBox sql
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        AppelDirect = 0
        Exit Function
    End If
    AppelDirect = 0
    If rs.EOF Then
        bad_util = True
    Else
        If rs("UAPP_Code").Value <> "" Then
            AppelDirect = UCase(rs("UAPP_Code").Value)
        End If
    End If
    rs.Close
    
End Function
Private Function FctOuvrirFichier(ByVal v_CheminFichier As String)
    
    If FICH_FichierExiste(v_CheminFichier) Then
        StartProcess v_CheminFichier
    Else
        MsgBox "Fichier " & v_CheminFichier & " introuvable"
    End If

End Function

Private Sub initialiser()

    Dim scmd As String
    Dim sret As String
    Dim direct As Boolean
    Dim stype_bdd As String, nom_bdd As String, nom_bddS As String
    Dim nbParam As Integer
    Dim ret As Integer
    Dim fd As Integer
    Dim param1 As String, param2 As String
    Dim sql As String, rs As rdoResultset
    Dim frm As Form, bcr As Boolean
    Dim Chemin_Résultats As String
    Dim NomFichierUser As String
    Dim fso As FileSystemObject
    Dim Dossier As Variant
    Dim fileItem As Variant
    Dim nbDossiers As Integer
    Dim NumForm As String, NumUtil As String, NumModele As String
    
    scmd = p_scmd
    RapportType.lblVersion.Caption = p_version_KaliRP

    
    p_bool_ModeDebug = True
'scmd = "DEBUG;C:\KaliDoc\RapportType;RapportType_dcss_dcss.ini;kalidoc_dcss_dcss;C:\KaliDoc\;;"
    
    ' utiliser le .ini local
    p_Mode_FctTrace = False
    If InStr(scmd, "DEBUG;") > 0 Then
        scmd = Replace(scmd, "DEBUG;", "")
        p_Mode_FctTrace = True
        p_bool_ModeDebug = True
    End If
    
    'If p_bool_ModeDebug Then MsgBox "scmd en entrée=" & scmd

    'Set p_HTTP_Form_Menu = RapportType
    'MsgBox scmd
    
Init:
    Me.MnuScmd.Caption = "Commande=" & scmd

Début:
    nbParam = STR_GetNbchamp(scmd, ";")
    
    ' 1 Chemin Export
    ' 2 Nom BDD
    ' 3 Chemin Application
    ' 4 num formulaire
    ' 5 Num Utilisateur

    ' variables déjà éventuellement initialisées :
    ' p_chemin_appli
    ' s_type_bdd
    ' nom_bdd
    ' p_nomini
    
    p_NumUtil = 0
    
    If STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT) <> "" Then
        NumForm = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMFORM)
        If NumForm = "" Then
            p_NumForm = 0
        Else
            p_NumForm = NumForm
        End If
        NumUtil = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMUTIL)
        If NumUtil = "" Then
            p_NumUtil = 0
        Else
            p_NumUtil = NumUtil
        End If
        NumModele = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMFILTRE)
        If NumModele = "" Then
            p_NumModele = 0
        Else
            p_NumModele = NumModele
        End If
    End If
'MsgBox NumForm
'MsgBox NumUtil
'MsgBox NumModele
    p_CheminRapportType_Ini = p_nomini
'MsgBox "p_CheminRapportType_Ini=" & p_CheminRapportType_Ini
    If FICH_FichierExiste(p_CheminRapportType_Ini) Then
        ' Voir si le fichier .ini existe
        Me.MnuFichierIni.Caption = "Fichier Ini = " & p_CheminRapportType_Ini
        ' connexion à la base de données
'MsgBox "avant test p_nomBDD_ODBC=" & p_nomBDD_ODBC
'MsgBox p_nomBDD_ODBC = ""
'MsgBox p_nomBDD_ODBC = "VIDE"
        
        If p_nomBDD_ODBC = "" Or p_nomBDD_ODBC = "VIDE" Then
            ' 2- Nom de la base
'MsgBox "avant nom_bdd"
            nom_bdd = SYS_GetIni("Base", "NOM", p_CheminRapportType_Ini)
'MsgBox "nom_bdd=" & nom_bdd
            If nom_bdd = "" Then
                Call MsgBox("Pas de nom de base (ODBC).", vbInformation + vbOKOnly)
                End
            Else
                Me.MnuBaseKD.Caption = "Base ODBC Locale = " & nom_bdd
                p_nomBDD_ODBC = nom_bdd
            End If
        Else
            Me.MnuBaseKD.Caption = "Base ODBC Locale = " & p_nomBDD_ODBC
        End If
        ' Connexion à la base
'MsgBox "avant odbc_init"
        If Odbc_Init("PG", p_nomBDD_ODBC) = P_ERREUR Then
            MsgBox "Connexion à la base " & nom_bdd & " impossible"
            End
        End If
'MsgBox "après odbc_init"
        
        ' Récupérer l'adresse HTTP du serveur (champ pg_serveur)
        
        'p_AdrServeur = SYS_GetIni("Chemins", "HTTP_SERVEUR", p_CheminRapportType_Ini)
        sql = "select PG_SERVEUR from PRMGEN_HTTP"
'MsgBox "sql=" & sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_SERVEUR à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_SERVEUR à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
'MsgBox "p_AdrServeur=" & p_AdrServeur
            p_AdrServeur = rs("pg_serveur")
'MsgBox "p_AdrServeur=" & p_AdrServeur
            rs.Close
            If p_AdrServeur = "" Then
                MsgBox "Adresse HTTP du serveur non renseignée"
            End If
        End If
            
        ' Récupérer HTTP : maxparfichier
        sql = "select PG_HTTP_MAXPARFICHIER from PRMGEN_HTTP"
'MsgBox "sql=" & sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_HTTP_MAXPARFICHIER à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_HTTP_MAXPARFICHIER à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
            p_HTTP_MaxParFichier = rs("pg_http_maxparfichier")
            rs.Close
            If p_HTTP_MaxParFichier = 0 Then
                MsgBox "Valeur HTTP de PG_HTTP_MAXPARFICHIER non renseignée"
            End If
            Me.mnuHTTPDConfig1.Caption = "Taille par Fichier " & p_HTTP_MaxParFichier
        End If
        
        ' Récupérer HTTP : maxparpaquet
        sql = "select PG_HTTP_MAXPARPAQUET from PRMGEN_HTTP"
'MsgBox "sql=" & sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_HTTP_MAXPARPAQUET à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_HTTP_MAXPARPAQUET à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
            p_HTTP_MaxParPaquet = rs("pg_http_maxparpaquet")
            rs.Close
            If p_HTTP_MaxParPaquet = 0 Then
                MsgBox "Valeur HTTP de PG_HTTP_MAXPARPAQUET non renseignée"
            End If
            Me.mnuHTTPDConfig2.Caption = "Taille par Paquet " & p_HTTP_MaxParPaquet
        End If
        
        ' Récupérer l'adresse s_Vers_Conf
        p_S_Vers_Conf = SYS_GetIni("Chemins", "S_VERS_CONF", p_CheminRapportType_Ini)
'MsgBox "p_S_Vers_Conf=" & p_S_Vers_Conf
        If p_S_Vers_Conf <> "" Then
            Me.MnuSVersConf.Caption = "s_vers_conf = " & p_S_Vers_Conf
        End If
        
        ' Récupérer le Chemin de Dépot HTTP du serveur  (champ pg_http_chemindepot)
        'p_HTTP_CheminDepot = SYS_GetIni("Chemins", "CHEMIN_DEPOT_HTTP", p_CheminRapportType_Ini)
        sql = "select PG_HTTP_CHEMINDEPOT from PRMGEN_HTTP"
'MsgBox "sql=" & sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_HTTP_CHEMINDEPOT à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_HTTP_CHEMINDEPOT à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
            p_HTTP_CheminDepot = rs("PG_HTTP_CHEMINDEPOT")
            rs.Close
            If p_HTTP_CheminDepot = "" Then
                MsgBox "Chemin de Dépot HTTP du serveur non renseignée"
            End If
        End If
        
        ' Récupérer le Chemin de destination des tableaux publiés dans KaliDoc (champ pg_chemindoc)
        'p_Drive_KaliDoc = SYS_GetIni("Chemins", "DRIVE_KALIDOC", p_CheminRapportType_Ini)
        'p_Path_KaliDoc = SYS_GetIni("Chemins", "PATH_KALIDOC", p_CheminRapportType_Ini)
        sql = "select PG_CHEMINDOC from PRMGEN_HTTP"
'MsgBox "sql=" & sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_CHEMINDOC à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_CHEMINDOC à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
            p_Path_KaliDoc = rs("PG_CHEMINDOC")
            rs.Close
            If p_Path_KaliDoc = "" Then
                MsgBox "Chemin de Publication dans KaliDoc non renseignée"
            End If
        End If
                
        If p_nomBDD_SERVEUR = "" Then
            ' 2- Nom de la base sur le serveur
            nom_bddS = SYS_GetIni("Base", "Nom_BDD_SERVEUR", p_CheminRapportType_Ini)
'MsgBox "nom_bddS=" & nom_bddS
            If nom_bddS = "" Then
                Call MsgBox("Pas de nom de base (Serveur).", vbInformation + vbOKOnly)
                End
            Else
                Me.MnuBaseServeur.Caption = "Base Serveur = " & nom_bddS
                p_nomBDD_SERVEUR = nom_bddS
            End If
        End If
                
'MsgBox "p_chemin_appli=" & p_chemin_appli
        ' 3- Chemin application : p_chemin_appli
        If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
            RapportType.MnuCheminAppli.Caption = p_chemin_appli
            p_chemin_appli = SYS_GetIni("Chemins", "Chemin_Appli", p_CheminRapportType_Ini)
            If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
                Call MsgBox("<Chemin Appli> est vide." & vbCr & vbLf _
                            & "Usage : KaliDoc <Chemin Export>;<Nom BDD>;<Chemin Appli>", vbInformation + vbOKOnly)
                End
            End If
        Else
            p_CheminRapportType = p_chemin_appli & "/KaliRP"
        End If
        
        sql = "select APP_Num from Application" _
            & " where APP_Code='KALIDOC'"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger p_appli_kalidoc à partir de la base"
            End
        End If
        If Not rs.EOF Then
            p_appli_kalidoc = rs("APP_Num").Value
        Else
            MsgBox "Impossible de déterminer p_appli_kalidoc"
            End
        End If
        rs.Close
            
        ' Récupérer le chemin des Modèles sur le serveur
        p_Drive_Modeles_Serveur = SYS_GetIni("Chemins", "Drive_Modeles_Serveur", p_CheminRapportType_Ini)
        p_Chemin_Modeles_Serveur = SYS_GetIni("Chemins", "Path_Modeles_Serveur", p_CheminRapportType_Ini)
        Me.MnuCheminMod_Serveur.Caption = "Modèles => " & p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur
        If p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur = "" Then
            MsgBox "Chemin des Modèles non renseigné"
            End
        ' ElseIf Not FICH_EstRepertoire(p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur, False) Then
        '     MsgBox "Chemin des Modèles " & p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur & " introuvable"
        End If
        
        ' Récupérer le chemin des Modèles en local
        p_Drive_Modeles_Local = SYS_GetIni("Chemins", "Drive_Modeles_Local", p_CheminRapportType_Ini)
        p_Chemin_Modeles_Local = SYS_GetIni("Chemins", "Path_Modeles_Local", p_CheminRapportType_Ini)
        Me.MnuCheminMod_Local.Caption = "Modèles => " & p_Drive_Modeles_Local & p_Chemin_Modeles_Local
        If p_Drive_Modeles_Local & p_Chemin_Modeles_Local = "" Then
            MsgBox "Chemin des Modèles non renseigné"
            End
        ElseIf Not FICH_EstRepertoire(p_Drive_Modeles_Local & p_Chemin_Modeles_Local, False) Then
            MsgBox "Chemin des Modèles " & p_Drive_Modeles_Local & p_Chemin_Modeles_Local & " introuvable"
        End If
            
        '' Récupérer le chemin de RapportType.exe sur le serveur
        'p_RapportTypeExe = SYS_GetIni("Chemins", "RapportTypeExe", p_CheminRapportType_Ini)
        'If p_RapportTypeExe = "" Then
        '    MsgBox "Chemin de RapportTypeExe non renseigné"
        '    End
        'End If
        
        ' Récupérer le chemin des Résultats (en local)
        p_Drive_Résultats = SYS_GetIni("Chemins", "Drive_Resultats", p_CheminRapportType_Ini)
        p_Chemin_Résultats = SYS_GetIni("Chemins", "Path_Resultats", p_CheminRapportType_Ini)
        Me.MnuCheminRes.Caption = "Résultats => " & p_Drive_Résultats & p_Chemin_Résultats
        If p_Drive_Résultats & p_Chemin_Résultats = "" Then
            MsgBox "Chemin des Résultats non renseigné"
            End
        ElseIf Not FICH_EstRepertoire(p_Drive_Résultats & p_Chemin_Résultats, False) Then
            MsgBox "Chemin des Résultats " & p_Drive_Résultats & p_Chemin_Résultats & " introuvable"
        End If
    End If

    p_Chemin_FichierTrace = p_chemin_appli & "\KaliRP_Trace.txt"


'MsgBox "ici"
    If p_NumUtil = 0 Then
        ret = P_SaisirUtilIdent(10, 10, 10, 10)
        If ret <> P_OUI Then
            'MsgBox "Utilisateur non autorisé"
            End
        End If
    Else
        p_CodeUtil = AppelDirect(p_NumUtil)
    End If
    
    If p_CodeUtil <> "ROOT" Then
        sql = "select * from utilisateur where u_num = " & p_NumUtil
        If Odbc_SelectV(sql, rs) <> P_ERREUR Then
            If Not rs.EOF Then
                Me.Caption = Me.Caption & "   (" & LCase(rs("u_prenom")) & " " & UCase(rs("u_nom")) & ")"
            Else
                MsgBox "Utilisateur non autorisé"
                End
            End If
        End If
    Else
        Me.Caption = Me.Caption & "   (SuperUtilisateur)"
        Me.MnuFichierIni.Caption = "Ouvrir le .Ini"
    End If

    p_Bool_Modif_Resolution = False
    
    TesterResolutionEcran

    Call FRM_ResizeForm(Me, Me.Width, Me.Height)

    Call VerifSiVide
    
    NomFichierUser = p_Drive_Modeles_Local & p_Chemin_Modeles_Local & "/User_" & p_NumUtil & ".User"
    If Not FICH_FichierExiste(NomFichierUser) Then
        If Not FICH_OuvrirFichier(NomFichierUser, FICH_ECRITURE, fd) = P_ERREUR Then
            Print #fd, "0"
            Close #fd
        End If
    End If
    
'p_NumModele = 0
'p_NumForm = 0
    If p_NumModele > 0 Then
EncoreModele:
        Set frm = PiloteExcelBis
        bcr = PiloteExcelBis.AppelFrm(p_NumModele, p_NumForm, "P")
        Set frm = Nothing
        If p_Appel_Création_Nouveau_Modele Then
            GoTo EncoreModele
        End If
        p_NumForm = 0
        'End
    End If
    If p_NumForm > 0 Then
EncoreFormulaire:
        Set frm = PiloteExcelBis
        bcr = PiloteExcelBis.AppelFrm(p_NumModele, p_NumForm, "P")
        Set frm = Nothing
        If p_Appel_Création_Nouveau_Modele Then
            GoTo EncoreFormulaire
        End If
        'End
    End If
    RapportType.MnuCheminAppli.Caption = "Chemin_Appli = " & p_chemin_appli
End Sub

Private Sub Form_Activate()
    
    If g_form_active Then
        Exit Sub
    End If
    
    Call FRM_ResizeForm(Me, 0, 0)
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub Form_Load()

    g_form_active = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Unload
    Exc_obj.Quit
    Set Exc_obj = Nothing
    
    If p_Bool_Modif_Resolution Then
        Call ResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    End If

Err_Unload:
   Resume Next
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Me.Text(0).Visible = False
    Me.Text(1).Visible = False
    Me.Text(2).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub MnuAide_Click()
    Call Appel_Aide
    
End Sub

Private Sub MnuFichierIni_Click()
    FctOuvrirFichier (p_CheminRapportType_Ini)
End Sub


Private Sub MnuQuitter_Click()
   On Error GoTo Err_Unload
   Exc_obj.Quit
   Set Exc_obj = Nothing
   End
Err_Unload:
   Resume Next
End Sub


Private Sub MnuTrace_Click()
    
    If FICH_FichierExiste(p_Chemin_FichierTrace) Then
        MnuTraceVider.Enabled = True
    Else
        MnuTraceVider.Enabled = False
    End If

End Sub

Private Sub MnuTraceActive_Click()
    
    If MnuTraceActive.Caption = "Activer" Then
        p_Mode_FctTrace = True
        MnuTraceActive.Caption = "Désactiver"
        Exit Sub
    End If
    If MnuTraceActive.Caption = "Désactiver" Then
        p_Mode_FctTrace = False
        MnuTraceActive.Caption = "Activer"
    End If
End Sub

Private Sub MnuTraceFichier_Click()
    
    FctOuvrirFichier (p_Chemin_FichierTrace)

End Sub

Private Sub MnuTraceVider_Click()
    If FICH_FichierExiste(p_Chemin_FichierTrace) Then
        Kill (p_Chemin_FichierTrace)
    End If
End Sub
