VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MenuG 
   BackColor       =   &H8000000C&
   Caption         =   "KaliRP"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11325
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   5175
      Index           =   1
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   11265
      TabIndex        =   0
      Top             =   0
      Width           =   11325
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
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
         Height          =   5175
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   11295
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   6120
            TabIndex        =   21
            Top             =   240
            Width           =   5055
            Begin VB.PictureBox Picture1 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1485
               Index           =   2
               Left            =   3120
               Picture         =   "Menu.frx":0000
               ScaleHeight     =   1485
               ScaleWidth      =   1815
               TabIndex        =   23
               Top             =   120
               Width           =   1815
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   975
               Index           =   0
               Left            =   120
               Picture         =   "Menu.frx":8D06
               ScaleHeight     =   975
               ScaleWidth      =   3135
               TabIndex        =   22
               Top             =   360
               Width           =   3135
            End
         End
         Begin VB.TextBox Text 
            BackColor       =   &H00C0C0FF&
            Height          =   2175
            Index           =   1
            Left            =   6120
            MultiLine       =   -1  'True
            TabIndex        =   19
            Text            =   "Menu.frx":136E0
            Top             =   2040
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.TextBox Text 
            BackColor       =   &H0080C0FF&
            Height          =   1935
            Index           =   3
            Left            =   6600
            MultiLine       =   -1  'True
            TabIndex        =   18
            Text            =   "Menu.frx":137BB
            Top             =   1920
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.CommandButton ComAssistant 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Créer un modèle de  rapport avec l'assistant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Picture         =   "Menu.frx":13843
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3840
            Width           =   5715
         End
         Begin VB.CommandButton ComResultats 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Accéder aux rapports générés"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Picture         =   "Menu.frx":15385
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   5715
         End
         Begin VB.CommandButton ComOuvrirModele 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Paramétrer un modèle de rapport"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Picture         =   "Menu.frx":1582A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2280
            Width           =   5715
         End
         Begin VB.CommandButton ComGénérer 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Générer et publier un rapport"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            MaskColor       =   &H00FFC0C0&
            Picture         =   "Menu.frx":15C31
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1440
            Visible         =   0   'False
            Width           =   5715
         End
         Begin VB.TextBox Text 
            BackColor       =   &H00FFC0C0&
            Height          =   1935
            Index           =   0
            Left            =   6600
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "Menu.frx":15FFF
            Top             =   1920
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.TextBox Text 
            BackColor       =   &H00C0E0FF&
            Height          =   1935
            Index           =   2
            Left            =   6600
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "Menu.frx":16167
            Top             =   1920
            Visible         =   0   'False
            Width           =   4335
         End
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
            TabIndex        =   8
            Top             =   1680
            Visible         =   0   'False
            Width           =   1755
         End
         Begin ComctlLib.ProgressBar PgFeuille 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   3120
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label lblWait 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   3360
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.Label LblGauge 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3960
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.Label lblVersion 
            BackColor       =   &H00C0C0FF&
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
            Height          =   495
            Left            =   6120
            TabIndex        =   14
            Top             =   4320
            Width           =   5055
         End
      End
      Begin VB.Frame FrmHTTPD 
         BackColor       =   &H00C0C0C0&
         Height          =   1935
         Left            =   6240
         TabIndex        =   1
         Top             =   2520
         Visible         =   0   'False
         Width           =   4935
         Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
            Height          =   255
            Left            =   2160
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
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
            TabIndex        =   6
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblHTTPDTemps 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblHTTPDTaille 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   1935
         End
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
      Begin VB.Menu MnuVersionExcel 
         Caption         =   "Mode Excel"
         Begin VB.Menu MnuVersionExcelCharts 
            Caption         =   "jusqu'à 2003"
         End
         Begin VB.Menu MnuVersionExcelShapes 
            Caption         =   "supérieure à 2003"
         End
      End
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
      Begin VB.Menu MnuCheminLocalCle 
         Caption         =   "Clé du Chemin Local"
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
            Caption         =   "Activer les Traces"
         End
         Begin VB.Menu MnuTraceFichier 
            Caption         =   "Fichier des Traces"
         End
         Begin VB.Menu MnuTraceVider 
            Caption         =   "Vider le Fichier"
         End
      End
   End
End
Attribute VB_Name = "MenuG"
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
   Dim Frm As Form
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
         Set Frm = Com_ChoixFichier
         nomfich = Com_ChoixFichier.AppelFrm("Choix du Modèle Excel", CheminFichier, "*.xls", False)
         Set Frm = Nothing
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

Private Sub ComAssistant_Click()
    Dim Frm As Form
    Dim bcr As Boolean
    
    p_ModePublication = "Param"
    
    FctTrace ("=========================================================")
    FctTrace ("RapportType Avant appel de PiloteExcelBis pour Assistant ")
    FctTrace ("=========================================================")

    p_ModeAssistant = True
    Set Frm = PiloteExcelBis
Faire:
    bcr = PiloteExcelBis.AppelFrm(0, 0, "P")
    Set Frm = Nothing
    
    FctTrace ("RapportType Après appel de PiloteExcelBis")
    
    Call VerifSiVide
    
    FctTrace ("ComAssistant_Click RapportType Après VerifSiVide")
    

End Sub

Private Sub ComAssistant_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Text(3).Visible = True
    Me.Text(0).Visible = False
    Me.Text(2).Visible = False
    Me.Text(1).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComAssistant.BackColor = Me.ComGénérer.MaskColor
    Me.ComAssistant.SetFocus

End Sub

Private Sub ComGénérer_Click()
    Dim Frm As Form
    Dim bcr As Boolean
    
    p_ModePublication = "Publier"
    p_ModeAssistant = False
        
    FctTrace ("======================================================")
    FctTrace ("RapportType Avant appel de PiloteExcelBis Pour Publier")
    FctTrace ("======================================================")
    
    Set Frm = PiloteExcelBis
    bcr = PiloteExcelBis.AppelFrm(0, 0, "G")
    Set Frm = Nothing

    If p_boolRetournerAuParam Then
        p_boolRetournerAuParam = False
        Me.LblGauge.Visible = True
        Me.PgFeuille.Visible = True
        FctTrace ("ComGénérer_Click appel ComOuvrirModele_Click")
        Call ComOuvrirModele_Click
        Me.LblGauge.Visible = False
        Me.PgFeuille.Visible = False
    End If

End Sub

Private Sub ComGénérer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Text(1).Visible = True
    Me.Text(0).Visible = False
    Me.Text(2).Visible = False
    Me.Text(3).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.ComGénérer.MaskColor
    Me.ComGénérer.SetFocus
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
    Me.ComAssistant.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub ComOuvrirModele_Click()
    Dim Frm As Form
    Dim bcr As Boolean
    
    p_ModePublication = "Param"
    p_ModeAssistant = False
    
    FctTrace ("=========================================================")
    FctTrace ("RapportType Avant appel de PiloteExcelBis pour Paramétrer")
    FctTrace ("=========================================================")

    Set Frm = PiloteExcelBis
Faire:
    bcr = PiloteExcelBis.AppelFrm(0, 0, "P")
    Set Frm = Nothing
    
    FctTrace ("RapportType Après appel de PiloteExcelBis")
    
    Call VerifSiVide
    
    FctTrace ("ComOuvrirModele_Click RapportType Après VerifSiVide")
    
End Sub

Private Function init_param_exe(ByVal v_scmd As String, _
                                ByRef r_numfor As Integer, _
                                ByRef r_numutil As Integer, _
                                ByRef r_nummodele As Integer, _
                                ByRef r_direct As Boolean) As Integer
    Dim numfor As String
    Dim nom_bdd As String
    Dim numutil As String
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
    MnuCheminAppli.Caption = p_chemin_appli
    If p_bool_ModeDebug Then MsgBox "p_chemin_appli = " & p_chemin_appli
    If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
    End If
    
    ' Connexion à la base
    If nom_bdd <> "" And nom_bdd <> "VIDE" Then
        If Odbc_Init("PG", nom_bdd, True) = P_ERREUR Then
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
    numutil = STR_GetChamp(v_scmd, ";", 5)
    If numutil <> "" Then
        r_numutil = val(numutil)
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

Private Sub ComOuvrirModele_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Text(0).Visible = True
    Me.Text(1).Visible = False
    Me.Text(2).Visible = False
    Me.Text(3).Visible = False
    Me.ComOuvrirModele.BackColor = Me.ComOuvrirModele.MaskColor
    Me.ComOuvrirModele.SetFocus
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
    Me.ComAssistant.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub ComResultats_Click()
    
    Dim Frm As Form
    Dim rp As String
    
    p_ModeAssistant = False
    Set Frm = VoirFichiers
    Call VoirFichiers.AppelFrm(0, "RES", rp)
    Set Frm = Nothing
    Call VerifSiVide

End Sub

Private Sub ComResultats_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Text(2).Visible = True
    Me.Text(0).Visible = False
    Me.Text(1).Visible = False
    Me.Text(3).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.ComResultats.MaskColor
    Me.ComAssistant.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.SetFocus
End Sub

Private Sub VerifSiVide()
    Dim sql As String, rs As rdoResultset
    Dim rstmp As rdoResultset
    Dim rs1 As rdoResultset
    Dim op As String
    Dim Chemin_Résultats As String
    Dim fso As FileSystemObject, fd As Integer
    Dim nbDossiers As Integer
    Dim Dossier As Variant
    Dim fileItem As Variant
    Dim lnb As Long
    Dim totlnb As Long
    
    FctTrace ("Début VerifSiVide")
    Me.ComResultats.Visible = False
    sql = "select * from rapport_type where rp_user_admin like '%U" & p_NumUtil & ";%'"
    sql = sql & " or rp_user_admin like '%U" & p_NumUtil & "=%'"
    If p_CodeUtil = "ROOT" Then
        sql = "select * from rapport_type"   ' where rp_user_admin like '%U" & p_NumUtil & ";%'"
        'sql = sql & " or rp_user_admin like '%U" & p_NumUtil & "=%'"
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        Me.ComGénérer.Visible = False
        Me.ComResultats.Visible = False
        rs.Close
        GoTo FinVerif
    Else
        sql = "select rp_num, rp_user_admin from rapport_type where (rp_user_admin like '%U" & p_NumUtil & "=PARAM:PUBLIER%' or rp_user_admin like '%U" & p_NumUtil & "=:PUBLIER:%')  " _
            & " and rp_num in (select rpd_rpnum from rp_document)"
        If p_CodeUtil = "ROOT" Then
            sql = "select rp_num, rp_user_admin from rapport_type where rp_num in (select rpd_rpnum from rp_document)"
        End If
        Call Odbc_SelectV(sql, rstmp)
        While Not rstmp.EOF
            'Debug.Print rstmp("rp_num") & " " & rstmp("rp_user_admin")
            sql = "select count(*) from rp_document where rpd_rpnum=" & rstmp("rp_num")
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                Exit Sub
            End If
            'Debug.Print lnb
            totlnb = totlnb + lnb
            rstmp.MoveNext
        Wend
        rstmp.Close
        rs.Close
        Me.ComGénérer.Visible = IIf(totlnb > 0, True, False)
        
        Chemin_Résultats = p_Chemin_Résultats
    
        FctTrace ("VerifSiVide Chemin_Résultats=" & Chemin_Résultats & " VerifSiVide CreateObject Scripting.FileSystemObject")
        'Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo 0
        nbDossiers = 0
        ' Lire les sous répertoires
        FctTrace ("VerifSiVide Chemin_Résultats=" & Chemin_Résultats)
        If KF_EstRepertoire(Chemin_Résultats, False) Then
            GoTo LabTesteResultat
            'For Each Dossier In fso.GetFolder(Chemin_Résultats).SubFolders
            '    Set fileItem = fso.GetFolder(Dossier)
            '    nbDossiers = 1
            '    Exit For
            'Next
        Else
            Me.ComResultats.Visible = False
        End If
    End If
    
    ' seconde méthode
LabTesteResultat:
            
    sql = "select count(*) from rp_fichier, rapport_type where rpf_rpnum=rp_num and "
    sql = sql & "( rp_user_admin like '%U" & p_NumUtil & "=PARAM:PUBLIER:RESULTAT%' Or rp_user_admin like '%U" & p_NumUtil & "=:PUBLIER:RESULTAT%' Or rp_user_admin like '%U" & p_NumUtil & "=:RESULTAT%')"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        Exit Sub
    End If
    If lnb > 0 Then
        Me.ComResultats.Visible = True
        If lnb = 1 Then
            Me.ComResultats.Caption = "Accéder au rapport généré"
        Else
            Me.ComResultats.Caption = "Accéder aux " & lnb & " rapports générés"
        End If
'        cmd(CMD_VOIR_RESULTATS).Visible = True
'        sql = "select count(*) from rp_fichier where rpf_rpnum=" & g_numModele _
            & " and rpf_diff_faite='f'"
'        Call Odbc_Count(sql, lnb)
'        If lnb > 0 Then
'            cmd(CMD_VOIR_RESULTATS).BackColor = vbRed
'        Else
'            cmd(CMD_VOIR_RESULTATS).BackColor = &HCCCCCC
'        End If
    Else
        Me.ComResultats.Visible = False
'        cmd(CMD_VOIR_RESULTATS).Visible = False
    End If
FinVerif:
    FctTrace ("Fin VerifSiVide")
    Exit Sub
ErrFSO:
    Call MsgBox("Erreur générée par Scripting.FileSystemObject" & vbCrLf & Error$)
    Resume LabTesteResultat
End Sub

Private Sub TesterResolutionEcran()
    Dim Anc_ScreenResolution As String
    
    Anc_ScreenResolution = ScreenResolution(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    '     ScreenResolution = "Vidéo " & r_Largeur_Ecran & " x  " & r_Hauteur_Ecran
    'MsgBox Anc_ScreenResolution
    'Call GetResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    Anc_ScreenResolution = "Vidéo " & pAnc_Largeur_Ecran & " x  " & pAnc_Hauteur_Ecran
    ' MsgBox Anc_ScreenResolution
    
    Me.Caption = Me.Caption & "  (" & Anc_ScreenResolution & ")"
    If pAnc_Largeur_Ecran < pNew_Largeur_Ecran Or pAnc_Hauteur_Ecran < pNew_Hauteur_Ecran Then
        MsgBox "La résolution de votre écran va être adaptée"
    
        Call MetResolutionEcran(pNew_Largeur_Ecran, pNew_Hauteur_Ecran)
        
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
    Dim sret As String, s As String
    Dim direct As Boolean
    Dim stype_bdd As String, nom_bdd As String, nom_bddS As String
    Dim nbParam As Integer, i As Integer, n As Integer
    Dim ret As Integer
    Dim fd As Integer
    Dim param1 As String, param2 As String
    Dim sql As String, rs As rdoResultset
    Dim Frm As Form, bcr As Boolean
    Dim Chemin_Résultats As String
    Dim NomFichierUser As String
    Dim fso As FileSystemObject
    Dim Dossier As Variant
    Dim fileItem As Variant
    Dim nbDossiers As Integer
    Dim NumForm As String, numutil As String, NumModele As String
    Dim CleTmp As String, NewCle As String
    Dim sCleTmp As String
    Dim cheminCle As String
    Dim lnb As Long
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    scmd = p_scmd
    lblVersion.Caption = p_version_KaliRP
    
    p_bool_ModeDebug = True
    
    ' utiliser le .ini local
    p_Mode_FctTrace = False
        
    p_Mode_FctTrace = IIf(UCase(SYS_GetIni("TRACES", "MODE_TRACE", p_nomini)) = "OUI", True, False)
    p_Mode_FctTrace = True
    
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
    p_SuperUser = 1
    
    If STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT) <> "" Then
        p_param_supplementaires = True
        NumForm = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMFORM)
        If NumForm = "" Then
            p_NumForm = 0
        Else
            p_NumForm = NumForm
        End If
        numutil = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMUTIL)
        If numutil = "" Then
            p_NumUtil = 0
        ElseIf numutil = "ROOT" Then
            p_NumUtil = p_SuperUser
            p_CodeUtil = "ROOT"
        Else
            p_NumUtil = numutil
        End If
        NumModele = STR_GetChamp(STR_GetChamp(scmd, ";", p_SCMD_PARAM_SUPPLEMENT), "|", p_SCMD_PARAM_NUMFILTRE)
        If NumModele = "" Then
            p_nummodele = 0
        Else
            p_nummodele = NumModele
        End If
    End If
    p_CheminRapportType_Ini = p_nomini
    If FICH_FichierExiste(p_CheminRapportType_Ini) Then
        ' Voir si le fichier .ini existe
        Me.MnuFichierIni.Caption = "Fichier Ini = " & p_CheminRapportType_Ini
        ' connexion à la base de données
        
        If p_nomBDD_ODBC = "" Or p_nomBDD_ODBC = "VIDE" Then
            ' 2- Nom de la base
            nom_bdd = SYS_GetIni("Base", "NOM", p_CheminRapportType_Ini)
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
        If Odbc_Init("PG", p_nomBDD_ODBC, True) = P_ERREUR Then
            MsgBox "Connexion à la base " & nom_bdd & " impossible"
            End
        End If
        
        P_MODCNUM_ou_MODELE = MODCNUM_ou_MODELE()
                
        p_estV4 = estV4()
        mode_Sites = Odbc_TableExiste("site")
        
        ' Récupérer l'adresse HTTP du serveur (champ pg_serveur)
        
        'p_AdrServeur = SYS_GetIni("Chemins", "HTTP_SERVEUR", p_CheminRapportType_Ini)
        sql = "select PG_SERVEUR from PRMGEN_HTTP"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger PG_SERVEUR à partir de la base(PRMGEN_HTTP)"
            End
        ElseIf rs.EOF Then
            MsgBox "Impossible de charger PG_SERVEUR à partir de la base(PRMGEN_HTTP EOF)"
            End
        Else
            p_AdrServeur = rs("pg_serveur")
            rs.Close
            If p_AdrServeur = "" Then
                MsgBox "Adresse HTTP du serveur non renseignée"
            End If
        End If
            
        ' Récupérer HTTP : maxparfichier
        sql = "select PG_HTTP_MAXPARFICHIER from PRMGEN_HTTP"
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
        If p_S_Vers_Conf <> "" Then
            Me.MnuSVersConf.Caption = "s_vers_conf = " & p_S_Vers_Conf
        End If
        
        ' Récupérer le Chemin de Dépot HTTP du serveur  (champ pg_http_chemindepot)
        sql = "select PG_HTTP_CHEMINDEPOT from PRMGEN_HTTP"
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
        sql = "select PG_CHEMINDOC from PRMGEN_HTTP"
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
                
        Call Odbc_RecupVal("Select pg_cheminkw from prmgen_http", p_cheminKW)
        
        If p_nomBDD_SERVEUR = "" Then
            ' 2- Nom de la base sur le serveur
            nom_bddS = SYS_GetIni("Base", "Nom_BDD_SERVEUR", p_CheminRapportType_Ini)
            If nom_bddS = "" Then
                Call MsgBox("Pas de nom de base (Serveur).", vbInformation + vbOKOnly)
                End
            Else
                Me.MnuBaseServeur.Caption = "Base Serveur = " & nom_bddS
                p_nomBDD_SERVEUR = nom_bddS
            End If
        End If
                
        ' 3- Chemin application : p_chemin_appli
        If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
            MnuCheminAppli.Caption = p_chemin_appli
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
        ' et le chemin des Résultats
        sql = "select pg_chemin_rapport, pg_http_rapport from prmgen_http"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            MsgBox "Impossible de charger pg_chemin_rapport à partir de la base"
            End
        End If
        p_Chemin_Modeles_Serveur = rs("pg_chemin_rapport").Value & "/Modeles_Serveur"
        If Not KF_EstRepertoire(p_Chemin_Modeles_Serveur, False) Then
            MsgBox "Chemin des Modèles " & p_Chemin_Modeles_Serveur & " introuvable"
            End
        End If
        p_Chemin_Résultats = rs("pg_chemin_rapport").Value & "/Resultats"
        If Not KF_EstRepertoire(p_Chemin_Résultats, False) Then
            MsgBox "Chemin des Résultats " & p_Chemin_Résultats & " introuvable"
            End
        End If
        Me.MnuCheminMod_Serveur.Caption = "Modèles => " & p_Chemin_Modeles_Serveur
        p_HTTP_Résultats = rs("pg_http_rapport").Value & "/Resultats"
        rs.Close
        
        ' les traces
        p_Mode_FctTrace = IIf(UCase(SYS_GetIni("TRACES", "MODE_TRACE", p_CheminRapportType_Ini)) = "OUI", True, False)
        p_Mode_FctTrace = True
        If p_Mode_FctTrace Then
            MnuTraceActive.Caption = "Désactiver les Traces"
        Else
            MnuTraceActive.Caption = "Activer les Traces"
        End If
        
        ' Récupérer le chemin des Modèles en local
        sCleTmp = "CLE_TMP"
        CleTmp = SYS_GetIni("Chemins", sCleTmp, p_CheminRapportType_Ini)
        NewCle = CleTmp
LabDosLocal:
        cheminCle = Rep_Documents(sCleTmp, CleTmp, NewCle)
        p_Chemin_Modeles_Local = cheminCle
        If NewCle = "" Or NewCle = "KaliDoc" Then
            NewCle = "KaliDoc"
            Call SYS_PutIni("Chemins", "CLE_TMP", NewCle, p_CheminRapportType_Ini)
        Else
            If CleTmp <> NewCle Then
                Call SYS_PutIni("Chemins", "CLE_TMP", NewCle, p_CheminRapportType_Ini)
            End If
            p_Chemin_Modeles_Local = Replace(UCase(p_Chemin_Modeles_Local), UCase(p_chemin_appli), cheminCle)
        End If
        p_Chemin_Modeles_Local = p_Chemin_Modeles_Local & "\KaliRP\Modeles_Local"
        p_CheminDossierTravailLocal = p_Chemin_Modeles_Local
        
        Me.MnuCheminLocalCle.Caption = sCleTmp & "=" & NewCle
        Me.MnuCheminMod_Local.Caption = "Modèles => " & p_Chemin_Modeles_Local
TestRep:
        If p_Chemin_Modeles_Local = "" Then
            MsgBox "Chemin des Modèles non renseigné"
            End
        ElseIf Not FICH_EstRepertoire(p_Chemin_Modeles_Local, False) Then
            Call FICH_CreerRepComp(p_Chemin_Modeles_Local, True, False)
            If Not FICH_EstRepertoire(p_Chemin_Modeles_Local, False) Then
                MsgBox "Impossible de créer le Dossier Local : " & p_Chemin_Modeles_Local
                CleTmp = "TEST"
                GoTo LabDosLocal
            End If
            GoTo TestRep
        End If
    End If

    p_Chemin_FichierTrace = p_Chemin_Modeles_Local & "\KaliRP_Trace.txt"
    
    If p_NumUtil = 0 Then
        ret = P_SaisirUtilIdent(10, 10, 10, 10)
        If ret <> P_OUI Then
            'MsgBox "Utilisateur non autorisé"
            End
        End If
    ElseIf p_NumUtil = p_SuperUser Then
    Else
        p_CodeUtil = AppelDirect(p_NumUtil)
    End If
    
    If p_CodeUtil <> "ROOT" Then
        sql = "select * from utilisateur where u_num = " & p_NumUtil
        If Odbc_SelectV(sql, rs) <> P_ERREUR Then
            If Not rs.EOF Then
                Me.Caption = Me.Caption & "   (" & left$(rs("u_prenom"), 1) & ". " & UCase(rs("u_nom")) & ")"
                rs.Close
            Else
                MsgBox "Utilisateur non autorisé"
                End
            End If
        End If
    Else
        Me.Caption = Me.Caption & "   (SuperUtilisateur)"
        Me.MnuFichierIni.Caption = "Ouvrir le .Ini"
    End If

    p_mode_acces = ""
    p_peut_creer = True
    sql = "select count(*) from fonction, fctok_util where fct_code='CR_KRP'" _
        & " and fu_unum=" & p_NumUtil _
         & " and fu_fctnum=fct_num"
    Call Odbc_Count(sql, lnb)
    If p_CodeUtil = "ROOT" Then
        lnb = 1
    End If
    If lnb = 0 Then
        p_peut_creer = False
        ComAssistant.Visible = False
        p_mode_acces = "SIMUL"
        sql = "select * from rapport_type where rp_user_admin like '%U" & p_NumUtil & "=%'"
        If p_NumUtil = p_SuperUser Then
            sql = "select * from rapport_type"
        End If
        If p_NumForm > 0 Then
            sql = sql & " and rp_formulaires like '%F" & p_NumForm & ";%'"
        End If
        Call Odbc_SelectV(sql, rs)
        If rs.EOF Then
            If p_NumForm > 0 Then
                Call MsgBox("Vous n'avez accès à aucun modèle de ce formulaire", vbOKOnly, "")
            Else
                Call MsgBox("Vous n'avez accès à aucun modèle", vbOKOnly, "")
            End If
            rs.Close
            End
        End If
        While Not rs.EOF
            n = STR_GetNbchamp(rs("rp_user_admin").Value, ";")
            For i = 0 To n - 1
                s = STR_GetChamp(rs("rp_user_admin").Value, ";", i)
                If InStr(s, "U" & p_NumUtil & "=") > 0 Then
                    If s <> "U" & p_NumUtil & "=:::SIMUL" Then
                        p_mode_acces = ""
                        Exit For
                    End If
                End If
            Next i
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    If p_mode_acces <> "SIMUL" Then
        Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    End If
    
    p_Bool_Modif_Resolution = False
    
    TesterResolutionEcran

    Call VerifSiVide
    
    NomFichierUser = p_Chemin_Modeles_Local & "/User_" & p_NumUtil & ".User"
    If Not FICH_FichierExiste(NomFichierUser) Then
        If Not FICH_OuvrirFichier(NomFichierUser, FICH_ECRITURE, fd) = P_ERREUR Then
            Print #fd, "0"
            Close #fd
        End If
    End If
    
    If p_mode_acces = "SIMUL" Then
LabEncore:
        FctTrace ("Initialiser appel ComOuvrirModele_Click")
        Call ComOuvrirModele_Click
        GoTo LabEncore
        Exit Sub
    End If
    
    If p_nummodele > 0 Then
EncoreModele:
        Set Frm = PiloteExcelBis
        bcr = PiloteExcelBis.AppelFrm(p_nummodele, p_NumForm, "P")
        Set Frm = Nothing
        If p_Appel_Création_Nouveau_Modele Then
            GoTo EncoreModele
        End If
        p_NumForm = 0
        'End
    End If
    If p_NumForm > 0 Then
EncoreFormulaire:
        Set Frm = PiloteExcelBis
        bcr = PiloteExcelBis.AppelFrm(p_nummodele, p_NumForm, "P")
        Set Frm = Nothing
        If p_Appel_Création_Nouveau_Modele Then
            GoTo EncoreFormulaire
        End If
    End If
    MnuCheminAppli.Caption = "Chemin_Appli = " & p_chemin_appli
End Sub

Private Sub MDIForm_Activate()
    
    If g_form_active Then
        Exit Sub
    End If
    
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub MDIForm_Load()

    g_form_active = False
    
    MenuG.MnuVersionExcel.tag = 0 ' indique que choix automatique
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If p_Bool_Modif_Resolution Then
        Call MetResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    End If
    
    On Error GoTo Err_Unload
    Exc_obj.Quit
    Set Exc_obj = Nothing
    

Err_Unload:
   Resume Next
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Text(0).Visible = False
    Me.Text(1).Visible = False
    Me.Text(2).Visible = False
    Me.Text(3).Visible = False
    Me.ComOuvrirModele.BackColor = Me.CmdBidon.BackColor
    Me.ComGénérer.BackColor = Me.CmdBidon.BackColor
    Me.ComResultats.BackColor = Me.CmdBidon.BackColor
    Me.ComAssistant.BackColor = Me.CmdBidon.BackColor
End Sub

Private Sub MnuAide_Click()
    Call Appel_Aide
    
End Sub

Private Sub MnuFichierIni_Click()
    FctOuvrirFichier (p_CheminRapportType_Ini)
End Sub


Private Sub MnuQuitter_Click()
   
    If p_Bool_Modif_Resolution Then
        Call MetResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    End If
   
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
    
    If MnuTraceActive.Caption = "Activer les Traces" Then
        p_Mode_FctTrace = True
        MnuTraceActive.Caption = "Désactiver les Traces"
        On Error Resume Next
        PiloteExcelBis.cmdTrace.Visible = True
        Exit Sub
    End If
    If MnuTraceActive.Caption = "Désactiver les Traces" Then
        p_Mode_FctTrace = False
        MnuTraceActive.Caption = "Activer les Traces"
        On Error Resume Next
        PiloteExcelBis.cmdTrace.Visible = False
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


Private Sub MnuVersionExcelCharts_Click()
    MenuG.MnuVersionExcelShapes.Caption = "supérieure à 2003"
    MenuG.MnuVersionExcelCharts.Caption = ">> jusqu'à 2003 (" & Exc_obj_Version & ")"
    p_VersionExcel = "2003"
    MnuVersionExcel.tag = 1 ' indique que choix manuel
    MnuVersionExcelCharts.Enabled = False
    MnuVersionExcelShapes.Enabled = True
    Call SYS_PutIni("EXCEL", "Version", p_VersionExcel, p_nomini)
End Sub

Private Sub MnuVersionExcelShapes_Click()
    MenuG.MnuVersionExcelShapes.Caption = ">> supérieure à 2003 (" & Exc_obj_Version & ")"
    MenuG.MnuVersionExcelCharts.Caption = "jusqu 'à 2003"
    p_VersionExcel = "2007"
    MnuVersionExcel.tag = 1 ' indique que choix manuel
    MnuVersionExcelCharts.Enabled = True
    MnuVersionExcelShapes.Enabled = False
    Call SYS_PutIni("EXCEL", "Version", p_VersionExcel, p_nomini)
End Sub
