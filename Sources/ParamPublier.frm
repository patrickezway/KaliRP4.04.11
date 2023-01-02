VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ParamPublier 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Publication des résultats"
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
      Height          =   8085
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10785
      Begin VB.Frame FrmHTTPD 
         BackColor       =   &H00C0C0C0&
         Height          =   1815
         Left            =   1440
         TabIndex        =   31
         Top             =   4920
         Visible         =   0   'False
         Width           =   8175
         Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
            Height          =   255
            Left            =   2880
            TabIndex        =   32
            Top             =   1320
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin ComctlLib.ProgressBar PgbarHTTPDTaille 
            Height          =   255
            Left            =   2880
            TabIndex        =   33
            Top             =   840
            Width           =   5175
            _ExtentX        =   9128
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
            TabIndex        =   36
            Top             =   240
            Width           =   7455
         End
         Begin VB.Label lblHTTPDTemps 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblHTTPDTaille 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Index           =   9
         Left            =   5400
         Picture         =   "ParamPublier.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Accès à l'aide"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Index           =   4
         Left            =   480
         Picture         =   "ParamPublier.frx":0359
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Paramétrage des documents à générer"
         Top             =   3120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
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
         Height          =   300
         Index           =   8
         Left            =   9960
         Picture         =   "ParamPublier.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Nouveau document"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   320
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   5
         Left            =   9960
         Picture         =   "ParamPublier.frx":0BDB
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer le document"
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   320
      End
      Begin VB.Frame FrmPublier 
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
         Height          =   4215
         Left            =   480
         TabIndex        =   17
         Top             =   3000
         Width           =   9975
         Begin VB.CheckBox ChkDocPublic 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tableau de Bord public"
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
            Left            =   3120
            TabIndex        =   29
            Top             =   1200
            Width           =   3735
         End
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
            Height          =   420
            Index           =   7
            Left            =   9480
            Picture         =   "ParamPublier.frx":1022
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter des destinataires"
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CheckBox ChkPublierKaliDoc 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Publier vers KaliDoc ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   21
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Dossier"
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
            Index           =   25
            Left            =   3120
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Nature"
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
            Index           =   21
            Left            =   4560
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Modèle"
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
            Index           =   20
            Left            =   7080
            TabIndex        =   18
            Top             =   360
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSFlexGridLib.MSFlexGrid grdDest 
            Height          =   2685
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Visible         =   0   'False
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   4736
            _Version        =   393216
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   8388608
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   16777215
            GridColorFixed  =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Destinataires"
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
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label LblDossier 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   9495
         End
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   2
         Left            =   5280
         Picture         =   "ParamPublier.frx":144D
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer le document"
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   320
      End
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
         Height          =   300
         Index           =   3
         Left            =   5280
         Picture         =   "ParamPublier.frx":1894
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Nouveau document"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   320
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Index           =   19
         Left            =   5400
         Picture         =   "ParamPublier.frx":1CEB
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Paramétrage des documents à générer"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin ComctlLib.ProgressBar PgBarChp 
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgBarGener 
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   7680
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgBarFeuille 
         Height          =   255
         Left            =   7080
         TabIndex        =   4
         Top             =   7680
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgBarDoc 
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   7680
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdFeuille 
         Height          =   1965
         Left            =   6000
         TabIndex        =   13
         Top             =   840
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   8388608
         BackColorFixed  =   8454143
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdDocument 
         Height          =   1965
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   8388608
         BackColorFixed  =   8454143
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Feuilles Excel"
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
         Index           =   4
         Left            =   6000
         TabIndex        =   16
         Top             =   600
         Width           =   3255
      End
      Begin ComctlLib.ImageList ImageListS 
         Left            =   9960
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   28
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   5
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ParamPublier.frx":2116
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ParamPublier.frx":2DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ParamPublier.frx":331A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ParamPublier.frx":38D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ParamPublier.frx":3E8E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documents publiés"
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
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblFeuille 
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
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Label lblDoc 
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   7080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   10785
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
         Picture         =   "ParamPublier.frx":45B8
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Enregistrer"
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
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
         Left            =   10080
         Picture         =   "ParamPublier.frx":4B21
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ParamPublier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_boolChkCliqué As Boolean

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_FERMER = 1
Private Const CMD_AJOUT_DEST = 7
Private Const CMD_ICONE_KALIDOC = 4
Private Const CMD_AJOUT_FEN = 8
Private Const CMD_AIDE = 9
Private Const CMD_SUPPR_FEN = 5
Private Const CMD_PARAM_PUBLIER = 19
Private Const CMD_AJOUT_DOC = 3
Private Const CMD_SUPPR_DOC = 2
Private Const CMD_SIMULATION_Un = 5
Private Const CMD_SIMULATION_Tous = 8
Private Const CMD_VOIR_RESULTATS = 9
Private Const CMD_CHOIX_DOSSIER = 25
Private Const CMD_VOIR_DIFFUSION = 11
Private Const CMD_FAIRE_DIFFUSION = 12
Private Const CMD_CHOIX_NATURE = 21
Private Const CMD_CHOIX_MODELE = 20

Private Const IMG_KALIDOC = 3
Private Const IMG_PAS_KALIDOC = 4

Private g_numfeuille As Integer
Private g_numDocument As Integer
Private g_strdest As String
Private g_bcr As Boolean
Private g_CheminModele As String
Private g_mode_saisie As Boolean
Private g_form_active As Boolean
Private g_DocParamDefaut As String

Private Const IMG_LOAD_EXCEL = 1
Private Const IMG_LOAD_HTML = 2

Private Faire_Doc_Click As Boolean

' pour le grid documents
Private Const ColGrdDocTitre = 0
Private Const ColGrdDocaFen = 1
Private Const ColGrdKaliDoc = 2
Private Const ColGrdDocLstFen = 3
Private Const ColGrdDocLstDest = 4
Private Const ColGrdDocExcel = 5
Private Const ColGrdDocCheminExcel = 6
Private Const ColGrdDocI_TabFichExcel = 7
Private Const ColGrdDocHTML = 8
Private Const ColGrdDocParam = 9
Private Const ColGrdDocCréé = 10
Private Const ColGrdNumDoc = 11
Private Const ColGrdDocChemin_Chemin = 12
Private Const ColGrdDocChemin_Nom = 13
Private Const ColGrdDocChemin_Extension = 14
Private Const ColGrdDocDocPublic = 15
Private Const ColGrdDocNum = 16

' pour le grid destinataires
Private Const ColGrdDestNum = 0
Private Const ColGrdDestLib = 1

' pour le grid feuilles
Private Const ColGrdFeuilNum = 0
Private Const ColGrdFeuilTag = 1
Private Const ColGrdFeuilaDoc = 2
Private Const ColGrdFeuilLib = 3

Dim p_BoolFaireChkClick As Boolean

' Tableau des fenetres
Private Type SFEN_EXCEL
    FenNum As Integer
    FenNom As String
    FenDest As String
End Type
Dim tbl_fen() As SFEN_EXCEL

' Tableau des diffusions
Private Type SDIFFUSION
    nomdoc As String
    CheminDoc As String
    NumDest As Integer
    nomdest As String
    numdoc As Integer
    Diffusé As Boolean
End Type
Dim tbl_diff() As SDIFFUSION

Public Function AppelFrm(ByRef v_bcr As Boolean, ByRef v_strdest As String, v_CheminModele As String) As String

    g_strdest = v_strdest
    g_bcr = v_bcr
    g_CheminModele = v_CheminModele
        
    Me.Show 1
    
    v_bcr = g_bcr
    
End Function

Private Sub RemplirTabFenetre()
    Dim I As Integer
    Dim prem As Integer
    Dim lig As Integer
    
    grdFeuille.Visible = True
    grdFeuille.Cols = 4
    grdFeuille.ColWidth(ColGrdFeuilNum) = 0
    grdFeuille.ColWidth(ColGrdFeuilTag) = 0
    grdFeuille.ColWidth(ColGrdFeuilaDoc) = 400
    grdFeuille.ColWidth(ColGrdFeuilLib) = grdFeuille.Width - 1300
    grdFeuille.Rows = 0
    prem = 0
    For I = 1 To Exc_obj.ActiveWorkbook.Sheets.Count
        ReDim Preserve tbl_fen(I) As SFEN_EXCEL
        p_bool_tbl_fenExcel = True
        tbl_fen(I).FenNum = I
        tbl_fen(I).FenNom = Exc_obj.ActiveWorkbook.Sheets(I).Name
        tbl_fen(I).FenDest = ""
        lig = I - 1
        If lig >= 0 Then
            grdFeuille.AddItem "", lig
        Else
            grdFeuille.AddItem ""
            lig = grdFeuille.Rows - 1
        End If
        grdFeuille.TextMatrix(lig, ColGrdFeuilNum) = I
        grdFeuille.TextMatrix(lig, ColGrdFeuilaDoc) = ""
        grdFeuille.TextMatrix(lig, ColGrdFeuilLib) = Exc_obj.ActiveWorkbook.Sheets(I).Name   'tbl_fen(v_i).FenNom
        grdFeuille.TextMatrix(lig, ColGrdFeuilTag) = ""
        'grdFeuille.col = grdFeuille.Cols - 1
        'grdFeuille.ColSel = grdFeuille.col
        'grdFeuille.RowSel = grdFeuille.Rows - 1
    Next I
End Sub


Private Sub initialiser()

    Dim NomFichierParam As String
    
    Erase p_tbl_FichExcelPublier()
    Erase tbl_diff()
    
    ' Charger le fichier
    ChargerFichier
    
    ' Remplir le tableau des fenetres
    ' ouvrir le fichier Excel
    Public_VerifOuvrir g_CheminModele & ".xls", False, False, p_tbl_FichExcelPublier
    
    RemplirTabFenetre
    
    cmd(CMD_ICONE_KALIDOC).Visible = False
    cmd(CMD_OK).Visible = False
    cmd(CMD_AJOUT_FEN).Visible = True
    cmd(CMD_PARAM_PUBLIER).Visible = False
    
    g_numDocument = -1
    
    Me.FrmPublier.Visible = False
    ' se mettre sur le premier document
    If grdDocument.Rows > 0 Then
        grdDocument.row = 0
        grddocument_click
    Else
        Me.lbl(2).Visible = False
        Me.cmd(CMD_AJOUT_DEST).Visible = False
    End If
    
    cmd(CMD_OK).Visible = False
    cmd(CMD_OK).Enabled = False
    
    Faire_Doc_Click = True
    
    p_BoolFaireChkClick = True
    
    g_mode_saisie = True

End Sub

Private Sub ChargerFichier()
    Dim s As String
    Dim fd As Integer
    Dim J As Integer
    Dim titreDoc As String
    Dim lstFen As String, numFen As String, strFen As String
    Dim lstdest As String, NumDest As Integer, typeDest As String
    Dim I As Integer, ligne As String
    Dim chemin As String, NomFichier As String, Extension As String
    Dim strDocChk As String
    Dim rs As rdoResultset
    
    J = 1
    grdDocument.Rows = 0
    grdDocument.Cols = 16
    grdDocument.ColWidth(ColGrdDocTitre) = 3800
    grdDocument.ColWidth(ColGrdDocaFen) = 500
    grdDocument.ColWidth(ColGrdKaliDoc) = 500
    grdDocument.ColWidth(ColGrdDocLstFen) = 0   '2000
    grdDocument.ColWidth(ColGrdDocLstDest) = 0  '2000
    grdDocument.ColWidth(ColGrdDocExcel) = 0    '500
    grdDocument.ColWidth(ColGrdDocHTML) = 0 '500
    grdDocument.ColWidth(ColGrdDocCheminExcel) = 0
    grdDocument.ColWidth(ColGrdDocI_TabFichExcel) = 0
    grdDocument.ColWidth(ColGrdDocParam) = 0
    grdDocument.ColWidth(ColGrdNumDoc) = 0
    grdDocument.ColWidth(ColGrdDocCréé) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Chemin) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Nom) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Extension) = 0
    grdDocument.ColWidth(ColGrdDocDocPublic) = 0
    
    sql = "select * from rp_dossier where rp_num=" & titreDoc
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
            titreDoc = tbl_fichExcel(I).CmdTitreDoc
            lstFen = tbl_fichExcel(I).CmdLstFen
            ' lstdest = tbl_fichExcel(i).CmdLstDest
            lstdest = STR_GetChamp(tbl_fichExcel(I).CmdLstDest, "%", 0)
            grdDocument.AddItem titreDoc & vbTab & "" & vbTab & "" & vbTab & lstFen & vbTab & lstdest & vbTab & "" & vbTab & "" & vbTab & I & vbTab & "" & vbTab & tbl_fichExcel(I).CmdMenFormeDoc & vbTab & "N"
            grdDocument.TextMatrix(grdDocument.Rows - 1, ColGrdDocDocPublic) = STR_GetChamp(tbl_fichExcel(I).CmdLstDest, "%", 1)
            strDocChk = STR_GetChamp(tbl_fichExcel(I).CmdMenFormeDoc, ";", 0)
            grdDocument.row = grdDocument.Rows - 1
            grdDocument.col = ColGrdKaliDoc
            If strDocChk = "1" Then
                Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
            Else
                Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            End If
            If tbl_fichExcel(I).CmdMenFormeDoc <> "" Then
                g_DocParamDefaut = tbl_fichExcel(I).CmdMenFormeDoc
            End If
            grdDocument.RowHeight(grdDocument.Rows - 1) = 400
            J = J + 1
        End If
    Wend
    rs.Close
    
End Sub

Private Sub ChargerFichier_old()
    Dim s As String
    Dim fd As Integer
    Dim J As Integer
    Dim titreDoc As String
    Dim lstFen As String, numFen As String, strFen As String
    Dim lstdest As String, NumDest As Integer, typeDest As String
    Dim I As Integer, ligne As String
    Dim chemin As String, NomFichier As String, Extension As String
    Dim strDocChk As String
    
    J = 1
    grdDocument.Rows = 0
    grdDocument.Cols = 16
    grdDocument.ColWidth(ColGrdDocTitre) = 3800
    grdDocument.ColWidth(ColGrdDocaFen) = 500
    grdDocument.ColWidth(ColGrdKaliDoc) = 500
    grdDocument.ColWidth(ColGrdDocLstFen) = 0   '2000
    grdDocument.ColWidth(ColGrdDocLstDest) = 0  '2000
    grdDocument.ColWidth(ColGrdDocExcel) = 0    '500
    grdDocument.ColWidth(ColGrdDocHTML) = 0 '500
    grdDocument.ColWidth(ColGrdDocCheminExcel) = 0
    grdDocument.ColWidth(ColGrdDocI_TabFichExcel) = 0
    grdDocument.ColWidth(ColGrdDocParam) = 0
    grdDocument.ColWidth(ColGrdNumDoc) = 0
    grdDocument.ColWidth(ColGrdDocCréé) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Chemin) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Nom) = 0
    grdDocument.ColWidth(ColGrdDocChemin_Extension) = 0
    grdDocument.ColWidth(ColGrdDocDocPublic) = 0
    
    On Error GoTo Err_Tab
    For I = 0 To UBound(tbl_fichExcel)
        If tbl_fichExcel(I).CmdType = "DOC" Then
            titreDoc = tbl_fichExcel(I).CmdTitreDoc
            lstFen = tbl_fichExcel(I).CmdLstFen
            ' lstdest = tbl_fichExcel(i).CmdLstDest
            lstdest = STR_GetChamp(tbl_fichExcel(I).CmdLstDest, "%", 0)
            grdDocument.AddItem titreDoc & vbTab & "" & vbTab & "" & vbTab & lstFen & vbTab & lstdest & vbTab & "" & vbTab & "" & vbTab & I & vbTab & "" & vbTab & tbl_fichExcel(I).CmdMenFormeDoc & vbTab & "N"
            grdDocument.TextMatrix(grdDocument.Rows - 1, ColGrdDocDocPublic) = STR_GetChamp(tbl_fichExcel(I).CmdLstDest, "%", 1)
            strDocChk = STR_GetChamp(tbl_fichExcel(I).CmdMenFormeDoc, ";", 0)
            grdDocument.row = grdDocument.Rows - 1
            grdDocument.col = ColGrdKaliDoc
            If strDocChk = "1" Then
                Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
            Else
                Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            End If
            If tbl_fichExcel(I).CmdMenFormeDoc <> "" Then
                g_DocParamDefaut = tbl_fichExcel(I).CmdMenFormeDoc
            End If
            grdDocument.RowHeight(grdDocument.Rows - 1) = 400
            J = J + 1
        End If
    Next I
Err_Tab:
    On Error GoTo 0
    
    'cmd(CMD_SIMULATION_Un).Visible = False
    'cmd(CMD_SIMULATION_Tous).Visible = False
    Exit Sub
    
End Sub


Private Sub OLDOuvrirModele(ByVal v_chemin As String, ByVal v_visible As Boolean, ByVal v_àSauver As Boolean)
    Dim encore As Boolean
    Dim retour As Integer
    Dim FichierIn As String, cmd As String
    Dim v_chemin_For As String, v_chemin_Fil As String, v_chemin_Excel As String
    ' Ouvrir le modele
'      Chemin_Parametrage = v_chemin_Excel _
'                         & "For" & tbParam(Ind_Numfor) & "\" _
'                         & "Fil" & tbParam(Ind_Special_Pere) & "\" _
'                         & "Valeurs\Param.xls"
      
    ' Ouvrir le modele
      If FICH_FichierExiste(v_chemin) Then
         If Excel_Init() = P_OK Then
            Excel_OuvrirDoc v_chemin, "", Exc_wrk, False
            Call Public_FichiersExcelOuverts(p_tbl_FichExcelPublier(), "Voir", v_chemin, v_visible, v_àSauver)
            Exc_obj.Visible = v_visible
         End If
      End If
End Sub


Private Function quitter() As Boolean

    Dim reponse As Integer
    Dim LaUbound As Integer
    Dim I As Integer
    Dim J As Integer
    
    If cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then
            quitter = False
            Exit Function
        End If
    End If
    
    ' voir si des fichiers Excel sont à fermer proprement
    LaUbound = 0
    On Error GoTo Faire
    LaUbound = UBound(p_tbl_FichExcelPublier(), 1) + 1
    For I = 0 To LaUbound - 1
        If p_tbl_FichExcelPublier(I).FichàSauver Then
            For J = 1 To Exc_obj.Workbooks.Count
                If UCase(Exc_obj.Workbooks(J).FullName) = UCase(p_tbl_FichExcelPublier(I).FichFullname) Then
                    Exc_obj.Workbooks(J).Close True
                Else
                    Exc_obj.Workbooks(J).Close False
                End If
            Next J
        End If
    Next I
Faire:
    
    If Exc_obj.Workbooks.Count = 0 Then
        Exc_obj.Quit
        Set Exc_obj = Nothing
    End If
    
    Unload Me
    
    quitter = True
    
End Function


Private Sub valider()

    Dim spm As String
    Dim I As Integer
    Dim J As Integer
    Dim num As Integer
    Dim Bound As Integer
    Dim sT As String
    Dim sF As String, sD As String
    Dim sMenF As String
    Dim titreDoc As String
    Dim ind As Integer, deja As Boolean
    Dim indDoc As Integer
    Dim s As String
    
    g_bcr = True
    ' Reconstituer les lignes de doc
    For I = 0 To UBound(tbl_fichExcel)
        If tbl_fichExcel(I).CmdType = "DOC" Then
            tbl_fichExcel(I).CmdType = "LIBRE"
        End If
    Next I
    
    For I = 0 To grdDocument.Rows - 1
        ind = -1
        sF = grdDocument.TextMatrix(I, ColGrdDocLstFen)
        sD = grdDocument.TextMatrix(I, ColGrdDocLstDest)
        sMenF = grdDocument.TextMatrix(I, ColGrdDocParam)
        titreDoc = grdDocument.TextMatrix(I, ColGrdDocTitre)
        If sD & sF <> "" Then
            For J = 0 To UBound(tbl_fichExcel)
                If tbl_fichExcel(J).CmdType = "LIBRE" Then
                    ind = J
                    Exit For
                End If
            Next J
            If ind = -1 Then
                ' en ajouter un
                Bound = UBound(tbl_fichExcel()) + 1
                ReDim Preserve tbl_fichExcel(Bound) As SFICH_PARAM_EXCEL
                p_bool_tbl_fichExcel = True
                tbl_fichExcel(Bound).CmdType = "DOC"
                tbl_fichExcel(Bound).CmdLstFen = sF
                tbl_fichExcel(Bound).CmdLstDest = sD & "%" & grdDocument.TextMatrix(I, ColGrdDocDocPublic)
                'tbl_fichExcel(Bound).CmdLstDest = tbl_fichExcel(ind).CmdLstDest
                tbl_fichExcel(Bound).CmdTitreDoc = titreDoc
                tbl_fichExcel(Bound).CmdMenFormeDoc = grdDocument.TextMatrix(I, ColGrdDocParam)
            Else
                ' remplacer le ind
                tbl_fichExcel(ind).CmdType = "DOC"
                tbl_fichExcel(ind).CmdLstFen = sF
                tbl_fichExcel(ind).CmdLstDest = sD & "%" & grdDocument.TextMatrix(I, ColGrdDocDocPublic)
                'tbl_fichExcel(ind).CmdLstDest = tbl_fichExcel(ind).CmdLstDest
                tbl_fichExcel(ind).CmdTitreDoc = titreDoc
                tbl_fichExcel(ind).CmdMenFormeDoc = grdDocument.TextMatrix(I, ColGrdDocParam)
            End If
        End If
    Next I
    
    Unload Me
    
End Sub

Private Sub ChkDocPublic_Click()
    If p_BoolFaireChkClick Then
        grdDocument.TextMatrix(grdDocument.row, ColGrdDocDocPublic) = ChkDocPublic.Value
    
        cmd(CMD_OK).Visible = True
        cmd(CMD_OK).Enabled = True
    End If
End Sub

Private Sub ChkPublierKaliDoc_Click()
    
    Dim strDocChk As String, strDocNum As String
    Dim strDocNature As String, strDocModele As String
    Dim FaiteAuto As Boolean
    Dim I As Integer
    
    If p_BoolFaireChkClick Then
        If Me.ChkPublierKaliDoc.Value = 1 Then
            If grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam) = "" Or grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam) = ";" Then
                strDocChk = "1"
                strDocNum = STR_GetChamp(g_DocParamDefaut, ";", 1)
                strDocNature = STR_GetChamp(g_DocParamDefaut, ";", 2)
                strDocModele = STR_GetChamp(g_DocParamDefaut, ";", 3)
                grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam) = strDocChk & ";" & strDocNum & ";" & strDocNature & ";" & strDocModele
                FaiteAuto = True
            End If
            
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
            Set cmd(CMD_ICONE_KALIDOC).Picture = ImageListS.ListImages(IMG_KALIDOC).Picture
            
        Else
            grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam) = ";"
            
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            Set cmd(CMD_ICONE_KALIDOC).Picture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            
        End If
        AfficheDest "D", grdDocument.row
        cmd(CMD_AJOUT_DEST).Visible = True
        If p_boolChkCliqué Then
            cmd(CMD_OK).Visible = True
            cmd(CMD_OK).Enabled = True
        End If
    End If
    p_BoolFaireChkClick = False
    Call MenF_Dossier
    p_BoolFaireChkClick = True
    If FaiteAuto And strDocNum = "" Then
        strDocNum = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 1)
        If strDocNum = "" Then
            cmd_Click (CMD_CHOIX_DOSSIER)
            strDocNum = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 1)
            If strDocNum <> "" Then
                strDocNature = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 2)
                If strDocNature = "" Then
                    cmd_Click (CMD_CHOIX_NATURE)
                    strDocNature = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 2)
                    If strDocNature <> "" Then
                        strDocModele = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 3)
                        If strDocModele = "" Then
                            cmd_Click (CMD_CHOIX_MODELE)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub VoirListeDiffusion()
    Dim lstdest As String, v_i As Integer
    Dim s As String
    Dim n As Integer
    Dim I As Integer
    Dim lib As String
    Dim srvnum As Long, srvnom As String
    Dim fctnum As Long, fctlibelle As String
    Dim prenom As String, nomutil As String, actif As Boolean
    Dim numutil As String, nomdoc As String
    Dim iRow As Integer
    Dim sql As String, rs As rdoResultset
    Dim UboundDiff As Integer, ii As Integer
    Dim NomFichier As String
    Dim nbdiff As Integer
    Dim nbrestediff As Integer
    Dim idiffuser As Integer
    Dim cheminRep As String
    Dim J As Integer
    Dim jj As Integer
    Dim strDocChk As String, strDocNum As String
    Dim strDocNature As String, strDocModele As String
    Dim op As String
    Dim numdos As String
    Dim DocàCréer As Boolean
    Dim iNumDos As Integer
    Dim iNumDoc As Long
    Dim nbligne As Integer
    Dim StrGrdDocParam As String
    Dim fd As Integer
    
    ' les destinataires sont dans tbl_diff
    'v_i = grdDocument.RowSel
    v_i = g_numDocument
    idiffuser = 2
Début:
    Call CL_Init
    nbligne = 0
    On Error Resume Next
    UboundDiff = 0
    UboundDiff = UBound(tbl_diff())
    For ii = 0 To UboundDiff
        If tbl_diff(ii).numdoc = v_i Then
            If Not tbl_diff(ii).Diffusé Then
                If idiffuser = 1 Or idiffuser = 2 Then
                    Call CL_AddLigne(tbl_diff(ii).nomdoc & " pour " & tbl_diff(ii).nomdest, tbl_diff(ii).NumDest, ii, True)
                    nbligne = nbligne + 1
                End If
            Else
                If idiffuser = 0 Or idiffuser = 2 Then
                    Call CL_AddLigne("(diffusion faite) >> " & tbl_diff(ii).nomdoc & " pour " & tbl_diff(ii).nomdest, tbl_diff(ii).NumDest, ii, False)
                    nbligne = nbligne + 1
                End If
            End If
        End If
    Next ii
    If nbligne = 0 Then
        MsgBox "Liste vide"
        If idiffuser = 2 Then
            Exit Sub
        End If
        idiffuser = 2
        GoTo Début
    End If
'    grdDest(v_i).Visible = True
'    n = STR_GetNbchamp(lstdest, ";")
'    For i = 0 To n - 1
'        s = STR_GetChamp(lstdest, ";", i)
'        If left(s, 1) = "U" Then
'            If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & Replace(s, "U", ""), _
'                                prenom, nomutil, actif) = P_ERREUR Then
'                s = "U0"
'                lib = "Utilisateur ???"
'            Else
'                lib = prenom & "." & nomutil
'            End If
'        ElseIf left(s, 1) = "F" Then
'            If Odbc_RecupVal("select FT_Num, FT_Libelle from fcttrav where FT_Num=" & Replace(s, "F", ""), _
'                                fctnum, fctlibelle) = P_ERREUR Then
'                s = "F0"
'                lib = "Fonction ???"
'            Else
'                lib = "Fonction : " & fctlibelle
'            End If
'        ElseIf left(s, 1) = "S" Then
'            If Odbc_RecupVal("select SRV_Num,SRV_Nom from service where SRV_Num=" & Replace(s, "S", ""), _
'                                 srvnum, srvnom) = P_ERREUR Then
'                s = "S0"
'                lib = "Service ???"
'            Else
'                lib = "Service : " & srvnom
'            End If
'        End If
        
'        Call CL_AddLigne(lib, i + 1, s, True)
        
'    Next i
        
    Call CL_InitTitreHelp("Liste de diffusion", "")
    Call CL_AddBouton("déjà diffusés", "", 0, 0, 2000)
    Call CL_AddBouton("non diffusés", "", 0, 0, 2000)
    Call CL_AddBouton("Tous", "", 0, 0, 2000)
    Call CL_AddBouton("Diffuser", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitTaille(0, -15)
    
    ChoixListe.Show 1
        
    If CL_liste.retour = 3 Then
        ' Diffuser
        strDocChk = STR_GetChamp(grdDocument.TextMatrix(v_i, ColGrdDocParam), ";", 0)
        strDocNum = STR_GetChamp(grdDocument.TextMatrix(v_i, ColGrdDocParam), ";", 1)
        iNumDos = val(strDocNum)

        If iNumDos = 0 Then
            MsgBox "Vous devez choisir un dossier pour la publication de ce document"
            Exit Sub
        End If
        
        nbdiff = 0
        nbrestediff = 0
        
        'Ouvrir le fichier de diffusion
        grdDocument.row = g_numDocument
        NomFichier = grdDocument.TextMatrix(I, ColGrdDocTitre)
        NomFichier = Replace(NomFichier, " ", "_")
        g_CheminModele = Replace(g_CheminModele, "/", "\")
        n = STR_GetNbchamp(g_CheminModele, "\")
        cheminRep = STR_GetChamp(g_CheminModele, "\", n - 1) & "\" & Replace(Date, "/", "_")
        If Not FICH_EstRepertoire(p_Chemin_Résultats & "\" & cheminRep, False) Then
            MsgBox "Le dossier " & p_Chemin_Résultats & "\" & cheminRep & " a été créé"
        End If
        Call FICH_OuvrirFichier(p_Chemin_Résultats & "\" & cheminRep & "\" & NomFichier & ".Diffusion", FICH_ECRITURE, fd)
        For I = 0 To UBound(CL_liste.lignes)
            If CL_liste.lignes(I).selected Then
                ii = CL_liste.lignes(I).tag
                ' voir dans le tableau
                If Not tbl_diff(ii).Diffusé Then
                    ' Voir si le document est à créer
                    If grdDocument.TextMatrix(v_i, ColGrdDocCréé) <> "O" Then
                        DocàCréer = True
                    Else
                        DocàCréer = False
                    End If
                    ' fabriquer la liste de diffusion pour ce document
                    lstdest = ""
                    op = ""
                    For J = 0 To UBound(CL_liste.lignes)
                        If CL_liste.lignes(J).selected Then
                            For jj = 0 To UBound(tbl_diff)
                                If tbl_diff(jj).nomdoc <> "" Then
                                    If tbl_diff(jj).nomdoc = tbl_diff(ii).nomdoc Then
                                        If tbl_diff(jj).NumDest = CL_liste.lignes(J).num Then
                                            lstdest = lstdest & op & "U" & tbl_diff(jj).NumDest
                                            tbl_diff(jj).Diffusé = True
                                            Print #fd, tbl_diff(jj).nomdest
                                            op = "|"
                                        End If
                                    End If
                                End If
                            Next jj
                        End If
                    Next J
                    nbdiff = nbdiff + 1
                    
                    'iRow = grdDocument.row
                    iRow = g_numDocument
                    strDocChk = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 0)
                    strDocNum = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 1)
                    strDocNature = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 2)
                    strDocModele = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 3)
                    'iNumDoc = CreerDoc(iNumDos, DocàCréer, strDocNature, strDocModele, grdDocument.TextMatrix(iRow, ColGrdDocTitre), grdDocument.TextMatrix(v_i, ColGrdDocCheminExcel), lstdest)
                    
                    If DocàCréer And iNumDoc = 0 Then
                        MsgBox "Document non créé ! Aucune diffusion effectuée"
                        Exit Sub
                    End If
                    
                    ' indique qu'il est créé
                    grdDocument.TextMatrix(v_i, ColGrdDocCréé) = "O"
                    grdDocument.TextMatrix(v_i, ColGrdNumDoc) = iNumDoc
                    tbl_diff(ii).Diffusé = True
                End If
            Else
                ii = CL_liste.lignes(I).tag
                If Not tbl_diff(ii).Diffusé Then
                    nbrestediff = nbrestediff + 1
                End If
            End If
        Next I
        Close #fd
        If nbdiff = 0 Then
            MsgBox "Pas de diffusion à faire"
        Else
            MsgBox nbdiff & " diffusion(s) effectuée(s)"
        End If
        'If nbrestediff = 0 Then
        '    Cmd(CMD_VOIR_DIFFUSION).Visible = False
        'End If
        Exit Sub
    ElseIf CL_liste.retour = 0 Or CL_liste.retour = 1 Or CL_liste.retour = 2 Then
        idiffuser = CL_liste.retour
        GoTo Début
    ElseIf CL_liste.retour = 4 Then
        Exit Sub
    End If
    '
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim frm As Form
    Dim iRow As Integer
    Dim s_bcr As String
    Dim retDos As Long
    Dim sql As String, rs As rdoResultset
    Dim strDocChk As String, strDocNum As String
    Dim strDocNature As String
    Dim strDocModele As String
    
    Select Case Index
    Case CMD_AIDE
        Call Appel_Aide
    Case CMD_ICONE_KALIDOC
        Me.FrmPublier.Visible = Not Me.FrmPublier.Visible = True
    Case CMD_AJOUT_FEN
        Call Ajout_FenDoc
    Case CMD_CHOIX_NATURE
        Call ChoixNature
        'Me.cmd(CMD_CHOIX_NATURE).Caption = "Nature"
    Case CMD_CHOIX_MODELE
        Call ChoixModele
        'Me.cmd(CMD_CHOIX_MODELE).Caption = "Modèle"
    Case CMD_VOIR_DIFFUSION
        Call VoirListeDiffusion
    Case CMD_CHOIX_DOSSIER
        strDocChk = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 0)
        strDocNum = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 1)
        strDocNature = ""
        strDocModele = ""
        
        iRow = grdDocument.row
        grdDocument.TextMatrix(iRow, ColGrdDocParam) = strDocChk & ";" & strDocNum & ";" & strDocNature & ";" & strDocModele
        Call MenF_Dossier
        
        retDos = ChoisirDocKalidoc(val(strDocNum))
        If retDos > 0 Then
            sql = "select * from Dossier where Ds_num = " & retDos
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            Me.LblDossier.Caption = "Choix du Dossier"   'ChercherNomDossier(retDos)
            If Not rs.EOF Then
                Me.LblDossier.Caption = "Dossier"   'ChercherNomDossier(retDos)
                Me.cmd(CMD_CHOIX_DOSSIER).tag = retDos
                grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam) = strDocChk & ";" & retDos
                cmd(CMD_OK).Visible = True
                cmd(CMD_OK).Enabled = True
            End If
            Call MenF_Dossier
    
            'strDocNature = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 2)
            'If strDocNature = "" Then
            '    cmd_Click (CMD_CHOIX_NATURE)
            '    strDocModele = STR_GetChamp(grdDocument.TextMatrix(grdDocument.row, ColGrdDocParam), ";", 3)
            '    If strDocModele = "" Then
            '        cmd_Click (CMD_CHOIX_MODELE)
            '    End If
            'End If
        End If
    Case CMD_AJOUT_DOC
        Call Ajout_Doc
    Case CMD_SUPPR_DOC
    'Case CMD_AJOUT_FENDOC
    '    Call Ajout_FenDoc
    Case CMD_AJOUT_DEST
        Call PrmDest
    Case CMD_OK
        Call valider
    Case CMD_FERMER
        Call quitter
    Case CMD_PARAM_PUBLIER
        'Me.FrmPublier.Visible = Not Me.FrmPublier.Visible = True
    End Select
End Sub

Private Function ChoixModele()
    Dim sql As String, modele As String
    Dim I As Integer, n As Integer
    Dim rs As rdoResultset
    Dim numdocs As Long
    Dim iRow As Integer
    Dim strDocChk As String
    Dim strDocNum As String
    Dim strDocNature As String
    Dim strDocModele As String
    
    If Me.cmd(CMD_CHOIX_DOSSIER).tag = "" Then
        Call MsgBox("Veuillez d'abord choisir un dossier.", vbOKOnly + vbInformation, "")
        ChoixModele = P_NON
        Exit Function
    End If
    
    If Me.cmd(CMD_CHOIX_NATURE).tag = "" Then
        Call MsgBox("Veuillez d'abord choisir une nature.", vbOKOnly + vbInformation, "")
        ChoixModele = P_NON
        Exit Function
    End If
    
    Call CL_Init
    
    sql = "select ds_donum from Dossier" _
        & " where DS_Num = " & Me.cmd(CMD_CHOIX_DOSSIER).tag
    If Odbc_RecupVal(sql, numdocs) = P_ERREUR Then
        ChoixModele = P_ERREUR
        Exit Function
    End If
    
    sql = "select distinct(DONM_Modele) from DocsNatureModele" _
        & " where DONM_DONum =" & p_NumDocs _
        & " and DONM_NDNum =" & Me.cmd(CMD_CHOIX_NATURE).tag _
        & " order by DONM_Modele"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ChoixModele = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("DONM_Modele").Value, n, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        Call MsgBox("Aucun modèle n'a été trouvé.", vbInformation + vbOKOnly, "")
        ChoixModele = P_NON
        Exit Function
    End If
        
    Call CL_InitTitreHelp("Choix d'un modèle", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    'Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        ChoixModele = P_NON
        Exit Function
    End If
    
    modele = CL_liste.lignes(CL_liste.pointeur).texte
    Me.cmd(CMD_CHOIX_MODELE).tag = modele

    iRow = grdDocument.row
    strDocChk = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 0)
    strDocNum = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 1)
    strDocNature = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 2)
    strDocModele = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 3)
    grdDocument.TextMatrix(iRow, ColGrdDocParam) = strDocChk & ";" & strDocNum & ";" & strDocNature & ";" & modele
    Call MenF_Dossier
    cmd(CMD_OK).Visible = True
    Me.cmd(CMD_OK).Enabled = True
    
    ChoixModele = P_OUI

End Function
Private Function ChoixNature()

    Dim sql As String
    Dim I As Integer, n As Integer
    Dim numnat As Long
    Dim rs As rdoResultset
    Dim numdocs As Long
    Dim iRow As Integer
    Dim strDocChk As String
    Dim strDocNum As String
    Dim strDocNature As String
    Dim strDocModele As String
    
    Call CL_Init
    
    If Me.cmd(CMD_CHOIX_DOSSIER).tag = "" Then
        Call MsgBox("Veuillez d'abord choisir un dossier.", vbOKOnly + vbInformation, "")
        ChoixNature = P_NON
        Exit Function
    End If
    
    sql = "select ds_donum from Dossier" _
        & " where DS_Num = " & Me.cmd(CMD_CHOIX_DOSSIER).tag
    If Odbc_RecupVal(sql, numdocs) = P_ERREUR Then
        ChoixNature = P_ERREUR
        Exit Function
    End If
    
    sql = "select * from NatureDoc" _
        & " where ND_Num in (select DON_NDNum from DocsNature where DON_DONum=" & numdocs & ")" _
        & " order by ND_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ChoixNature = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("ND_Nom").Value, rs("ND_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        MsgBox "Aucune nature n'a été trouvée.", vbExclamation + vbOKOnly, ""
        ChoixNature = P_NON
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Choix d'une nature", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        ChoixNature = P_NON
        Exit Function
    End If
    
    numnat = CL_liste.lignes(CL_liste.pointeur).num
    Me.cmd(CMD_CHOIX_NATURE).tag = numnat
    
    iRow = grdDocument.row
    strDocChk = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 0)
    strDocNum = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 1)
    strDocNature = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 2)
    strDocModele = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 3)
    grdDocument.TextMatrix(iRow, ColGrdDocParam) = strDocChk & ";" & strDocNum & ";" & numnat & ";" & strDocModele
    Call MenF_Dossier
    cmd(CMD_OK).Visible = True
    Me.cmd(CMD_OK).Enabled = True
    
    ChoixNature = P_OUI
    
End Function

Private Function ChercherNomNature(ByVal v_numNature)
    Dim sql As String, rs As rdoResultset
    
    If v_numNature <> "" Then
        sql = "select * from NatureDoc where ND_num = " & v_numNature
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            ChercherNomNature = ""
            Exit Function
        End If
        If Not rs.EOF Then
            ChercherNomNature = "Nature : " & rs("ND_Nom")
            Me.cmd(CMD_CHOIX_NATURE).tag = v_numNature
            Me.cmd(CMD_CHOIX_NATURE).Caption = "Nature"
        Else
            ChercherNomNature = ""
            Me.cmd(CMD_CHOIX_NATURE).tag = ""
            Me.cmd(CMD_CHOIX_NATURE).Caption = "Choix de la Nature"
        End If
        rs.Close
    End If
End Function

Private Function ChercherNomModele(ByVal v_NumModele)
    Dim sql As String, rs As rdoResultset
    
    If v_NumModele <> "" Then
        ChercherNomModele = "Modèle : " & v_NumModele
        Me.cmd(CMD_CHOIX_MODELE).Caption = "Modèle"
    Else
        ChercherNomModele = ""
        Me.cmd(CMD_CHOIX_MODELE).Caption = "Choix du Modèle"
    End If
End Function

Private Function ChercherNomDossier(ByVal v_numdos)
    Dim sql As String, rs As rdoResultset
    
    If v_numdos <> "" Then
        sql = "select * from Dossier where Ds_num = " & v_numdos
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            ChercherNomDossier = ""
            Exit Function
        End If
        If Not rs.EOF Then
            ChercherNomDossier = "Dossier : " & rs("Ds_Titre")
            Me.cmd(CMD_CHOIX_DOSSIER).Caption = "Dossier"
        Else
            ChercherNomDossier = ""
            Me.cmd(CMD_CHOIX_DOSSIER).Caption = "Choix du Dossier"
        End If
        rs.Close
    End If
End Function



Private Function ChoisirDocKalidoc(ByVal v_numdos As Long) As Long

    Dim sret As Integer
    Dim frm As Form
    
    Set frm = ChoixDossier
    sret = ChoixDossier.AppelFrm(0, 0)
    Set frm = Nothing
    ChoisirDocKalidoc = sret
End Function

Private Function UtilEstSuperviseur(ByVal v_numutil As Long, _
                                     ByVal v_numdocs As Long, _
                                     ByRef r_idroit As Integer) As Boolean

    Dim sql As String, slsts As String, s As String
    Dim I As Integer, n As Integer
    
    sql = "select DO_LstSuperv from Documentation" _
        & " where DO_Num=" & v_numdocs
    If Odbc_RecupVal(sql, slsts) = P_ERREUR Then
        UtilEstSuperviseur = False
        Exit Function
    End If
    
    n = STR_GetNbchamp(slsts, "|")
    For I = 0 To n - 1
        s = STR_GetChamp(slsts, "|", I)
        ' L'utilisateur est superviseur
        If CLng(Mid$(STR_GetChamp(s, ";", 0), 2)) = v_numutil Then
            r_idroit = STR_GetChamp(s, ";", 2)
            UtilEstSuperviseur = True
            Exit Function
        End If
    Next I
    
    UtilEstSuperviseur = False

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


Private Sub Ajout_Doc()
    Dim NbLig As Integer
    Dim stitre As String
    
    NbLig = grdDocument.Rows
    stitre = InputBox("Indiquez le titre du Document", "Titre du Document", "Document " & NbLig + 1)
    If stitre <> "" Then
        grdDocument.AddItem stitre & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
        g_numDocument = grdDocument.Rows
        grdDocument.row = grdDocument.Rows - 1
        grddocument_click
        ' se mettre en ajout de fenetres
        Ajout_FenDoc
    End If
End Sub

Private Sub Suppr_FenDoc()
    Dim lstFen As String
    Dim I As Integer
    
    ' Ajouter cette fenetre dans le document
    MsgBox g_numfeuille
    MsgBox g_numDocument
    lstFen = grdDocument.TextMatrix(g_numDocument, ColGrdDocLstFen)
    lstFen = Replace(lstFen, "F" & g_numfeuille & ";", "")
    lstFen = Replace(lstFen, ";;", ";")
    grdDocument.TextMatrix(g_numDocument, ColGrdDocLstFen) = lstFen
    ' mettre la fenetre en non surbrillance
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.col = ColGrdFeuilLib
        grdFeuille.row = I
        If grdFeuille.row = g_numfeuille - 1 Then
            grdFeuille.CellFontBold = False
        End If
    Next I
    grdFeuille_Click
    ' mettre enregistrer en visible
    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = True
End Sub

Private Sub Ajout_FenDoc()
    Dim lstFen As String

    Dim sql As String, sret As String, sfct As String
    Dim n As Integer, I As Integer
    Dim leIndex As Integer
    Dim rs As rdoResultset
    Dim laDim As Integer
    Dim numfor As Integer
    Dim stitre As String
    Dim NumFiltre As Integer
    Dim Indfiltre As Integer
    Dim bajout As Boolean
    Dim selected As Boolean
    Dim trouve As Boolean
    Dim ProchainIndex As Integer
    Dim s As String
    
    Call CL_Init
    n = 0
    grdFeuille.col = ColGrdFeuilLib
    ' ceux qui y sont déja
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.row = I
        If grdFeuille.CellFontBold = True Then
            selected = True
            Call CL_AddLigne(grdFeuille.TextMatrix(I, ColGrdFeuilLib), I + 1, "", selected)
        Else
            selected = False
            Call CL_AddLigne(grdFeuille.TextMatrix(I, ColGrdFeuilLib), I + 1, "", selected)
        End If
        'Call CL_AddLigne(grdFeuille.TextMatrix(i, ColGrdFeuilLib), i + 1, "", selected)
    Next I
    
    stitre = grdDocument.TextMatrix(grdDocument.row, ColGrdDocTitre)
    Call CL_InitTitreHelp(stitre & " : Feuilles Excel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    ' Ajouter cette fenetre dans le document
    lstFen = ""
    For I = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(I).selected Then
            lstFen = lstFen & "F" & CL_liste.lignes(I).num & ";"
        End If
    Next I
    grdDocument.TextMatrix(g_numDocument, ColGrdDocLstFen) = lstFen
    grddocument_click
    
    ' mettre enregistrer en visible
    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = True
End Sub

Private Sub PrmDest()
    
    Dim frm As Form
    Dim bret As Boolean
    Dim lstdest As String
    Dim I As Integer, n As Integer, s As String
    Dim prenom As String, nomutil As String, actif As Boolean, lib As String
    Dim fctnum As Integer, fctlibelle As String
    Dim srvnum As Integer, srvnom As String
    Dim lig As Integer, i_dest As Integer
    Dim lstDest_Out As String, op As String
    Dim lstDest_In As String
    Dim grpnum As Integer, grpnom As String, grpcode As String
    
    lstdest = grdDocument.TextMatrix(g_numDocument, ColGrdDocLstDest)
    lstDest_In = lstdest
    lstdest = Replace(lstdest, ";", "|")
    lstDest_Out = ""
    op = ""
    
    Set frm = ChoixDestinataire
    bret = ChoixDestinataire.AppelFrm(lstdest, _
                                     "Liste des destinataires")
    Set frm = Nothing
    If Not bret Then
        Exit Sub
    End If
    
    i_dest = g_numDocument
    lstdest = Replace(lstdest, "|", ";")
    
    'mettre les destinataires dans le grid
    n = STR_GetNbchamp(lstdest, ";")
    grdDest(i_dest).Rows = 0
    For I = 0 To n - 1
        s = STR_GetChamp(lstdest, ";", I)
        If left(s, 1) = "U" Then
            If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & Replace(s, "U", ""), _
                             prenom, nomutil, actif) = P_ERREUR Then
                s = "U0"
                lib = "Utilisateur ???"
            Else
                lib = prenom & "." & nomutil
            End If
            lstDest_Out = lstDest_Out & op & s
            op = ";"
        ElseIf left(s, 1) = "F" Then
            If Odbc_RecupVal("select FT_Num, FT_Libelle from fcttrav where FT_Num=" & Replace(s, "F", ""), _
                             fctnum, fctlibelle) = P_ERREUR Then
                s = "F0"
                lib = "Fonction ???"
            Else
                lib = "Fonction : " & fctlibelle
            End If
            lstDest_Out = lstDest_Out & op & s
            op = ";"
        ElseIf left(s, 1) = "G" Then
            If Odbc_RecupVal("select GU_Num, GU_Nom, GU_Code from groupeutil where GU_Num=" & Replace(s, "G", ""), _
                             grpnum, grpnom, grpcode) = P_ERREUR Then
                s = "G0"
                lib = "Fonction ???"
            Else
                lib = "Groupe : " & grpnom & " (" & grpcode & ")"
            End If
            lstDest_Out = lstDest_Out & op & s
            op = ";"
        ElseIf left(s, 1) = "S" Then
            If Odbc_RecupVal("select SRV_Num,SRV_Nom from service where SRV_Num=" & Replace(s, "S", ""), _
                             srvnum, srvnom) = P_ERREUR Then
                s = "S0"
                lib = "Service ???"
            Else
                lib = "Service : " & srvnom
            End If
            lstDest_Out = lstDest_Out & op & s
            op = ";"
        End If
        If s <> "" Then
            If grdDest(i_dest).Rows = 0 Then
                lig = 0
                grdDest(i_dest).Rows = 1
            Else
                lig = grdDest(i_dest).Rows
                grdDest(i_dest).Rows = grdDest(i_dest).Rows + 1
            End If
            grdDest(i_dest).TextMatrix(lig, ColGrdDestNum) = s
            grdDest(i_dest).TextMatrix(lig, ColGrdDestLib) = lib
        End If
    Next I
    
    lstDest_Out = lstDest_Out & ";"
    
    ' Modifiée ?
    If lstDest_In <> lstDest_Out Then
        cmd(CMD_OK).Visible = True
        cmd(CMD_OK).Enabled = True
        'on le remet à sa place
        grdDocument.TextMatrix(g_numDocument, ColGrdDocLstDest) = lstDest_Out
    End If
End Sub
Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Index = CMD_FERMER Then
        g_mode_saisie = False
    End If
    
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
        Call valider
    ElseIf KeyCode = vbKeyEscape Then
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
        If Not quitter() Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub grddocument_click()
    Dim iRow As Integer
    Dim lstFen As String
    Dim lstdest As String
    Dim n As Integer, I As Integer, J As Integer
    Dim s As String
    Dim laCol As Integer
    Dim FichierHTML As String
    Dim strDocNum As String, strDocChk As String
    Dim strDocNature As String, strDocModele As String
    Dim nb As Integer
    Dim AncBoolFaireChkClick As Boolean
    
    iRow = grdDocument.row
    laCol = grdDocument.ColSel
    'Me.FrmPublier.Visible = False
    If iRow < 0 Then
        ' pas encore de document : proposer création
        MsgBox "Vous devez d'abord créer un document"
        Call Ajout_Doc
        iRow = grdDocument.row
        'If iRow < 0 Then
        '    Exit Sub
        'End If
    End If
    
    If grdDocument.Rows > 0 And Faire_Doc_Click Then
        g_numDocument = iRow
        
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            grdDocument.TextMatrix(I, ColGrdDocaFen) = ""
            If I = iRow Then
                For J = 0 To grdDocument.Cols - 1
                    grdDocument.col = J
                    If J = ColGrdDocaFen Then grdDocument.TextMatrix(I, J) = ">>"
                    grdDocument.CellBackColor = grdDocument.BackColorFixed
                    grdDocument.CellFontBold = True
                Next J
            Else
                For J = 0 To grdDocument.Cols - 1
                    grdDocument.col = J
                    If J = ColGrdDocaFen Then grdDocument.TextMatrix(I, J) = ""
                    grdDocument.CellBackColor = grdDocument.BackColorBkg
                    grdDocument.CellFontBold = False
                Next J
            End If
        Next I
        grdDocument.row = iRow
        
        'cmd(CMD_PARAM_PUBLIER).Visible = True
        'cmd(CMD_AJOUT_FENDOC).Visible = True
        'cmd(CMD_SIMULATION_Un).Visible = True
        'cmd(CMD_SIMULATION_Tous).Visible = True
        'cmd(CMD_VOIR_RESULTATS).Visible = True
        
        lstFen = grdDocument.TextMatrix(grdDocument.row, ColGrdDocLstFen)
        lstdest = grdDocument.TextMatrix(grdDocument.row, ColGrdDocLstDest)
        ' mettre la ligne en surbrillance
        Faire_Doc_Click = False
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            If I = iRow Then
                grdDocument.CellFontBold = True
            Else
                grdDocument.CellFontBold = False
            End If
            'AfficheDest (i)
        Next I
        ' mettre les fenetres en non surbrillance
        For I = 0 To grdFeuille.Rows - 1
            grdFeuille.row = I
            grdFeuille.RowHeight(I) = 0
            grdFeuille.col = ColGrdFeuilLib
            grdFeuille.CellFontBold = False
            grdFeuille.TextMatrix(I, ColGrdFeuilaDoc) = ""
        Next I
        ' mettre les fenetres concernées en surbrillance
        If lstFen = "" Then
            cmd(CMD_AJOUT_FEN).Visible = True
            cmd(CMD_SUPPR_FEN).Visible = False
            For J = 0 To grdFeuille.Rows - 1
                grdFeuille.row = J
                grdFeuille.RowHeight(J) = 0
                grdFeuille.TextMatrix(J, ColGrdFeuilaDoc) = ""
            Next J
        Else
            n = STR_GetNbchamp(lstFen, ";")
            For I = 0 To n - 1
                s = STR_GetChamp(lstFen, ";", I)
                s = Replace(s, "F", "")
                For J = 0 To grdFeuille.Rows - 1
                    grdFeuille.row = J
                    'grdFeuille.col = 1
                    grdFeuille.TextMatrix(J, ColGrdFeuilTag) = " "
                    If grdFeuille.TextMatrix(J, ColGrdFeuilNum) = s Then
                        grdFeuille.TextMatrix(J, ColGrdFeuilaDoc) = ">>"
                        grdFeuille.RowHeight(J) = 400
                        grdFeuille.col = ColGrdFeuilaDoc
                        grdFeuille.CellFontBold = True
                        grdFeuille.col = ColGrdFeuilLib
                        grdFeuille.CellFontBold = True
                        cmd(CMD_SUPPR_FEN).Visible = True
                        nb = nb + 1
                        Exit For
                    End If
                Next J
            Next I
            If nb = grdFeuille.Rows Then
                cmd(CMD_AJOUT_FEN).Visible = False
            Else
                cmd(CMD_AJOUT_FEN).Visible = True
            End If
        End If
        ' grdDest en invisible
        For I = 0 To grdDocument.Rows - 1
            On Error Resume Next
            grdDest(I).Visible = False
        Next I
        On Error GoTo 0
        
        strDocChk = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 0)
        
        If strDocChk = "1" Then
            Me.ChkPublierKaliDoc.Visible = True
            AncBoolFaireChkClick = p_BoolFaireChkClick
            p_BoolFaireChkClick = False
            Me.ChkPublierKaliDoc.Value = 1
            p_BoolFaireChkClick = AncBoolFaireChkClick
            
            Set cmd(CMD_ICONE_KALIDOC).Picture = ImageListS.ListImages(IMG_KALIDOC).Picture
            cmd(CMD_ICONE_KALIDOC).Visible = True
            cmd(CMD_ICONE_KALIDOC).ToolTipText = "tableau de bord publié dans KaliDoc"
            
            grdDocument.row = iRow
            grdDocument.col = ColGrdKaliDoc
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
        Else
            Me.ChkPublierKaliDoc.Visible = True
            AncBoolFaireChkClick = p_BoolFaireChkClick
            p_BoolFaireChkClick = False
            Me.ChkPublierKaliDoc.Value = 0
            p_BoolFaireChkClick = AncBoolFaireChkClick
            
            Me.lbl(2).Visible = False
            Me.cmd(CMD_AJOUT_DEST).Visible = False
            Set cmd(CMD_ICONE_KALIDOC).Picture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            cmd(CMD_ICONE_KALIDOC).Visible = True
            cmd(CMD_ICONE_KALIDOC).ToolTipText = "tableau de bord local non publié dans KaliDoc"
            cmd(CMD_SIMULATION_Un).Visible = True
            cmd(CMD_SIMULATION_Tous).Visible = True
            
            grdDocument.row = iRow
            grdDocument.col = ColGrdKaliDoc
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
        
        End If
        AfficheDest "D", iRow
        grdDocument.row = iRow
        
        Faire_Doc_Click = True
        
        If grdDocument.ColSel = ColGrdDocExcel Then
            grdDocument.row = g_numDocument
            If grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_EXCEL).Picture Then
                Public_VerifOuvrir grdDocument.TextMatrix(g_numDocument, ColGrdDocCheminExcel), True, True, p_tbl_FichExcelPublier
            End If
        ElseIf grdDocument.ColSel = ColGrdDocHTML Then
            grdDocument.row = g_numDocument
            If grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_HTML).Picture Then
                FichierHTML = Replace(grdDocument.TextMatrix(g_numDocument, ColGrdDocCheminExcel), ".xls", ".htm")
                If FICH_FichierExiste(FichierHTML) Then
                    StartProcess FichierHTML
                Else
                    MsgBox "Fichier " & FichierHTML & " introuvable"
                End If
            End If
        End If
        grdDocument.row = g_numDocument
        'Me.FrmDocument.Visible = True
        'Me.FrmDocument.Caption = grdDocument.TextMatrix(iRow, ColGrdDocTitre)
        
        p_boolChkCliqué = False
        Call MenF_Dossier
        p_boolChkCliqué = True
    End If
    
    If iRow >= 0 Then grdDocument.row = iRow
    
    If Me.ChkPublierKaliDoc.Value = 1 Then
       'cmd(CMD_VOIR_DIFFUSION).Visible = True
    Else
       'cmd(CMD_VOIR_DIFFUSION).Visible = False
    End If
    p_BoolFaireChkClick = True
End Sub

Private Sub MenF_Dossier()
    Dim strDocChk As String
    Dim strDocNum As String
    Dim strDocNature As String
    Dim strDocModele As String
    Dim iRow As Integer
    
    iRow = grdDocument.row
    strDocChk = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 0)
    strDocNum = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 1)
    strDocNature = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 2)
    strDocModele = STR_GetChamp(grdDocument.TextMatrix(iRow, ColGrdDocParam), ";", 3)
         
    If strDocChk = "" Then
        If p_BoolFaireChkClick Then Me.ChkPublierKaliDoc.Value = 0
        Me.cmd(CMD_CHOIX_DOSSIER).Visible = False
        Me.cmd(CMD_CHOIX_NATURE).Visible = False
        Me.cmd(CMD_CHOIX_MODELE).Visible = False
        Me.LblDossier.Visible = False
    Else
        Me.cmd(CMD_CHOIX_DOSSIER).Visible = True
        Me.cmd(CMD_CHOIX_DOSSIER).tag = strDocNum
        Me.cmd(CMD_CHOIX_NATURE).tag = strDocNature
        Me.cmd(CMD_CHOIX_MODELE).tag = strDocModele
        If strDocNum <> "" Then
            Me.cmd(CMD_CHOIX_NATURE).Visible = True
        End If
        If strDocNature <> "" Then
            Me.cmd(CMD_CHOIX_MODELE).Visible = True
        End If
        
        Me.LblDossier.Visible = True
        Me.LblDossier.Caption = FaitLib(strDocNum, strDocNature, strDocModele)
        If p_BoolFaireChkClick Then Me.ChkPublierKaliDoc.Value = 1
    End If

End Sub

Private Function FaitLib(v_dossier As String, v_Nature As String, v_modele As String)
    FaitLib = "[" & ChercherNomDossier(v_dossier) & "]"
    FaitLib = FaitLib & "  [" & ChercherNomNature(v_Nature) & "]"
    FaitLib = FaitLib & "  [" & ChercherNomModele(v_modele) & "]"
End Function

Private Sub grdDocument_DblClick()
    Dim stitre As String
    
    stitre = grdDocument.TextMatrix(grdDocument.row, ColGrdDocTitre)
    stitre = InputBox("Indiquez le titre du document", "Titre du document", stitre)
    If stitre <> "" Then
        If stitre <> grdDocument.TextMatrix(grdDocument.row, ColGrdDocTitre) Then
            grdDocument.TextMatrix(grdDocument.row, ColGrdDocTitre) = stitre
            cmd(CMD_OK).Visible = True
            cmd(CMD_OK).Enabled = True
        End If
    End If

End Sub

Private Sub grdFeuille_Click()
    Dim I As Integer
    Dim J As Integer
    Dim Anc_Faire_Doc_Click As Boolean
    Dim isel As Integer
    Dim s As String
    Dim AncDocRow As Integer
    
    For I = 0 To grdDocument.Rows - 1
        On Error Resume Next
        grdDest(I).Visible = False
    Next I
    cmd(CMD_AJOUT_DEST).Visible = False
    
    g_numfeuille = grdFeuille.RowSel + 1
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.col = ColGrdFeuilLib
        grdFeuille.row = I
        grdFeuille.CellFontBold = False
        grdFeuille.TextMatrix(I, ColGrdFeuilTag) = " "
        If grdFeuille.row = g_numfeuille - 1 Then
            For J = 0 To grdFeuille.Cols - 1
                grdFeuille.col = J
                If J = ColGrdFeuilaDoc Then grdFeuille.TextMatrix(I, J) = "<<"
                grdFeuille.CellBackColor = grdFeuille.BackColorFixed
                grdFeuille.CellFontBold = True
            Next J
        Else
            For J = 0 To grdFeuille.Cols - 1
                grdFeuille.col = J
                If J = ColGrdFeuilaDoc Then grdFeuille.TextMatrix(I, J) = ""
                grdFeuille.CellBackColor = grdFeuille.BackColorBkg
                grdFeuille.CellFontBold = False
            Next J
        End If
    Next I
    grdFeuille.row = g_numfeuille - 1
    
    ' mettre les documents qui comportent cette fenetre en surbrillance
    Anc_Faire_Doc_Click = Faire_Doc_Click
    Faire_Doc_Click = False
    AncDocRow = grdDocument.row
    For I = 0 To grdDocument.Rows - 1
        grdDocument.row = I
        s = grdDocument.TextMatrix(I, ColGrdDocLstFen)
        
        If InStr(s, "F" & g_numfeuille & ";") > 0 Then
            grdDocument.col = ColGrdDocTitre
            grdDocument.TextMatrix(I, ColGrdDocaFen) = "<<"
            grdDocument.CellFontBold = True
            grdDocument.col = ColGrdDocaFen
            grdDocument.CellFontBold = True
        Else
            grdDocument.col = ColGrdDocTitre
            grdDocument.TextMatrix(I, ColGrdDocaFen) = ""
            grdDocument.CellFontBold = False
            grdDocument.col = ColGrdDocaFen
            grdDocument.CellFontBold = False
        End If
    Next I
    grdDocument.row = AncDocRow
    Faire_Doc_Click = Anc_Faire_Doc_Click

    grdFeuille.row = g_numfeuille - 1
    grdFeuille.col = ColGrdFeuilNum
    Me.SetFocus

End Sub

Private Sub AfficheDest(v_Trait As String, v_i As Integer)
    Dim sdest As String
    Dim n As Integer
    Dim I As Integer, s As String
    Dim lstdest As String
    Dim prenom As String, nomutil As String, actif As Boolean
    Dim srvnum As Integer, srvnom As String
    Dim fctnum As Integer, fctlibelle As String
    Dim lib As String
    Dim lig As Integer
    
    Me.lbl(2).Visible = False
    Me.cmd(CMD_AJOUT_DEST).Visible = False
    If v_Trait = "D" Then
        grdDocument.Visible = True
        cmd(CMD_AJOUT_DEST).Visible = True
        ' les destinataires sont dans grdDocument de v_i
        lstdest = grdDocument.TextMatrix(v_i, ColGrdDocLstDest)
        
        p_BoolFaireChkClick = False
        Me.ChkDocPublic.Value = IIf(grdDocument.TextMatrix(v_i, ColGrdDocDocPublic) = "", 0, grdDocument.TextMatrix(v_i, ColGrdDocDocPublic))
        
        'Me.ChkDocPublic.Value = STR_GetChamp(lstdest, "%", 1)
        'lstdest = STR_GetChamp(lstdest, "%", 0)
        ' charger le grid
        On Error GoTo Err_Load
        Load grdDest(v_i)
        AfficheDest "D", grdDocument.row
        'grdDest(v_i).ColWidth(ColGrdDestNum) = 1000
        'grdDest(v_i).ColWidth(ColGrdDestLib) = grdDest(v_i).Width - 100
        GoTo Suite_Load
Err_Load:
        Resume Suite_Load
Suite_Load:
        grdDest(v_i).Rows = 0
        grdDest(v_i).Cols = 2
        grdDest(v_i).ColWidth(ColGrdDestNum) = 0    '1000
        grdDest(v_i).ColWidth(ColGrdDestLib) = grdDest(v_i).Width - 100
        On Error GoTo 0
        grdDest(v_i).Visible = True
        Me.lbl(2).Visible = True
        Me.cmd(CMD_AJOUT_DEST).Visible = True
        n = STR_GetNbchamp(lstdest, ";")
        For I = 0 To n - 1
            s = STR_GetChamp(lstdest, ";", I)
            If left(s, 1) = "U" Then
                If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & Replace(s, "U", ""), _
                                 prenom, nomutil, actif) = P_ERREUR Then
                    s = "U0"
                    lib = "Utilisateur ???"
                Else
                    lib = prenom & "." & nomutil
                End If
            ElseIf left(s, 1) = "F" Then
                If Odbc_RecupVal("select FT_Num, FT_Libelle from fcttrav where FT_Num=" & Replace(s, "F", ""), _
                                 fctnum, fctlibelle) = P_ERREUR Then
                    s = "F0"
                    lib = "Fonction ???"
                Else
                    lib = "Fonction : " & fctlibelle
                End If
            ElseIf left(s, 1) = "S" Then
                If Odbc_RecupVal("select SRV_Num,SRV_Nom from service where SRV_Num=" & Replace(s, "S", ""), _
                                 srvnum, srvnom) = P_ERREUR Then
                    s = "S0"
                    lib = "Service ???"
                Else
                    lib = "Service : " & srvnom
                End If
            End If
            If grdDest(v_i).Rows = 0 Then
                lig = 0
                grdDest(v_i).Rows = 1
            Else
                lig = grdDest(v_i).Rows
                grdDest(v_i).Rows = grdDest(v_i).Rows + 1
            End If
            grdDest(v_i).TextMatrix(lig, ColGrdDestNum) = s
            grdDest(v_i).TextMatrix(lig, ColGrdDestLib) = lib
        Next I
    End If
End Sub

