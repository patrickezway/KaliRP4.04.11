VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrmPublier 
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
      Caption         =   "Paramètres de la publication d'un rapport"
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
         TabIndex        =   30
         Top             =   4920
         Visible         =   0   'False
         Width           =   8175
         Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
            Height          =   255
            Left            =   2880
            TabIndex        =   31
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
            TabIndex        =   32
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
            TabIndex        =   35
            Top             =   240
            Width           =   7455
         End
         Begin VB.Label lblHTTPDTemps 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblHTTPDTaille 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Index           =   9
         Left            =   5400
         Picture         =   "PrmPublier.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Accès à l'aide"
         Top             =   240
         UseMaskColor    =   -1  'True
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
         Picture         =   "PrmPublier.frx":0359
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
         Picture         =   "PrmPublier.frx":07B0
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
            TabIndex        =   28
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
            Picture         =   "PrmPublier.frx":0BF7
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
            Left            =   270
            TabIndex        =   21
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Dossier"
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
            Index           =   25
            Left            =   2700
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Nature"
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
               Size            =   8.25
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
         Picture         =   "PrmPublier.frx":1022
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
         Picture         =   "PrmPublier.frx":1469
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
         Picture         =   "PrmPublier.frx":18C0
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
               Picture         =   "PrmPublier.frx":1CEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmPublier.frx":299D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmPublier.frx":2EEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmPublier.frx":34A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmPublier.frx":3A63
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
      Top             =   7965
      Width           =   10785
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Gérer les résultats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1845
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
         Height          =   510
         Index           =   0
         Left            =   480
         Picture         =   "PrmPublier.frx":418D
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
         Picture         =   "PrmPublier.frx":46F6
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmPublier"
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
Private Const CMD_FAIRE_DIFFUSION = 12
Private Const CMD_CHOIX_NATURE = 21
Private Const CMD_CHOIX_MODELE = 20
Private Const CMD_PUBLIER = 4

Private Const IMG_KALIDOC = 3
Private Const IMG_PAS_KALIDOC = 4

Private g_numModele As Long
Private g_numfeuille As Integer
Private g_numDocument As Integer
Private g_bcr As Boolean
Private g_CheminModele As String
Private g_mode_saisie As Boolean
Private g_form_active As Boolean
Private g_DocParamDefaut As String

Private Const IMG_LOAD_EXCEL = 1
Private Const IMG_LOAD_HTML = 2

Private Faire_Doc_Click As Boolean

' pour le grid documents
Private Const GRDDOC_NUMDOC = 0
Private Const GRDDOC_TITRE = 1
Private Const GRDDOC_LSTFEN = 2
Private Const GRDDOC_LSTDEST = 3
Private Const GRDDOC_PUBLIC = 4
Private Const GRDDOC_PUBLIER_KD = 5
Private Const GRDDOC_NUMNAT = 6
Private Const GRDDOC_NUMDOS = 7
Private Const GRDDOC_MODELE = 8
Private Const GRDDOC_SEL = 9
Private Const GRDDOC_IMGKD = 10
Private Const GRDDOC_IDMODELE = 11

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

Public Function AppelFrm(ByRef v_nummodele As Long, _
                         ByVal v_CheminModele As String) As Boolean

    g_numModele = v_nummodele
    g_CheminModele = v_CheminModele
        
    Show 1
    
    AppelFrm = g_bcr
    
End Function

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
    Public_VerifOuvrir g_CheminModele & p_PointExtensionXls, False, False, p_tbl_FichExcelPublier
    
    RemplirTabFenetre
    
    FrmPublier.Visible = False
    cmd(CMD_OK).Visible = False
    cmd(CMD_AJOUT_FEN).Visible = True
    cmd(CMD_PARAM_PUBLIER).Visible = False
    
    g_numDocument = -1
    
    FrmPublier.Visible = False
    ' se mettre sur le premier document
    If grdDocument.Rows > 0 Then
        grdDocument.row = 0
        grddocument_click
    Else
        lbl(2).Visible = False
        cmd(CMD_AJOUT_DEST).Visible = False
    End If
    
    cmd(CMD_OK).Visible = False
    cmd(CMD_OK).Enabled = False
    
    Faire_Doc_Click = True
    
    p_BoolFaireChkClick = True
    
    g_mode_saisie = True

End Sub

Private Sub ChargerFichier()
    
    Dim sql As String
    Dim titre As String
    Dim lig As Integer
    Dim rs As rdoResultset
    Dim rs1 As rdoResultset
    
    grdDocument.Rows = 0
    grdDocument.Cols = 12
    grdDocument.ColWidth(GRDDOC_NUMDOC) = 0
    grdDocument.ColWidth(GRDDOC_LSTFEN) = 0
    grdDocument.ColWidth(GRDDOC_LSTDEST) = 0
    grdDocument.ColWidth(GRDDOC_PUBLIC) = 0
    grdDocument.ColWidth(GRDDOC_PUBLIER_KD) = 0
    grdDocument.ColWidth(GRDDOC_NUMNAT) = 0
    grdDocument.ColWidth(GRDDOC_NUMDOS) = 0
    grdDocument.ColWidth(GRDDOC_MODELE) = 0
    grdDocument.ColWidth(GRDDOC_IDMODELE) = 0
    grdDocument.ColWidth(GRDDOC_TITRE) = 3800
    grdDocument.ColWidth(GRDDOC_SEL) = 500
    grdDocument.ColWidth(GRDDOC_IMGKD) = 500
    
    sql = "select * from rp_document where rpd_rpnum=" & g_numModele
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        cmd(CMD_PUBLIER).Visible = True
    End If
    While Not rs.EOF
        grdDocument.AddItem ""
        lig = grdDocument.Rows - 1
        grdDocument.TextMatrix(lig, GRDDOC_NUMDOC) = rs("rpd_num").Value
        grdDocument.TextMatrix(lig, GRDDOC_TITRE) = rs("rpd_titre").Value
        grdDocument.TextMatrix(lig, GRDDOC_LSTFEN) = rs("rpd_lstfeuille").Value
        grdDocument.TextMatrix(lig, GRDDOC_LSTDEST) = rs("rpd_lstdest").Value
        grdDocument.TextMatrix(lig, GRDDOC_PUBLIC) = rs("rpd_public").Value
        grdDocument.TextMatrix(lig, GRDDOC_PUBLIER_KD) = rs("rpd_publier_kd").Value
        grdDocument.row = lig
        grdDocument.col = GRDDOC_IMGKD
        If rs("rpd_publier_kd").Value Then
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
        Else
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
        End If
        grdDocument.TextMatrix(lig, GRDDOC_NUMNAT) = rs("rpd_ndnum").Value
        grdDocument.TextMatrix(lig, GRDDOC_NUMDOS) = rs("rpd_dsnum").Value
        If IsNumeric(rs("rpd_modele").Value) Then
            sql = "select * from modeledoc where modc_num = " & rs("rpd_modele").Value
            If Odbc_SelectV(sql, rs1) = P_ERREUR Then
                Exit Sub
            End If
            titre = " " & rs("rpd_modele").Value & "non trouvé"
            If Not rs1.EOF Then
                titre = rs1("modc_titre").Value
            End If
            grdDocument.TextMatrix(lig, GRDDOC_MODELE) = titre  ' rs("rpd_modele").Value
            grdDocument.TextMatrix(lig, GRDDOC_IDMODELE) = rs("rpd_modele").Value
        Else
            MsgBox "Veuillez sélectionner un modèle"
            grdDocument.TextMatrix(lig, GRDDOC_MODELE) = ""
            grdDocument.TextMatrix(lig, GRDDOC_IDMODELE) = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    
End Sub


Private Function quitter() As Boolean

    Dim reponse As Integer
    Dim LaUbound As Integer
    Dim I As Integer
    Dim j As Integer
    
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
            For j = 1 To Exc_obj.Workbooks.Count
                If UCase(Exc_obj.Workbooks(j).FullName) = UCase(p_tbl_FichExcelPublier(I).FichFullname) Then
                    Exc_obj.Workbooks(j).Close True
                Else
                    Exc_obj.Workbooks(j).Close False
                End If
            Next j
        End If
    Next I
Faire:
    
    If Exc_obj Is Nothing Then
        GoTo lab_fin
    End If
    If Exc_obj.Workbooks.Count = 0 Then
        Exc_obj.Quit
        Set Exc_obj = Nothing
    End If

lab_fin:
    Unload Me
    
    quitter = True
    
End Function

Private Sub valider()

    Dim sF As String, sD As String
    Dim I As Integer
    Dim lbid As Long
    
    For I = 0 To grdDocument.Rows - 1
        sF = grdDocument.TextMatrix(I, GRDDOC_LSTFEN)
        sD = grdDocument.TextMatrix(I, GRDDOC_LSTDEST)
        If sD & sF <> "" Then
            If grdDocument.TextMatrix(I, GRDDOC_NUMDOC) = "" Then
                ' Ajout
                Call Odbc_AddNew("rp_document", "rpd_num", "rpd_seq", False, lbid, _
                                 "rpd_rpnum", g_numModele, _
                                 "rpd_titre", grdDocument.TextMatrix(I, GRDDOC_TITRE), _
                                 "rpd_lstfeuille", grdDocument.TextMatrix(I, GRDDOC_LSTFEN), _
                                 "rpd_public", CBool(grdDocument.TextMatrix(I, GRDDOC_PUBLIC)), _
                                 "rpd_lstdest", grdDocument.TextMatrix(I, GRDDOC_LSTDEST), _
                                 "rpd_publier_kd", CBool(grdDocument.TextMatrix(I, GRDDOC_PUBLIER_KD)), _
                                 "rpd_dsnum", grdDocument.TextMatrix(I, GRDDOC_NUMDOS), _
                                 "rpd_ndnum", grdDocument.TextMatrix(I, GRDDOC_NUMNAT), _
                                 "rpd_modele", grdDocument.TextMatrix(I, GRDDOC_IDMODELE))
            Else
                ' Modif
                Call Odbc_Update("rp_document", "rpd_num", "where rpd_num=" & grdDocument.TextMatrix(I, GRDDOC_NUMDOC), _
                                 "rpd_titre", grdDocument.TextMatrix(I, GRDDOC_TITRE), _
                                 "rpd_lstfeuille", grdDocument.TextMatrix(I, GRDDOC_LSTFEN), _
                                 "rpd_public", CBool(grdDocument.TextMatrix(I, GRDDOC_PUBLIC)), _
                                 "rpd_lstdest", grdDocument.TextMatrix(I, GRDDOC_LSTDEST), _
                                 "rpd_publier_kd", CBool(grdDocument.TextMatrix(I, GRDDOC_PUBLIER_KD)), _
                                 "rpd_dsnum", grdDocument.TextMatrix(I, GRDDOC_NUMDOS), _
                                 "rpd_ndnum", grdDocument.TextMatrix(I, GRDDOC_NUMNAT), _
                                 "rpd_modele", grdDocument.TextMatrix(I, GRDDOC_IDMODELE))
            End If
        End If
    Next I
    
    Unload Me
    
End Sub

Private Sub ChkDocPublic_Click()
    
    If p_BoolFaireChkClick Then
        grdDocument.TextMatrix(grdDocument.row, GRDDOC_PUBLIC) = IIf(ChkDocPublic.Value = 1, True, False)
        cmd(CMD_OK).Visible = True
        cmd(CMD_OK).Enabled = True
        cmd(CMD_PUBLIER).Visible = False
    End If

End Sub

Private Sub ChkPublierKaliDoc_Click()
    
    Dim FaiteAuto As Boolean
    
    FaiteAuto = False
    If p_BoolFaireChkClick Then
        If ChkPublierKaliDoc.Value = 1 Then
            grdDocument.TextMatrix(grdDocument.row, GRDDOC_PUBLIER_KD) = True
            If grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOS) = 0 Then
                FaiteAuto = True
            End If
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
            cmd(CMD_CHOIX_DOSSIER).Visible = True
        Else
            grdDocument.TextMatrix(grdDocument.row, GRDDOC_PUBLIER_KD) = False
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
            cmd(CMD_CHOIX_DOSSIER).Visible = False
        End If
        Call AfficheDest("D", grdDocument.row)
        cmd(CMD_AJOUT_DEST).Visible = True
        If p_boolChkCliqué Then
            cmd(CMD_OK).Visible = True
            cmd(CMD_OK).Enabled = True
            cmd(CMD_PUBLIER).Visible = False
        End If
    End If
    p_BoolFaireChkClick = False
    Call MenF_Dossier
    p_BoolFaireChkClick = True
    If FaiteAuto Then
        If grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOS) = 0 Then
            If Choisir_dossier() = P_OUI Then
                If Choisir_Nature() = P_OUI Then
                    Call Choisir_Modele
                End If
            End If
        End If
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim frm As Form
    
    Select Case Index
    Case CMD_AIDE
        Call Appel_Aide
    Case CMD_AJOUT_FEN
        Call Ajout_FenDoc
    Case CMD_CHOIX_NATURE
        Call Choisir_Nature
    Case CMD_CHOIX_MODELE
        Call Choisir_Modele
    Case CMD_CHOIX_DOSSIER
        Call Choisir_dossier
    Case CMD_AJOUT_DOC
        Call Ajouter_Doc
    Case CMD_SUPPR_DOC
        Call supprimer_doc
    Case CMD_AJOUT_DEST
        Call PrmDest
    Case CMD_OK
        Call valider
    Case CMD_FERMER
        Call quitter
    Case CMD_PUBLIER
        Set frm = Publier
        Publier.AppelFrm (g_numModele)
        Set frm = Nothing
    End Select
    
End Sub

Private Function Choisir_dossier() As Integer

    Dim sql As String, sret As String
    Dim numDos As Long
    Dim frm As Form
    Dim rs As rdoResultset
    
    Choisir_dossier = P_NON
    
    Set frm = ChoixDossier
    sret = ChoixDossier.AppelFrm(0, 0)
    Set frm = Nothing
    If sret = "" Then
        Exit Function
    End If
    
    numDos = sret
    If numDos > 0 Then
        sql = "select * from Dossier where Ds_num = " & numDos
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Choisir_dossier = P_ERREUR
            Exit Function
        End If
        LblDossier.Caption = "Choix du Dossier"   'ChercherNomDossier(retDos)
        If Not rs.EOF Then
            LblDossier.Caption = "Dossier"   'ChercherNomDossier(retDos)
            cmd(CMD_CHOIX_DOSSIER).tag = numDos
            grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOS) = numDos
            cmd(CMD_OK).Visible = True
            cmd(CMD_OK).Enabled = True
            cmd(CMD_PUBLIER).Visible = False
            Choisir_dossier = P_OUI
        End If
        rs.Close
        Call MenF_Dossier
    End If

End Function

Private Function Choisir_Modele()
    
    Dim sql As String, modele As String
    Dim rs1 As rdoResultset
    Dim titre As String
    Dim I As Integer, n As Integer
    Dim numdocs As Long
    Dim IdModele As Variant
    Dim rs As rdoResultset
    
    If cmd(CMD_CHOIX_DOSSIER).tag = 0 Then
        Call MsgBox("Veuillez d'abord choisir un dossier.", vbOKOnly + vbInformation, "")
        Choisir_Modele = P_NON
        Exit Function
    End If
    
    If cmd(CMD_CHOIX_NATURE).tag = 0 Then
        Call MsgBox("Veuillez d'abord choisir une nature.", vbOKOnly + vbInformation, "")
        Choisir_Modele = P_NON
        Exit Function
    End If
    
    Call CL_Init
    
    sql = "select ds_donum from Dossier" _
        & " where DS_Num = " & cmd(CMD_CHOIX_DOSSIER).tag
    If Odbc_RecupVal(sql, numdocs) = P_ERREUR Then
        Choisir_Modele = P_ERREUR
        Exit Function
    End If
    sql = "select distinct(" & P_MODCNUM_ou_MODELE & ") from DocsNatureModele" _
        & " where DONM_DONum =" & numdocs _
        & " and DONM_NDNum =" & cmd(CMD_CHOIX_NATURE).tag _
        & " order by " & P_MODCNUM_ou_MODELE
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Choisir_Modele = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        IdModele = rs(0).Value
        sql = "select * from modeledoc where MODC_Num = " & IdModele
        If Odbc_SelectV(sql, rs1) = P_ERREUR Then
            Choisir_Modele = P_ERREUR
            Exit Function
        End If
        titre = "Modèle " & IdModele & " non trouvé"
        If Not rs1.EOF Then
            titre = rs1("MODC_Titre")
        End If
        Call CL_AddLigne(titre, n, IdModele, False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        Call MsgBox("Aucun modèle n'a été trouvé.", vbInformation + vbOKOnly, "")
        Choisir_Modele = P_NON
        Exit Function
    End If
        
    Call CL_InitTitreHelp("Choix d'un modèle", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    'Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Choisir_Modele = P_NON
        Exit Function
    End If
    
    IdModele = CL_liste.lignes(CL_liste.pointeur).tag
    cmd(CMD_CHOIX_MODELE).tag = IdModele
    grdDocument.TextMatrix(grdDocument.row, GRDDOC_MODELE) = CL_liste.lignes(CL_liste.pointeur).texte
    grdDocument.TextMatrix(grdDocument.row, GRDDOC_IDMODELE) = IdModele
    Call MenF_Dossier
    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = True
    cmd(CMD_PUBLIER).Visible = False
    
    Choisir_Modele = P_OUI

End Function

Private Function Choisir_Nature() As Integer

    Dim sql As String
    Dim I As Integer, n As Integer
    Dim numnat As Long, numdocs As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    If cmd(CMD_CHOIX_DOSSIER).tag = 0 Then
        Call MsgBox("Veuillez d'abord choisir un dossier.", vbOKOnly + vbInformation, "")
        Choisir_Nature = P_NON
        Exit Function
    End If
    
    sql = "select ds_donum from Dossier" _
        & " where DS_Num = " & cmd(CMD_CHOIX_DOSSIER).tag
    If Odbc_RecupVal(sql, numdocs) = P_ERREUR Then
        Choisir_Nature = P_ERREUR
        Exit Function
    End If
    
    sql = "select * from NatureDoc" _
        & " where ND_Num in (select DON_NDNum from DocsNature where DON_DONum=" & numdocs & ")" _
        & " order by ND_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Choisir_Nature = P_ERREUR
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
        Choisir_Nature = P_NON
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Choix d'une nature", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Choisir_Nature = P_NON
        Exit Function
    End If
    
    numnat = CL_liste.lignes(CL_liste.pointeur).num
    cmd(CMD_CHOIX_NATURE).tag = numnat
    grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMNAT) = numnat
    Call MenF_Dossier
    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = True
    cmd(CMD_PUBLIER).Visible = False
    
    Choisir_Nature = P_OUI
    
End Function

Private Function ChercherNomNature(ByVal v_numNature)
    Dim sql As String, rs As rdoResultset
    
    If v_numNature <> "" Then
        sql = "select * from NatureDoc where ND_num = " & v_numNature
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            cmd(CMD_CHOIX_NATURE).tag = 0
            ChercherNomNature = ""
            Exit Function
        End If
        If Not rs.EOF Then
            ChercherNomNature = "Nature : " & rs("ND_Nom")
            cmd(CMD_CHOIX_NATURE).tag = v_numNature
            cmd(CMD_CHOIX_NATURE).Caption = "Nature"
        Else
            ChercherNomNature = ""
            cmd(CMD_CHOIX_NATURE).tag = 0
            cmd(CMD_CHOIX_NATURE).Caption = "Choix de la Nature"
        End If
        rs.Close
    End If
End Function

Private Function ChercherNomModele(ByVal v_nummodele)
    Dim sql As String, rs As rdoResultset
    
    If v_nummodele <> "" Then
        ChercherNomModele = "Modèle : " & v_nummodele
        cmd(CMD_CHOIX_MODELE).Caption = "Modèle"
    Else
        ChercherNomModele = ""
        cmd(CMD_CHOIX_MODELE).Caption = "Choix du Modèle"
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
            cmd(CMD_CHOIX_DOSSIER).Caption = "Dossier"
        Else
            ChercherNomDossier = ""
            cmd(CMD_CHOIX_DOSSIER).Caption = "Choix du Dossier"
        End If
        rs.Close
    End If
End Function

Private Sub supprimer_doc()
MsgBox "fonctionnalité à implémanter"
End Sub

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

Private Sub Ajouter_Doc()
    
    Dim stitre As String
    Dim lig As Integer
    
    stitre = InputBox("Indiquez le titre du Document", "Titre du Document", "")
    If stitre <> "" Then
        grdDocument.AddItem ""
        lig = grdDocument.Rows - 1
        grdDocument.TextMatrix(lig, GRDDOC_TITRE) = stitre
        grdDocument.TextMatrix(lig, GRDDOC_PUBLIC) = False
        grdDocument.TextMatrix(lig, GRDDOC_PUBLIER_KD) = False
        grdDocument.TextMatrix(lig, GRDDOC_NUMDOS) = 0
        grdDocument.TextMatrix(lig, GRDDOC_NUMNAT) = 0
        grdDocument.row = lig
        grdDocument.col = GRDDOC_IMGKD
        Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
        g_numDocument = grdDocument.Rows
        grdDocument.row = grdDocument.Rows - 1
        grddocument_click
        ' se mettre en ajout de fenetres
        Ajout_FenDoc
    End If
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
    
    'Si aucun document publié
    If grdDocument.row < 0 Then
        MsgBox "Aucun document publié pour cette feuille"
        Exit Sub
    End If
    
    stitre = grdDocument.TextMatrix(grdDocument.row, GRDDOC_TITRE)
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
    grdDocument.TextMatrix(g_numDocument, GRDDOC_LSTFEN) = lstFen
    grddocument_click
    
    ' mettre enregistrer en visible
    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = True
    cmd(CMD_PUBLIER).Visible = False
End Sub

Private Sub PrmDest()
    
    Dim frm As Form
    Dim bret As Boolean
    Dim lstdest As String, s As String, s2 As String
    Dim I As Integer, n As Integer, n2 As Integer
    Dim prenom As String, nomutil As String, actif As Boolean, lib As String
    Dim fctnum As Integer, fctlibelle As String, s_sp As String
    Dim srvnum As Integer, srvnom As String
    Dim ponum As Integer, ponom As String
    Dim lig As Integer, i_dest As Integer
    Dim lstDest_Out As String, op As String
    Dim lstDest_In As String
    Dim grpnum As Integer, grpnom As String, grpcode As String
    
    lstdest = grdDocument.TextMatrix(g_numDocument, GRDDOC_LSTDEST)
    lstDest_In = lstdest
    s = ""
    n = STR_GetNbchamp(lstdest, ";")
    For I = 0 To n - 1
        s2 = STR_GetChamp(lstdest, ";", I)
        If left$(s2, 1) <> "P" Then
            If left$(s2, 1) = "S" Then
                s = s & s2 & ";|"
            Else
                s = s & s2 & "|"
            End If
        Else
            s_sp = build_arbor_serv(s2)
            If left$(s2, 1) = "S" Then
                s = s & ";"
            End If
            s = s & s_sp & "|"
        End If
    Next I
    lstdest = s
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
    s = ""
    n = STR_GetNbchamp(lstdest, "|")
    For I = 0 To n - 1
        s2 = STR_GetChamp(lstdest, "|", I)
        If left$(s2, 1) <> "S" Then
            s = s & s2 & ";"
        Else
            n2 = STR_GetNbchamp(s2, ";")
            s = s & STR_GetChamp(s2, ";", n2 - 1) & ";"
        End If
    Next I
    lstdest = s
    
    'mettre les destinataires dans le grid
    n = STR_GetNbchamp(lstdest, ";")
    grdDest(i_dest).Rows = 0
    For I = 0 To n - 1
        s = STR_GetChamp(lstdest, ";", I)
        If left(s, 1) = "U" Then
            If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & Replace(s, "U", ""), _
                             prenom, nomutil, actif) = P_OK Then
                lib = prenom & "." & nomutil
                lstDest_Out = lstDest_Out & op & s
                op = ";"
            End If
        ElseIf left(s, 1) = "F" Then
            If Odbc_RecupVal("select FT_Num, FT_Libelle from fcttrav where FT_Num=" & Replace(s, "F", ""), _
                             fctnum, fctlibelle) = P_OK Then
                lib = "Fonction : " & fctlibelle
                lstDest_Out = lstDest_Out & op & s
                op = ";"
            End If
        ElseIf left(s, 1) = "G" Then
            If Odbc_RecupVal("select GU_Num, GU_Nom, GU_Code from groupeutil where GU_Num=" & Replace(s, "G", ""), _
                             grpnum, grpnom, grpcode) = P_OK Then
                lib = "Groupe : " & grpnom & " (" & grpcode & ")"
                lstDest_Out = lstDest_Out & op & s
                op = ";"
            End If
        ElseIf left(s, 1) = "S" Then
            If Odbc_RecupVal("select SRV_Num, SRV_Nom from service where SRV_Num=" & Replace(s, "S", ""), _
                             srvnum, srvnom) = P_OK Then
                lib = "Service : " & srvnom
                lstDest_Out = lstDest_Out & op & s
                op = ";"
            End If
        ElseIf left(s, 1) = "P" Then
            If Odbc_RecupVal("select PO_Num, PO_Libelle from poste where PO_Num=" & Replace(s, "P", ""), _
                             ponum, ponom) = P_OK Then
                lib = "Poste : " & ponom
                lstDest_Out = lstDest_Out & op & s
                op = ";"
            End If
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
        cmd(CMD_PUBLIER).Visible = False
        'on le remet à sa place
        grdDocument.TextMatrix(g_numDocument, GRDDOC_LSTDEST) = lstDest_Out
    End If
End Sub
Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

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
    Dim n As Integer, I As Integer, j As Integer
    Dim s As String
    Dim laCol As Integer
    Dim FichierHTML As String
    Dim strDocNum As String, strDocChk As String
    Dim strDocNature As String, strDocModele As String
    Dim nb As Integer
    Dim AncBoolFaireChkClick As Boolean
    
    iRow = grdDocument.row
    laCol = grdDocument.ColSel
    'FrmPublier.Visible = False
    If iRow < 0 Then
        ' pas encore de document : proposer création
        MsgBox "Vous devez d'abord créer un document"
        Call Ajouter_Doc
        iRow = grdDocument.row
        'If iRow < 0 Then
        '    Exit Sub
        'End If
    End If
    
    If grdDocument.Rows > 0 And Faire_Doc_Click Then
        
        FrmPublier.Visible = True
        g_numDocument = iRow
        
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            grdDocument.TextMatrix(I, GRDDOC_SEL) = ""
            If I = iRow Then
                For j = 0 To grdDocument.Cols - 1
                    grdDocument.col = j
                    If j = GRDDOC_SEL Then grdDocument.TextMatrix(I, j) = ">>"
                    grdDocument.CellBackColor = grdDocument.BackColorFixed
                    grdDocument.CellFontBold = True
                Next j
            Else
                For j = 0 To grdDocument.Cols - 1
                    grdDocument.col = j
                    If j = GRDDOC_SEL Then grdDocument.TextMatrix(I, j) = ""
                    grdDocument.CellBackColor = grdDocument.BackColorBkg
                    grdDocument.CellFontBold = False
                Next j
            End If
        Next I
        grdDocument.row = iRow
        
        'cmd(CMD_PARAM_PUBLIER).Visible = True
        'cmd(CMD_AJOUT_FENDOC).Visible = True
        'cmd(CMD_SIMULATION_Un).Visible = True
        'cmd(CMD_SIMULATION_Tous).Visible = True
        'cmd(CMD_VOIR_RESULTATS).Visible = True
        
        lstFen = grdDocument.TextMatrix(grdDocument.row, GRDDOC_LSTFEN)
        lstdest = grdDocument.TextMatrix(grdDocument.row, GRDDOC_LSTDEST)
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
            For j = 0 To grdFeuille.Rows - 1
                grdFeuille.row = j
                grdFeuille.RowHeight(j) = 0
                grdFeuille.TextMatrix(j, ColGrdFeuilaDoc) = ""
            Next j
        Else
            n = STR_GetNbchamp(lstFen, ";")
            For I = 0 To n - 1
                s = STR_GetChamp(lstFen, ";", I)
                s = Replace(s, "F", "")
                For j = 0 To grdFeuille.Rows - 1
                    grdFeuille.row = j
                    'grdFeuille.col = 1
                    grdFeuille.TextMatrix(j, ColGrdFeuilTag) = " "
                    If grdFeuille.TextMatrix(j, ColGrdFeuilNum) = s Then
                        grdFeuille.TextMatrix(j, ColGrdFeuilaDoc) = ">>"
                        grdFeuille.RowHeight(j) = 400
                        grdFeuille.col = ColGrdFeuilaDoc
                        grdFeuille.CellFontBold = True
                        grdFeuille.col = ColGrdFeuilLib
                        grdFeuille.CellFontBold = True
                        cmd(CMD_SUPPR_FEN).Visible = True
                        nb = nb + 1
                        Exit For
                    End If
                Next j
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
        
        If CBool(grdDocument.TextMatrix(iRow, GRDDOC_PUBLIER_KD)) Then
            ChkPublierKaliDoc.Visible = True
            AncBoolFaireChkClick = p_BoolFaireChkClick
            p_BoolFaireChkClick = False
            ChkPublierKaliDoc.Value = 1
            p_BoolFaireChkClick = AncBoolFaireChkClick
            grdDocument.row = iRow
            grdDocument.col = GRDDOC_IMGKD
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
        Else
            ChkPublierKaliDoc.Visible = True
            AncBoolFaireChkClick = p_BoolFaireChkClick
            p_BoolFaireChkClick = False
            ChkPublierKaliDoc.Value = 0
            p_BoolFaireChkClick = AncBoolFaireChkClick
            
            lbl(2).Visible = False
            cmd(CMD_AJOUT_DEST).Visible = False
            cmd(CMD_SIMULATION_Un).Visible = True
            cmd(CMD_SIMULATION_Tous).Visible = True
            
            grdDocument.row = iRow
            grdDocument.col = GRDDOC_IMGKD
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
        
        End If
        AfficheDest "D", iRow
        grdDocument.row = iRow
        
        Faire_Doc_Click = True
        
        grdDocument.row = g_numDocument
        'FrmDocument.Visible = True
        'FrmDocument.Caption = grdDocument.TextMatrix(iRow, ColGrdDocTitre)
        
        p_boolChkCliqué = False
        Call MenF_Dossier
        p_boolChkCliqué = True
    End If
    
    If iRow >= 0 Then grdDocument.row = iRow
    
    If ChkPublierKaliDoc.Value = 1 Then
       'cmd(CMD_VOIR_DIFFUSION).Visible = True
    Else
       'cmd(CMD_VOIR_DIFFUSION).Visible = False
    End If
    p_BoolFaireChkClick = True
End Sub

Private Sub MenF_Dossier()
    
    Dim iRow As Integer
    
    iRow = grdDocument.row
         
    If Not CBool(grdDocument.TextMatrix(iRow, GRDDOC_PUBLIER_KD)) Then
        If p_BoolFaireChkClick Then ChkPublierKaliDoc.Value = 0
        cmd(CMD_CHOIX_DOSSIER).Visible = False
        cmd(CMD_CHOIX_NATURE).Visible = False
        cmd(CMD_CHOIX_MODELE).Visible = False
        LblDossier.Visible = False
    Else
        cmd(CMD_CHOIX_DOSSIER).Visible = True
        cmd(CMD_CHOIX_DOSSIER).tag = grdDocument.TextMatrix(iRow, GRDDOC_NUMDOS)
        cmd(CMD_CHOIX_NATURE).tag = grdDocument.TextMatrix(iRow, GRDDOC_NUMNAT)
        cmd(CMD_CHOIX_MODELE).tag = grdDocument.TextMatrix(iRow, GRDDOC_IDMODELE)
        If grdDocument.TextMatrix(iRow, GRDDOC_NUMDOS) <> 0 Then
            cmd(CMD_CHOIX_NATURE).Visible = True
            If grdDocument.TextMatrix(iRow, GRDDOC_NUMNAT) <> 0 Then
                cmd(CMD_CHOIX_MODELE).Visible = True
            End If
        End If
        LblDossier.Visible = True
        LblDossier.Caption = FaitLib(grdDocument.TextMatrix(iRow, GRDDOC_NUMDOS), _
                                        grdDocument.TextMatrix(iRow, GRDDOC_NUMNAT), _
                                        grdDocument.TextMatrix(iRow, GRDDOC_MODELE))
        If p_BoolFaireChkClick Then ChkPublierKaliDoc.Value = 1
    End If

End Sub

Private Function FaitLib(ByVal v_dossier As String, _
                         ByVal v_Nature As String, _
                         ByVal v_modele As String) As String
                         
    FaitLib = "[" & ChercherNomDossier(v_dossier) & "]"
    FaitLib = FaitLib & "  [" & ChercherNomNature(v_Nature) & "]"
    FaitLib = FaitLib & "  [" & ChercherNomModele(v_modele) & "]"

End Function

Private Sub grdDocument_DblClick()
    Dim stitre As String
    
    stitre = grdDocument.TextMatrix(grdDocument.row, GRDDOC_TITRE)
    stitre = InputBox("Indiquez le titre du document", "Titre du document", stitre)
    If stitre <> "" Then
        If stitre <> grdDocument.TextMatrix(grdDocument.row, GRDDOC_TITRE) Then
            grdDocument.TextMatrix(grdDocument.row, GRDDOC_TITRE) = stitre
            cmd(CMD_OK).Visible = True
            cmd(CMD_OK).Enabled = True
            cmd(CMD_PUBLIER).Visible = False
        End If
    End If

End Sub

Private Sub grdFeuille_Click()
    Dim I As Integer
    Dim j As Integer
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
            For j = 0 To grdFeuille.Cols - 1
                grdFeuille.col = j
                If j = ColGrdFeuilaDoc Then grdFeuille.TextMatrix(I, j) = "<<"
                grdFeuille.CellBackColor = grdFeuille.BackColorFixed
                grdFeuille.CellFontBold = True
            Next j
        Else
            For j = 0 To grdFeuille.Cols - 1
                grdFeuille.col = j
                If j = ColGrdFeuilaDoc Then grdFeuille.TextMatrix(I, j) = ""
                grdFeuille.CellBackColor = grdFeuille.BackColorBkg
                grdFeuille.CellFontBold = False
            Next j
        End If
    Next I
    grdFeuille.row = g_numfeuille - 1
    
    ' mettre les documents qui comportent cette fenetre en surbrillance
    Anc_Faire_Doc_Click = Faire_Doc_Click
    Faire_Doc_Click = False
    AncDocRow = grdDocument.row
    For I = 0 To grdDocument.Rows - 1
        grdDocument.row = I
        s = grdDocument.TextMatrix(I, GRDDOC_LSTFEN)
        
        If InStr(s, "F" & g_numfeuille & ";") > 0 Then
            grdDocument.TextMatrix(I, GRDDOC_SEL) = "<<"
            grdDocument.CellFontBold = True
        Else
            grdDocument.TextMatrix(I, GRDDOC_SEL) = ""
            grdDocument.CellFontBold = False
        End If
    Next I
    If AncDocRow < 0 Then
        MsgBox "Aucun document publié n'est associé à cette feuille." & vbCrLf & " Vous devez d'abord créer un document"
        Call Ajouter_Doc
    Else
        grdDocument.row = AncDocRow
    End If
    Faire_Doc_Click = Anc_Faire_Doc_Click

    grdFeuille.row = g_numfeuille - 1
    grdFeuille.col = ColGrdFeuilNum
    SetFocus

End Sub

Private Sub AfficheDest(v_Trait As String, v_i As Integer)
    Dim sdest As String
    Dim n As Integer
    Dim I As Integer, s As String
    Dim lstdest As String
    Dim prenom As String, nomutil As String, actif As Boolean
    Dim srvnum As Long, srvnom As String
    Dim ponum As Long, ponom As String
    Dim fctnum As Long, fctlibelle As String
    Dim grpnum As Long, grplibelle As String
    Dim lib As String
    Dim lig As Integer
    
    lbl(2).Visible = False
    cmd(CMD_AJOUT_DEST).Visible = False
    If v_Trait = "D" Then
        grdDocument.Visible = True
        cmd(CMD_AJOUT_DEST).Visible = True
        ' les destinataires sont dans grdDocument de v_i
        lstdest = grdDocument.TextMatrix(v_i, GRDDOC_LSTDEST)
        
        p_BoolFaireChkClick = False
        ChkDocPublic.Value = IIf(CBool(grdDocument.TextMatrix(v_i, GRDDOC_PUBLIC)), 1, 0)
        
        'ChkDocPublic.Value = STR_GetChamp(lstdest, "%", 1)
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
        lbl(2).Visible = True
        cmd(CMD_AJOUT_DEST).Visible = True
        n = STR_GetNbchamp(lstdest, ";")
        For I = 0 To n - 1
            s = STR_GetChamp(lstdest, ";", I)
            If left(s, 1) = "U" Then
                If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & Replace(s, "U", ""), _
                                 prenom, nomutil, actif) = P_OK Then
                    lib = prenom & "." & nomutil
                End If
            ElseIf left(s, 1) = "G" Then
                If Odbc_RecupVal("select GU_Num, GU_Nom from groupeutil where GU_Num=" & Replace(s, "G", ""), _
                                 grpnum, grplibelle) = P_OK Then
                    lib = "Groupe : " & grplibelle
                End If
            ElseIf left(s, 1) = "F" Then
                If Odbc_RecupVal("select FT_Num, FT_Libelle from fcttrav where FT_Num=" & Replace(s, "F", ""), _
                                 fctnum, fctlibelle) = P_OK Then
                    lib = "Fonction : " & fctlibelle
                End If
            ElseIf left(s, 1) = "S" Then
                If Odbc_RecupVal("select SRV_Num,SRV_Nom from service where SRV_Num=" & Replace(s, "S", ""), _
                                 srvnum, srvnom) = P_OK Then
                    lib = "Service : " & srvnom
                End If
            ElseIf left(s, 1) = "P" Then
                If Odbc_RecupVal("select PO_Num, PO_libelle from poste where PO_Num=" & Replace(s, "P", ""), _
                                 ponum, ponom) = P_OK Then
                    lib = "Poste : " & ponom
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

