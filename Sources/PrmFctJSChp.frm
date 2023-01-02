VERSION 5.00
Begin VB.Form PrmFctJSChp 
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conditions supplémentaires :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10995
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Choisir un ..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   26
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">="
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   25
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Supérieur ou égal"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<="
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Inférieur ou égal"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Frame Frm 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Expression d'une formule"
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
         Height          =   3645
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   10245
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Compris"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   28
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Inférieur ou égal"
            Top             =   960
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fenêtres"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   3120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   9495
         End
         Begin VB.Frame FrmSrvDet 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   2760
            TabIndex        =   34
            Top             =   840
            Width           =   2895
            Begin VB.CheckBox ChkSrvDet 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Détaillé ?"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   180
               Visible         =   0   'False
               Width           =   2535
            End
         End
         Begin VB.TextBox TxtValeur 
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   2280
            Width           =   5055
         End
         Begin VB.TextBox TxtOperateur 
            Height          =   375
            Left            =   2880
            TabIndex        =   32
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox TxtChamp 
            Height          =   375
            Left            =   360
            TabIndex        =   31
            Top             =   2280
            Width           =   2415
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saisir une date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   6
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   480
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   3
            Left            =   6960
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vide"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   9000
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   825
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "choix d'une valeur"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   2295
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "("
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   24
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   ")"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   23
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "choix d'un champ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   7
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ET"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OU"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   8760
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   11
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "égal à "
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "<>"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   30
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Différent de"
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Tag             =   "0"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   13
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   14
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   15
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Tag             =   "3"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   16
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Tag             =   "4"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   17
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "5"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   18
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "6"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   19
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Tag             =   "7"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   20
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Tag             =   "8"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   21
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "9"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   ","
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   22
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   ","
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   10995
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
         Left            =   10320
         Picture         =   "PrmFctJSChp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
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
         Left            =   360
         Picture         =   "PrmFctJSChp.frx":05B9
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmFctJSChp.frx":0A12
         Height          =   510
         Index           =   4
         Left            =   4440
         Picture         =   "PrmFctJSChp.frx":0E69
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Supprimer cette action"
         Top             =   195
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmFctJSChp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' index dans PrmFormEtapeChp
Private Const TXT_BUTTON_TITLE = 16

' Index des objets frm
Private Const FRM_DECLENCHEMENT = 2
Private Const FRM_FORMULE = 1

Private p_BoolSaisieDate As Boolean
Private p_boolSaisieListe As Boolean
Private p_boolSaisieListeHierar As Boolean
Private p_BoolSaisieAutre As Boolean
Private p_TypeSaisieAutre As String

Private p_TypeChamp As String
Private g_modif_val_directe As Boolean

' Index des objets cmd
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_VIDE = 2
Private Const CMD_SLASH = 3
Private Const CMD_SAISIR_DATE = 6
Private Const CMD_MOINS_ACTION = 8
Private Const CMD_SUPPRIMER = 4
Private Const CMD_OP_ET = 9
Private Const CMD_OP_OU = 10
Private Const CMD_OP_EGAL = 11
Private Const CMD_OP_DIFF = 30
Private Const CMD_CHOIX_CHP = 7
Private Const CMD_CHOIX_VAL = 5
Private Const CMD_PAR_OUV = 24
Private Const CMD_PAR_FER = 23
Private Const CMD_VIRGULE = 22
Private Const CMD_MONTER = 40
Private Const CMD_DESCENDRE = 41
Private Const CMD_OP_SUPERIEUR = 25
Private Const CMD_OP_INFERIEUR = 8
Private Const CMD_CHOIX_AUTRE = 26
Private Const CMD_CHOIX_FENETRE = 27

Private Const CMD_BOUT_0 = 12
Private Const CMD_BOUT_1 = 13
Private Const CMD_BOUT_2 = 14
Private Const CMD_BOUT_3 = 15
Private Const CMD_BOUT_4 = 16
Private Const CMD_BOUT_5 = 17
Private Const CMD_BOUT_6 = 18
Private Const CMD_BOUT_7 = 19
Private Const CMD_BOUT_8 = 20
Private Const CMD_BOUT_9 = 21

Private Const DEBUT_CMD = 12
Private Const FIN_CMD = 22

Private Const CMD_DATE_COMPRIS = 28

' Index des objets chk
Private Const CHK_MODIF = 0
Private Const CHK_LOAD = 1

' Index des objets txt
Private Const TXT_CONDF = 0
Private Const TXT_TITRE = 1
Private Const TXT_CONDPF = 2

' Index des objets frm
Private Const FRM_PROP = 0

Private g_Trait As String
Private g_numaction As Long
Private g_straction As String
Private g_sqlaction As String
Private g_numChpCnd As String
Private g_numchp As Long
Private g_numfor As Long
Private g_boolListeFen As Boolean
Private g_ListeFen As String

Private g_numetape As Long
Private g_strchp As String
Private g_TypChp As String

Private g_ConfirmerSortie As Boolean

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

Private g_retour_PrmFctJS As String

Private g_numlst As Integer
Private g_boolOper As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String

Public Function AppelFrm(ByVal v_Trait As String, _
                         ByVal v_boolListeFen As Boolean, _
                         ByVal v_numaction As Long, _
                         ByVal v_straction As String, _
                         ByVal v_numChpCnd As String, _
                         ByVal v_numfor As Long)

    g_numaction = v_numaction
    g_straction = v_straction
    g_numChpCnd = v_numChpCnd
    g_numfor = v_numfor
    g_boolListeFen = v_boolListeFen
    If v_boolListeFen Then
        g_ListeFen = STR_GetChamp(v_straction, "¤", 3)
        If g_ListeFen = "" Then
            g_ListeFen = "*"
        End If
    End If
    
    g_ConfirmerSortie = True
    If v_Trait = "AjoutSpecialPrem" Then
        v_Trait = "Ajout"
        g_ConfirmerSortie = False
    End If
    g_Trait = v_Trait
    
    Me.Show 1
    
    If g_Trait = "Ajout" Then
        AppelFrm = g_retour_PrmFctJS
    End If
    If g_Trait = "Modif" Then
        AppelFrm = g_retour_PrmFctJS
    End If
    
End Function

Private Function ChoisirChamp(ByVal V_ChpNum As String) As String
                                
    Dim sql As String, rs As rdoResultset
    Dim n As Integer, numlst As Integer
    Dim sNumLst As String
    Dim strG As String, strD As String, s As String
    Dim strGPF As String, strDPF As String
    Dim ancnumchp As Integer, i As Integer
    Dim LibType As String, fornums As String, sin As String
    Dim nbiter As Integer
    Dim àGarder As Boolean
    
    On Error Resume Next
    ancnumchp = cmd(CMD_CHOIX_CHP).tag
    cmd(CMD_CHOIX_VAL).Visible = False
    cmd(CMD_CHOIX_VAL).tag = 0
    cmd(CMD_CHOIX_CHP).tag = 0
    
    Call CL_Init
    Call CL_InitTitreHelp("Choix du Champ de formulaire", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    If ancnumchp > 0 Then
        sql = "select * from formetapechp where forec_num=" & ancnumchp
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        Call CL_AddLigne(rs("FOREC_Label").Value, rs("FOREC_Num").Value, "", True)
    End If
    If p_derchamp > 0 Then
        sql = "select * from formetapechp where forec_num=" & p_derchamp
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        Call CL_AddLigne(rs("FOREC_Label").Value, rs("FOREC_Num").Value, "", True)
    End If
    sql = "select ff_fornums from filtreform where ff_num = " & g_numfor
    Call Odbc_RecupVal(sql, fornums)
    sin = "("
    n = STR_GetNbchamp(fornums, "*")
    For i = 1 To n - 1
        s = STR_GetChamp(fornums, "*", i)
        If i > 1 Then
            sin = sin + ","
        End If
        sin = sin + s
    Next i
    sin = sin + ")"
    sql = "select * from formetapechp where forec_fornum in " & sin & " and forec_Type <> 'BUTTON' order by forec_fornum,forec_numetape,forec_ordre"
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    While Not rs.EOF
        àGarder = False
        If rs("forec_num") <> ancnumchp Then
            'Debug.Print rs("forec_label") & " " & rs("forec_formule")
            'LibType = rs("FOREC_Nom").Value & " : " & rs("FOREC_Label").Value & " (" & rs("FOREC_Type")
            LibType = rs("FOREC_Type")
            If rs("FOREC_Type") = "TEXT" Then
                'MsgBox rs("FOREC_valeurs_possibles")
                If rs("forec_fctvalid") <> "" Then
                    If rs("forec_fctvalid") = "%NUMUTIL" Then
                        LibType = LibType & " : Utilisateur"
                        àGarder = True
                    ElseIf InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
                        LibType = LibType & " : Date"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%NUMSERVICE" Then
                        LibType = LibType & " : Service"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%NUMFCT" Then
                        LibType = LibType & " : Fonction"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%NUM" Or rs("forec_fctvalid") = "%ENTIER" Then
                        LibType = LibType & " : Entier"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%HEURE" Then
                        LibType = LibType & " : Heure"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%MTT" Then
                        LibType = LibType & " : Nombre Décimal"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%LISTUTIL" Then
                        LibType = LibType & " : Liste d'utilisateurs"
                        àGarder = True
                    ElseIf rs("forec_fctvalid") = "%MAIL" Then
                        LibType = LibType & " : Adresse Mail"
                        àGarder = True
                    ElseIf Mid(rs("forec_fctvalid"), 1, 9) = "%NUMCAUSE" Then
                        LibType = LibType & " : Causes"
                        àGarder = True
                    ElseIf Mid(rs("forec_fctvalid"), 1, 10) = "%NUMCONSEQ" Then
                        LibType = LibType & " : Conséquences"
                        àGarder = True
                    ElseIf Mid(rs("forec_fctvalid"), 1, 10) = "%TELEPHONE" Then
                        LibType = LibType & " : Téléphone"
                        àGarder = True
                    ElseIf Mid(rs("forec_fctvalid"), 1, 3) = "%PJ" Then
                        àGarder = False
                    Else
                        MsgBox rs("forec_fctvalid")
                        LibType = LibType & " : Type indéterminé"
                        àGarder = False
                    End If
                ElseIf Mid(rs("forec_formule"), 1, 9) = "=calculer" Then
                    LibType = LibType & " : Champ calculé"
                    àGarder = True
                End If
            ElseIf rs("FOREC_Type") = "HIERARCHIE" Then
                LibType = "Liste hiér."
                àGarder = True
            ElseIf rs("FOREC_Type") = "TEXTAREA" Then
                LibType = "Texte illimité"
                àGarder = True
            ElseIf rs("FOREC_Type") = "SELECT" Then
                LibType = "Liste"
                àGarder = True
            ElseIf rs("FOREC_Type") = "RADIO" Then
                LibType = "Boutons Radio"
                àGarder = True
            ElseIf rs("FOREC_Type") = "CHECK" Then
                LibType = "Cases à cocher"
                àGarder = True
            End If
            If àGarder Then
                If rs("FOREC_Num").Value = V_ChpNum Then
                    Call CL_AddLigne(rs("FOREC_Nom").Value & vbTab & rs("FOREC_Label").Value & vbTab & LibType, rs("FOREC_Num").Value, "", True)
                Else
                    Call CL_AddLigne(rs("FOREC_Nom").Value & vbTab & rs("FOREC_Label").Value & vbTab & LibType, rs("FOREC_Num").Value, "", False)
                End If
            End If
        End If
        rs.MoveNext
    Wend
    
    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    
    ' Quitter
    If CL_liste.retour = 1 Then
        Exit Function
    End If
    
    ' Choix d'un champ
    ChoisirChamp = FctNomChp(CL_liste.lignes(CL_liste.pointeur).num)
    ' chercher si c'est une liste de valeurs
    sql = "select * from formetapechp where forec_num = " & CL_liste.lignes(CL_liste.pointeur).num
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    sNumLst = ""
    numlst = -1
    g_numlst = 0
    p_TypeChamp = ""
    If Not rs.EOF Then
        If InStr(rs("forec_fctvalid"), "%NUM") = 0 Then
            sNumLst = rs("forec_valeurs_possibles").Value
            If sNumLst <> "" Then
                g_numlst = rs("forec_valeurs_possibles").Value
            End If
        End If
        p_BoolSaisieDate = False
        p_boolSaisieListe = False
        p_boolSaisieListeHierar = False
        p_BoolSaisieAutre = False
        p_TypeSaisieAutre = ""
        
        Me.TxtChamp.Text = ""
        Me.TxtChamp.tag = ""
        Me.TxtOperateur.Text = ""
        Me.TxtOperateur.tag = ""
        Me.TxtValeur.Text = ""
        Me.TxtValeur.tag = ""
        cmd(CMD_CHOIX_VAL).tag = 0
        cmd(CMD_CHOIX_AUTRE).tag = 0
        Me.FrmSrvDet.Visible = False
        
        If InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
            p_BoolSaisieDate = True
            Me.cmd(CMD_SAISIR_DATE).Visible = True
            Me.cmd(CMD_CHOIX_VAL).Visible = False
            Me.cmd(CMD_CHOIX_AUTRE).Visible = False
            Me.cmd(CMD_OP_EGAL).Visible = True
            Me.cmd(CMD_OP_DIFF).Visible = True
            Me.cmd(CMD_OP_SUPERIEUR).Visible = True
            Me.cmd(CMD_OP_INFERIEUR).Visible = True
            Me.cmd(CMD_DATE_COMPRIS).Visible = True
        ElseIf rs("forec_fctvalid") = "%NUMFCT" Or rs("forec_fctvalid") = "%NUMSERVICE" Then
            Me.cmd(CMD_SAISIR_DATE).Visible = False
            Me.cmd(CMD_CHOIX_VAL).Visible = False
            Me.cmd(CMD_CHOIX_AUTRE).Visible = True
            Me.cmd(CMD_OP_EGAL).Visible = True
            Me.cmd(CMD_OP_DIFF).Visible = False
            Me.cmd(CMD_OP_SUPERIEUR).Visible = False
            Me.cmd(CMD_OP_INFERIEUR).Visible = False
            Me.cmd(CMD_DATE_COMPRIS).Visible = False
            p_TypeSaisieAutre = Replace(rs("forec_fctvalid"), "%", "")
            If p_TypeSaisieAutre = "NUMSERVICE" Then
                cmd(CMD_CHOIX_AUTRE).Caption = "Choisir un Service"
                p_TypeChamp = "SRV"
                Me.FrmSrvDet.Visible = True
                Me.ChkSrvDet.Visible = True
            ElseIf p_TypeSaisieAutre = "NUMFCT" Then
                cmd(CMD_CHOIX_AUTRE).Caption = "Choisir une Fonction"
            End If
        Else
            Me.cmd(CMD_SAISIR_DATE).Visible = False
            Me.cmd(CMD_CHOIX_AUTRE).Visible = False
            Me.cmd(CMD_DATE_COMPRIS).Visible = False
            If rs("forec_type").Value = "CHECK" Or rs("forec_type").Value = "RADIO" Or rs("forec_type").Value = "SELECT" Then
                p_boolSaisieListe = True
                Me.cmd(CMD_CHOIX_VAL).Visible = True
                Me.cmd(CMD_OP_EGAL).Visible = True
                Me.cmd(CMD_OP_DIFF).Visible = True
                Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                Me.cmd(CMD_OP_INFERIEUR).Visible = False
            ElseIf rs("forec_type").Value = "HIERARCHIE" Then
                p_boolSaisieListeHierar = True
                Me.FrmSrvDet.Visible = True
                Me.ChkSrvDet.Visible = True
                Me.cmd(CMD_CHOIX_VAL).Visible = True
                Me.cmd(CMD_OP_EGAL).Visible = True
                Me.cmd(CMD_OP_DIFF).Visible = True
                Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                Me.cmd(CMD_OP_INFERIEUR).Visible = False
            Else
                Me.cmd(CMD_CHOIX_VAL).Visible = False
                Me.cmd(CMD_OP_EGAL).Visible = True
                Me.cmd(CMD_OP_DIFF).Visible = True
                Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                Me.cmd(CMD_OP_INFERIEUR).Visible = True
                g_modif_val_directe = True
            End If
        End If
    End If
    
    Me.TxtChamp.tag = rs("FOREC_Num")
    '
    cmd(CMD_CHOIX_CHP).tag = CL_liste.lignes(CL_liste.pointeur).num
    If g_numlst > 0 Then
        cmd(CMD_CHOIX_VAL).Visible = True
        cmd(CMD_CHOIX_VAL).tag = g_numlst
        cmd(CMD_CHOIX_AUTRE).tag = 0
    Else
        cmd(CMD_CHOIX_VAL).tag = 0
        cmd(CMD_CHOIX_AUTRE).tag = 0
    End If
    
    ChoisirChamp = cmd(CMD_CHOIX_CHP).tag
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    ChoisirChamp = P_ERREUR
    Exit Function
    
End Function

Private Function ChoisirFenetres(ByVal V_ListeFen As String) As String
                                
    Dim sql As String, rs As rdoResultset
    Dim sret As String
    Dim n As Integer
    Dim i As Integer
    Dim bTous As Boolean
    
    Call CL_Init
    Call CL_InitTitreHelp("Choix des Feuilles", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)
    
    If V_ListeFen = "*" Then bTous = True
    If p_bool_tbl_fenExcel Then
        For i = 1 To UBound(tbl_fenExcel)
            If bTous Or InStr(V_ListeFen, tbl_fenExcel(i).FenNum) > 0 Then
                Call CL_AddLigne(tbl_fenExcel(i).FenNom, tbl_fenExcel(i).FenNum, tbl_fenExcel(i).FenNum, True)
            Else
                Call CL_AddLigne(tbl_fenExcel(i).FenNom, tbl_fenExcel(i).FenNum, tbl_fenExcel(i).FenNum, False)
            End If
        Next i
    End If
    
    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    
    ' Quitter
    If CL_liste.retour = 1 Then
        Exit Function
    End If
    bTous = True
    For n = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(n).selected Then
            sret = sret & CL_liste.lignes(n).tag & ";"
        Else
            bTous = False
        End If
    Next n
    If bTous Then
        sret = "*"
    End If
    
    ChoisirFenetres = sret
    
End Function


Private Function choisir_valeur_autre() As String

    Dim sql As String, s As String
    Dim numchp As Long, numlst As Long
    Dim rs As rdoResultset
    Dim strNom As String
    Dim i As Integer
    Dim StrFct As String, libfct As String
    Dim sCond As String, stype As String, fctvalid As String
    Dim ret As String, iret As String
    
    choisir_valeur_autre = ""
    'MsgBox "ici"
            
    sql = "select FOREC_Label, FOREC_Type, FOREC_FctValid from FormEtapeChp" _
        & " where FOREC_Num=" & Me.TxtChamp.tag
    If Odbc_RecupVal(sql, sCond, stype, fctvalid) = P_ERREUR Then
        MsgBox "Erreur SQL " & sql
    End If
    
    
    If fctvalid = "%NUMSERVICE" Then
        p_TypeChamp = "SRV"
        ret = PrmFormatChp.ChoisirService(Replace(Me.TxtValeur.tag, "_DET", ""))
        If ret <> "" And ret <> "0" Then
            For i = 0 To UBound(CL_liste.lignes)
                s = CL_liste.lignes(i).texte
                ' recupérer le dernier
                iret = Mid$(STR_GetChamp(s, ";", STR_GetNbchamp(s, ";") - 1), 2)
                Call P_RecupSrvNom(iret, strNom)
                'Me.LabValChp.Caption = strNom
                choisir_valeur_autre = IIf(choisir_valeur_autre = "", "S" & iret & ";", choisir_valeur_autre & "S" & iret & ";")
                'Me.LabValChp.Visible = True
            Next i
        Else
            'Me.LabValChp.Visible = True
            'Me.LabValChp.Caption = "Tous les Services"
            choisir_valeur_autre = 0
        End If
    ElseIf fctvalid = "%NUMFCT" Then
        p_TypeChamp = "FCT"
        StrFct = Me.TxtValeur.tag
        ret = choisir_fonction(StrFct, StrFct, libfct)
        If ret <> "" Then
            choisir_valeur_autre = StrFct
        Else
            'Me.LabValChp.Visible = True
            'Me.LabValChp.Caption = "Tous les Services"
            choisir_valeur_autre = 0
        End If
    End If

End Function

Private Function choisir_fonction(ByVal v_numfct As String, ByRef r_strFct As String, ByRef r_LibFct As String) As Integer

    Dim sret As String, sql As String
    Dim n As Integer
    Dim bDeja As Boolean
    Dim LeNum As String
    Dim nofct As Long
    Dim rs As rdoResultset
    Dim sep As String
    Dim SepF As String
    Call FRM_ResizeForm(Me, 0, 0)

lab_affiche:
    Call CL_Init
    Call CL_InitMultiSelect(True, True)
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_fonction = P_ERREUR
        Exit Function
    End If
    ' Celles déjà cochée
    While Not rs.EOF
        For n = 0 To STR_GetNbchamp(v_numfct, ";")
            LeNum = STR_GetChamp(v_numfct, ";", n)
            If rs("FT_Num").Value = LeNum Then
                Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", True)
                Exit For
            End If
        Next n
        rs.MoveNext
    Wend
    ' Les autres
    rs.MoveFirst
    While Not rs.EOF
        bDeja = False
        For n = 0 To STR_GetNbchamp(v_numfct, ";")
            If rs("FT_Num").Value = STR_GetChamp(v_numfct, ";", n) Then
                bDeja = True
                Exit For
            End If
        Next n
        If Not bDeja Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Liste des fonctions", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_c_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        choisir_fonction = P_NON
        Exit Function
    End If
    
    sep = ""
    r_strFct = ""
    r_LibFct = ""
    For n = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(n).selected Then
            'choisir_valeur = choisir_valeur & Sep & "VAL:" & CL_liste.lignes(I).num & "[" & CL_liste.lignes(I).texte & "]"
            r_strFct = r_strFct & sep & CL_liste.lignes(n).num
            r_LibFct = r_LibFct & SepF & CL_liste.lignes(n).texte
            SepF = " - "
            sep = ";"
        End If
    Next n
    
    p_TypeChamp = "FCT"
    Me.TxtOperateur.Text = "Parmi"
    Me.TxtValeur.tag = r_strFct
    Me.TxtValeur.Text = r_LibFct
    
    choisir_fonction = P_OUI

End Function

Private Function choisir_valeur() As String

    Dim sql As String, s As String
    Dim numchp As Long, numlst As Long
    Dim rs As rdoResultset
    Dim sep2 As String
    Dim bMettre As Boolean
    Dim i As Integer, sep As String
    
    choisir_valeur = ""
    
    Call CL_Init
    
    Call CL_InitMultiSelect(True, True)
    
    numlst = cmd(CMD_CHOIX_VAL).tag
    
    sql = "select VC_Num, VC_Lib from ValChp" _
        & " where VC_LVCNum=" & numlst
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Aucune valeur n'a été trouvée.", vbInformation + vbOKOnly, "")
        Exit Function
    End If
    ' Ceux qui sont selected
    While Not rs.EOF
        For i = 0 To STR_GetNbchamp(Me.TxtValeur.tag, ";")
            If STR_GetChamp(Me.TxtValeur.tag, ";", i) = rs("VC_Num").Value Then
                Call CL_AddLigne(rs("VC_Lib").Value, rs("VC_Num").Value, "", True)
            End If
        Next i
        rs.MoveNext
    Wend
    ' Ceux qui ne sont pas selected
    rs.MoveFirst
    While Not rs.EOF
        bMettre = True
        For i = 0 To STR_GetNbchamp(Me.TxtValeur.tag, ";")
            If STR_GetChamp(Me.TxtValeur.tag, ";", i) = rs("VC_Num").Value Then
                bMettre = False
                Exit For
            End If
        Next i
        If bMettre Then
            Call CL_AddLigne(rs("VC_Lib").Value, rs("VC_Num").Value, "", False)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Liste des valeurs", p_chemin_appli + "\help\kalidoc.chm" & ";" & "form_etape.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 0 Then
        choisir_valeur = ""
        sep = ""
        sep2 = ""
        Me.TxtValeur.tag = ""
        Me.TxtValeur.Text = ""
        For i = 0 To UBound(CL_liste.lignes)
            If CL_liste.lignes(i).selected Then
                choisir_valeur = choisir_valeur & sep & "VAL:" & CL_liste.lignes(i).num & "[" & CL_liste.lignes(i).texte & "]"
                Me.TxtValeur.tag = Me.TxtValeur.tag & sep & CL_liste.lignes(i).num
                Me.TxtValeur.Text = Me.TxtValeur.Text & sep2 & CL_liste.lignes(i).texte
                sep2 = " - "
                sep = ";"
            End If
        Next i
    End If

End Function

Private Function choisir_valeur_hierar()

    Dim sval As String, sep As String, sep2 As String, s As String
    Dim n As Integer, i As Integer, n2 As Integer
    Dim numlst As Long
    
    choisir_valeur_hierar = ""
    
    Call CL_Init
    
    Call CL_InitMultiSelect(True, True)
    
    numlst = cmd(CMD_CHOIX_VAL).tag
    
    Call CL_AddLigne("<Non renseigné>", 0, 0, IIf(InStr(TxtValeur.tag, " ;") > 0, True, False))
    sval = ""
    n = STR_GetNbchamp(TxtValeur.tag, ";")
    For i = 0 To n - 1
        s = STR_GetChamp(TxtValeur.tag, ";", i)
        sval = sval & "M" & s & ";"
    Next i
    Call ajouter_hierar_fils(-numlst, sval, 0)
    If UBound(CL_liste.lignes()) = 1 Then
        Call MsgBox("Aucune valeur n'a été trouvée.", vbInformation + vbOKOnly, "")
        Exit Function
    End If
    Call CL_InitTitreHelp("Liste des valeurs", p_chemin_appli + "\help\kalidoc.chm" & ";" & "form_etape.htm")
    If UBound(CL_liste.lignes) < 10 Then
        n = UBound(CL_liste.lignes) + 1
    Else
        n = 10
    End If
    Call CL_InitTaille(0, -n)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        Exit Function
    End If
    
    If CL_liste.retour = 0 Then
        choisir_valeur_hierar = ""
        sep = ""
        sep2 = ""
        Me.TxtValeur.tag = ""
        Me.TxtValeur.Text = ""
        For i = 0 To UBound(CL_liste.lignes)
            If CL_liste.lignes(i).selected Then
                choisir_valeur_hierar = choisir_valeur_hierar & sep & "VAL:" & CL_liste.lignes(i).num & "[" & CL_liste.lignes(i).texte & "]"
                Me.TxtValeur.tag = Me.TxtValeur.tag & sep & CL_liste.lignes(i).num
                Me.TxtValeur.Text = Me.TxtValeur.Text & sep2 & CL_liste.lignes(i).texte
                sep2 = " - "
                sep = ";"
            End If
        Next i
    End If

End Function

Private Function FctNomChp(V_ChpNum)
    
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from formetapechp where forec_num = " & V_ChpNum
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    
    If Not rs.EOF Then
        FctNomChp = rs("forec_nom").Value
    Else
        FctNomChp = V_ChpNum
    End If

End Function

Private Function FctOnclick_Ou_Onchange(v_numchp As Long)
    
    Dim sql As String
    Dim rs As rdoResultset
    
    sql = "select FOREC_type from FormEtapeChp" _
        & " where FOREC_Num=" & v_numchp
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If Not rs.EOF Then
        If rs("FOREC_type") = "TEXT" Then
            FctOnclick_Ou_Onchange = "onchange"
        Else
            FctOnclick_Ou_Onchange = "onclick"
        End If
    End If

End Function

Private Function FaitListeFenetres(v_listefenetres)
    Dim rs As rdoResultset, sql As String
    Dim s As String
    Dim i As Integer
    Dim nb As Integer
    
    nb = 0
    If v_listefenetres = "*" Then
        FaitListeFenetres = "Toutes les Feuilles"
    Else
        For i = 0 To STR_GetNbchamp(v_listefenetres, ";")
            s = STR_GetChamp(v_listefenetres, ";", i)
            If s <> "" Then
                If p_bool_tbl_fenExcel Then
                    On Error Resume Next
                    FaitListeFenetres = FaitListeFenetres & IIf(FaitListeFenetres <> "", " + ", "") & tbl_fenExcel(s).FenNom
                End If
            End If
        Next i
        FaitListeFenetres = "Feuilles : " & FaitListeFenetres
    End If
End Function

Private Sub initialiser()

    Dim i As Integer
    Dim op As String
    Dim sql As String
    Dim rs As rdoResultset
    Dim n As Integer
    Dim sChamp As String
    Dim sCond As String, stype As String
    Dim NumChamp As String
    Dim sValeur As String
    Dim sep2 As String
    Dim sep As String
    Dim sG As String, sD As String
    Dim strG As String, strD As String
    Dim s As String
    Dim numServ As String
    Dim Forec_Formule As String
    Dim nomFct As String
    Dim LaStr As String
    Dim sNumLst As String
    Dim slstServ As String
    Dim numlst As Integer
    Dim Label As String, fctvalid As String
    Dim BoolDetail As Boolean
    Dim nomServ As String
    Dim nomServs As String
            
    ' lire l'action
    g_mode_saisie = False
    
    cmd(CMD_CHOIX_VAL).Visible = False
    cmd(CMD_CHOIX_VAL).tag = 0
    cmd(CMD_CHOIX_CHP).tag = 0
    Me.FrmSrvDet.Visible = False
    
    cmd(CMD_CHOIX_FENETRE).Visible = g_boolListeFen
    cmd(CMD_CHOIX_FENETRE).tag = g_ListeFen
    
    If g_boolListeFen Then
        cmd(CMD_CHOIX_FENETRE).Visible = True
        If g_ListeFen = "*" Then
            cmd(CMD_CHOIX_FENETRE).Caption = "Toutes les fenêtres"
            cmd(CMD_CHOIX_FENETRE).BackColor = &HE0E0E0
        Else
            cmd(CMD_CHOIX_FENETRE).Caption = FaitListeFenetres(g_ListeFen)
            cmd(CMD_CHOIX_FENETRE).BackColor = &H8080FF
        End If
    Else
        cmd(CMD_CHOIX_FENETRE).Visible = False
    End If
        
    If g_Trait = "Ajout" Then
        ' mettre la condition
        'txt(TXT_CONDF).Text = ""
        
        If g_numChpCnd <> "" And g_numChpCnd <> 0 Then
            ' chercher si c'est une liste de valeurs
            sql = "select * from formetapechp where forec_num = " & g_numChpCnd
            Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
            sNumLst = ""
            numlst = -1
            g_numlst = 0
            If Not rs.EOF Then
                sNumLst = rs("forec_valeurs_possibles").Value
                If sNumLst <> "" Then
                    g_numlst = rs("forec_valeurs_possibles").Value
                End If
                p_BoolSaisieDate = False
                p_boolSaisieListe = False
                p_boolSaisieListeHierar = False
        
                If InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
                    g_modif_val_directe = True
                    p_BoolSaisieDate = True
                    Me.cmd(CMD_SAISIR_DATE).Visible = True
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_OP_EGAL).Visible = True
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                    Me.cmd(CMD_OP_INFERIEUR).Visible = True
                Else
                    Me.cmd(CMD_SAISIR_DATE).Visible = False
                    If rs("forec_type").Value = "CHECK" Or rs("forec_type").Value = "RADIO" Or rs("forec_type").Value = "SELECT" Then
                        g_modif_val_directe = False
                        p_boolSaisieListe = True
                        p_boolSaisieListeHierar = False
                        Me.cmd(CMD_CHOIX_VAL).Visible = True
                        Me.cmd(CMD_OP_EGAL).Visible = True
                        Me.cmd(CMD_OP_DIFF).Visible = True
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                        Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    ElseIf rs("forec_type").Value = "HIERARCHIE" Then
                        g_modif_val_directe = False
                        p_boolSaisieListe = False
                        p_boolSaisieListeHierar = True
                        Me.cmd(CMD_CHOIX_VAL).Visible = False
                        Me.cmd(CMD_OP_EGAL).Visible = True
                        Me.cmd(CMD_OP_DIFF).Visible = True
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                        Me.cmd(CMD_OP_INFERIEUR).Visible = True
                    End If
                End If
            End If
            '
            cmd(CMD_CHOIX_CHP).tag = g_numChpCnd
            If g_numlst > 0 Then
                cmd(CMD_CHOIX_VAL).Visible = True
                cmd(CMD_CHOIX_VAL).tag = g_numlst
            Else
                'cmd(CMD_CHOIX_VAL).visible = False
                cmd(CMD_CHOIX_VAL).tag = 0
            End If
            'txt(TXT_CONDF).Text = "CHP:" & g_numChpCnd & ":" & rs("forec_nom") & " "
            'txt(TXT_CONDF).SelStart = 0
            'txt(TXT_CONDF).SelLength = Len(txt(TXT_CONDF).Text)
            'txt(TXT_CONDF).SetFocus
        End If
    Else
        ' lire le champ : mettre le titre et la formule
        Me.cmd(CMD_DATE_COMPRIS).Visible = False
        If g_Trait = "BTN" Then
            'txt(TXT_TITRE).Text = PrmFormEtapeChp.txt(TXT_BUTTON_TITLE).Text
            frm(FRM_DECLENCHEMENT).Visible = False
            frm(FRM_FORMULE).Visible = False
        End If
        
        ' Découper en 3
        If g_straction <> "" Then
            sChamp = STR_GetChamp(g_straction, "¤", 0)
            Me.TxtChamp.Text = STR_GetChamp(g_straction, "¤", 0)
            NumChamp = STR_GetChamp(STR_GetChamp(g_straction, "¤", 0), ":", 1)
            If val(g_numChpCnd) = 0 Then
                g_numChpCnd = NumChamp
            End If
            sql = "select FOREC_Formule, FOREC_FctValid, FOREC_Label, FOREC_Type, FOREC_Valeurs_Possibles from FormEtapeChp" _
                & " where FOREC_Num=" & NumChamp
            If Odbc_RecupVal(sql, Forec_Formule, fctvalid, Label, stype, sNumLst) = P_ERREUR Then
                MsgBox "Erreur SQL " & sql
            End If
            Me.TxtChamp.Text = Label & " (" & STR_GetChamp(STR_GetChamp(g_straction, "¤", 0), ":", 2) & ")"
            Me.TxtChamp.tag = NumChamp
        
            sValeur = STR_GetChamp(g_straction, "¤", 2)
            sValeur = Replace(sValeur, "VAL:", "")
            sValeur = Replace(sValeur, stype & ":", "")
            If stype = "RADIO" Or stype = "SELECT" Or stype = "CHECK" Then
                If sValeur = "" Then
                    sValeur = ""
                Else
                    Me.TxtValeur.tag = ""
                    ' MODIF LN 27/06/11
                    For n = 0 To STR_GetNbchamp(sValeur, ";") - 1
                        s = STR_GetChamp(sValeur, ";", n)
                        If s <> "" And s <> "<NR>" And s <> "<R>" Then
                            sql = "select VC_Lib from ValChp" _
                                & " where VC_Num=" & s
                            If Odbc_RecupVal(sql, s) = P_ERREUR Then
                                MsgBox "Erreur SQL " & sql
                            End If
                            Me.TxtValeur.Text = Me.TxtValeur.Text & sep2 & s
                            Me.TxtValeur.tag = Me.TxtValeur.tag & sep & STR_GetChamp(sValeur, ";", n)
                            sep = ";"
                            sep2 = " - "
                        Else
                            Me.TxtValeur.Text = "[VIDE]"
                            Me.TxtValeur.tag = "<NR>"
                            Exit For
                        End If
                    Next n
                End If
                'scond = scond + s
            ElseIf stype = "HIERARCHIE" Then
                If sValeur = "" Then
                    sValeur = ""
                Else
                    Me.TxtValeur.tag = ""
                    BoolDetail = IIf(InStr(sValeur, "_DET") > 0, True, False)
                    sValeur = Replace(sValeur, "_DET", "")
                    sValeur = Replace(sValeur, "_", "")
                    op = STR_GetChamp(g_straction, "¤", 1)
                    If sValeur = "0" Then
                        Me.TxtValeur.Text = "<non renseigné>"
                        Me.TxtValeur.tag = "0"
                        Me.ChkSrvDet.Visible = False
                        Me.FrmSrvDet.Visible = False
                        GoTo LabFinHier
                    Else
                        For n = 0 To STR_GetNbchamp(sValeur, ";")
                            s = STR_GetChamp(sValeur, ";", n)
                            If s <> "" Then
                                sql = "select HVC_Nom from HierarValChp" _
                                    & " where HVC_Num=" & s
                                If Odbc_RecupVal(sql, s) = P_ERREUR Then
                                    MsgBox "Erreur SQL " & sql
                                End If
                                Me.TxtValeur.Text = Me.TxtValeur.Text & sep2 & s
                                Me.TxtValeur.tag = Me.TxtValeur.tag & sep & STR_GetChamp(sValeur, ";", n)
                                sep = ";"
                                sep2 = " - "
                            End If
                        Next n
                    End If
                End If
                Me.ChkSrvDet.Visible = True
                Me.FrmSrvDet.Visible = True
                If BoolDetail Then
                    Me.TxtOperateur.Text = "Fait partie de"
                    Me.TxtValeur.Text = Me.TxtValeur.Text & " (D)"
                    Me.ChkSrvDet.Value = 1
                    Me.ChkSrvDet.Visible = True
                    Me.FrmSrvDet.Visible = True
                Else
                    Me.TxtOperateur.Text = "Strictement égal à"
                    Me.ChkSrvDet.Value = 0
                End If
                If op = "OP:NE" And stype = "HIERARCHIE" Then
                    BoolDetail = True
                    Me.TxtOperateur.Text = "Différent de"
                    Me.ChkSrvDet.Value = 1
                    Me.ChkSrvDet.Visible = False
                    Me.FrmSrvDet.Visible = False
                End If
LabFinHier:
            Else
                If fctvalid = "%NUMSERVICE" Then
                    'MsgBox sValeur
                    sValeur = Replace(sValeur, "NUMSERVICE:", "")
                    BoolDetail = IIf(InStr(sValeur, "_DET") > 0, True, False)
                    numServ = Replace(sValeur, "_DET", "")
                    numServ = Replace(numServ, "_", "")
                    numServ = Replace(numServ, "S", "")
                    If numServ = "<NR>" Then
                        nomServs = "<non renseigné>"
                        slstServ = "<NR>"
                        GoTo suite
                    End If
                    nomServs = ""
                    slstServ = ""
                    For i = 0 To STR_GetNbchamp(numServ, ";")
                        s = STR_GetChamp(numServ, ";", i)
                        If s <> "" Then
                            Call P_RecupSrvNom(s, nomServ)
                            nomServs = nomServs & IIf(nomServs = "", "", " Ou ") & nomServ
                            slstServ = slstServ & "S" & s & ";"
                        End If
                    Next i
                    slstServ = Mid(slstServ, 1, Len(slstServ) - 1)
                    slstServ = slstServ & IIf(BoolDetail, "_DET", "")
suite:
                    Me.TxtValeur.Text = nomServs & IIf(BoolDetail, " (D)", "")
                    Me.TxtValeur.tag = slstServ
                    Me.TxtOperateur.tag = "="
                    If BoolDetail Then
                        Me.TxtOperateur.Text = "Fait partie de"
                    Else
                        Me.TxtOperateur.Text = "Strictement égal à"
                    End If
                    Me.ChkSrvDet.Visible = True
                    Me.FrmSrvDet.Visible = True
                    If BoolDetail Then
                        Me.ChkSrvDet.Value = 1
                        Me.ChkSrvDet.Visible = True
                        Me.FrmSrvDet.Visible = True
                    Else
                        Me.ChkSrvDet.Value = 0
                    End If
                    Me.cmd(CMD_SAISIR_DATE).Visible = False
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_CHOIX_AUTRE).Visible = True
                    Me.cmd(CMD_OP_EGAL).Visible = True
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                    Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    cmd(CMD_CHOIX_AUTRE).Caption = "Choisir un Service"
                    p_TypeSaisieAutre = "NUMSERVICE"
                ElseIf fctvalid = "%NUMFCT" Then
                    'MsgBox sValeur
                    op = ""
                    For n = 0 To STR_GetNbchamp(sValeur, ";")
                        s = STR_GetChamp(sValeur, ";", n)
                        If s <> "" Then
                            Call P_RecupNomFonction(val(s), nomFct)
                            Me.TxtValeur.Text = Me.TxtValeur.Text & op & nomFct
                            op = " ou "
                        End If
                    Next n
                    p_TypeChamp = "FCT"
                    Me.TxtValeur.tag = sValeur
                    Me.cmd(CMD_SAISIR_DATE).Visible = False
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_CHOIX_AUTRE).Visible = True
                    Me.cmd(CMD_OP_EGAL).Visible = True
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                    Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    cmd(CMD_CHOIX_AUTRE).Caption = "Choisir une Fonction"
                    p_TypeSaisieAutre = "NUMFCT"
                ElseIf InStr(fctvalid, "%DATE") > 0 Then
                    'MsgBox sValeur
                    'Me.TxtValeur.Text = s
                    'p_TypeChamp = "FCT"
                    g_modif_val_directe = True
                    Me.TxtValeur.tag = sValeur
                    Me.TxtValeur.Text = sValeur
                    Me.TxtValeur.Enabled = False
                    op = STR_GetChamp(g_straction, "¤", 1)
                    op = Replace(op, "OP:", "")
                    Me.cmd(CMD_SAISIR_DATE).Visible = True
                    Me.cmd(CMD_SAISIR_DATE).Visible = True
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_CHOIX_AUTRE).Visible = False
                    Me.cmd(CMD_OP_EGAL).Visible = False
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                    Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    Me.cmd(CMD_DATE_COMPRIS).Visible = False
                    If op = "SU" Then
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                    ElseIf op = "IN" Then
                        Me.cmd(CMD_OP_INFERIEUR).Visible = True
                    ElseIf op = "EG" Or op = "=" Then
                        Me.cmd(CMD_OP_EGAL).Visible = True
                    ElseIf op = "NE" Then
                        Me.cmd(CMD_OP_DIFF).Visible = True
                    ElseIf op = "COMPRIS" Then
                        Me.cmd(CMD_DATE_COMPRIS).Visible = True
                    End If
                    'cmd(CMD_CHOIX_AUTRE).Caption = "Choisir une Fonction"
                    'p_TypeSaisieAutre = "NUMFCT"
                ElseIf Mid(Forec_Formule, 1, 9) = "=calculer" Or InStr(fctvalid, "%NUM") > 0 Or InStr(fctvalid, "%ENTIER") > 0 Then
                    'MsgBox sValeur
                    'Me.TxtValeur.Text = s
                    'p_TypeChamp = "FCT"
                    Me.TxtValeur.tag = sValeur
                    Me.TxtValeur.Text = sValeur
                    Me.TxtValeur.Enabled = False
                    op = STR_GetChamp(g_straction, "¤", 1)
                    op = Replace(op, "OP:", "")
                    Me.cmd(CMD_SAISIR_DATE).Visible = False
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_CHOIX_AUTRE).Visible = False
                    Me.cmd(CMD_OP_EGAL).Visible = False
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                    Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    If op = "SU" Then
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                    ElseIf op = "IN" Then
                        Me.cmd(CMD_OP_INFERIEUR).Visible = True
                    ElseIf op = "EG" Or op = "=" Then
                        Me.cmd(CMD_OP_EGAL).Visible = True
                    ElseIf op = "NE" Then
                        Me.cmd(CMD_OP_DIFF).Visible = True
                    End If
                    'cmd(CMD_CHOIX_AUTRE).Caption = "Choisir une Fonction"
                    'p_TypeSaisieAutre = "NUMFCT"
                Else
                    MsgBox "Cas à voir"
                    Me.TxtValeur.Text = s
                End If
            End If
            If Me.TxtOperateur.Text = "" Then
                Me.TxtOperateur.Text = ChercheOperateur(STR_GetChamp(g_straction, "¤", 1))
            End If
            Me.TxtOperateur.tag = STR_GetChamp(g_straction, "¤", 1)
        End If
        
        'txt(TXT_CONDF).Text = STR_GetChamp(g_straction, "µ", 1)
        'txt(TXT_CONDF).Text = g_straction
        
        If g_numChpCnd <> "" Then
            ' chercher si c'est une liste de valeurs
            sql = "select * from formetapechp where forec_num = " & g_numChpCnd
            Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
            sNumLst = ""
            numlst = -1
            g_numlst = 0
            If Not rs.EOF Then
                sNumLst = rs("forec_valeurs_possibles").Value
                If sNumLst <> "" Then
                    g_numlst = rs("forec_valeurs_possibles").Value
                End If
                p_BoolSaisieDate = False
                p_boolSaisieListe = False
                p_boolSaisieListeHierar = False
        
                If InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
                    p_BoolSaisieDate = True
                    Me.cmd(CMD_SAISIR_DATE).Visible = True
                    Me.cmd(CMD_CHOIX_VAL).Visible = False
                    Me.cmd(CMD_OP_EGAL).Visible = True
                    Me.cmd(CMD_OP_DIFF).Visible = False
                    Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                    Me.cmd(CMD_OP_INFERIEUR).Visible = True
                    Me.cmd(CMD_DATE_COMPRIS).Visible = True
                Else
                    Me.cmd(CMD_SAISIR_DATE).Visible = False
                    If rs("forec_type").Value = "CHECK" Or rs("forec_type").Value = "RADIO" Or rs("forec_type").Value = "SELECT" Then
                        p_boolSaisieListe = True
                        Me.cmd(CMD_CHOIX_VAL).Visible = True
                        Me.cmd(CMD_OP_EGAL).Visible = True
                        Me.cmd(CMD_OP_DIFF).Visible = True
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                        Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    ElseIf rs("forec_type").Value = "HIERARCHIE" Then
                        p_boolSaisieListeHierar = True
                        Me.cmd(CMD_CHOIX_VAL).Visible = True
                        Me.cmd(CMD_OP_EGAL).Visible = True
                        Me.cmd(CMD_OP_DIFF).Visible = True
                        Me.cmd(CMD_OP_SUPERIEUR).Visible = False
                        Me.cmd(CMD_OP_INFERIEUR).Visible = False
                    Else
                        If fctvalid <> "%NUMSERVICE" And fctvalid <> "%NUMFCT" Then
                            Me.cmd(CMD_CHOIX_VAL).Visible = False
                            Me.cmd(CMD_OP_EGAL).Visible = True
                            Me.cmd(CMD_OP_DIFF).Visible = True
                            Me.cmd(CMD_OP_SUPERIEUR).Visible = True
                            Me.cmd(CMD_OP_INFERIEUR).Visible = True
                        End If
                    End If
                End If
            End If
            '
            cmd(CMD_CHOIX_CHP).tag = g_numChpCnd
            If g_numlst > 0 Then
                cmd(CMD_CHOIX_VAL).Visible = True
                cmd(CMD_CHOIX_VAL).tag = g_numlst
            Else
                'cmd(CMD_CHOIX_VAL).visible = False
                cmd(CMD_CHOIX_VAL).tag = 0
            End If
        End If
    End If
    '
    'txt(TXT_TITRE).SetFocus
    g_mode_saisie = True

End Sub

Private Sub ajouter_hierar_fils(ByVal v_numval As Long, _
                                ByVal v_LstVal As String, _
                                ByVal v_niveau As Integer)

    Dim sql As String, sdecal As String, s As String
    Dim trouve As Boolean
    Dim rs As rdoResultset
    
    If v_numval < 0 Then
        sql = "select * from HierarValChp" _
            & " where HVC_LHCNum=" & -v_numval _
            & " and HVC_Numpere=0" _
            & " and HVC_Actif=true" _
            & " order by HVC_Ordre, HVC_Nom"
    Else
        sql = "select * from HierarValChp" _
            & " where HVC_Numpere=" & v_numval _
            & " and HVC_Actif=true" _
            & " order by HVC_Ordre, HVC_Nom"
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    sdecal = String$(v_niveau * 3, " ")
    While Not rs.EOF
        trouve = False
        If v_LstVal <> "" Then
            's = v_shierar & "M" & rs("HVC_Num").Value & ";"
            s = "M" & rs("HVC_Num").Value & ";"
            If InStr(v_LstVal, s) > 0 Then
                trouve = True
            End If
        End If
        Call CL_AddLigne(sdecal & rs("HVC_nom").Value, rs("HVC_Num").Value, s, trouve)
        Call ajouter_hierar_fils(rs("HVC_Num").Value, v_LstVal, v_niveau + 1)
        rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Function ChercheOperateur(ByVal v_operateur As String)

    'MsgBox "ici " & v_operateur
    If v_operateur = "OP:EG" Then
        ChercheOperateur = "Egal"
    ElseIf v_operateur = "OP:COMPRIS" Then
        ChercheOperateur = "Compris entre "
    ElseIf v_operateur = "<>" Or v_operateur = "OP:NE" Then
        ChercheOperateur = "Différent de"
    ElseIf v_operateur = ">" Then
        ChercheOperateur = "Supérieur"
    ElseIf v_operateur = ">=" Then
        ChercheOperateur = "Supérieur ou Egal"
    ElseIf v_operateur = "OP:SU" Then
        ChercheOperateur = "Supérieur ou Egal"
    ElseIf v_operateur = "<" Then
        ChercheOperateur = "Inférieur"
    ElseIf v_operateur = "<=" Then
        ChercheOperateur = "Inférieur ou Egal"
    ElseIf v_operateur = "OP:IN" Then
        ChercheOperateur = "Inférieur ou Egal"
    ElseIf v_operateur = "SRV" Then
        ChercheOperateur = "Inférieur ou Egal"
    ElseIf v_operateur = "FCT" Then
        ChercheOperateur = "Inférieur ou Egal"
    Else
        MsgBox "Cas à traiter : " & v_operateur
    End If
End Function

Private Function quitter(ByVal v_bforce As Boolean) As Boolean

    Dim reponse As Integer
    
    If Not v_bforce Then
        If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
            If g_ConfirmerSortie Then
                reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
                If reponse = vbNo Then
                    quitter = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' retourner à l'appelant
    g_retour_PrmFctJS = ""
    
    Unload Me
    
    quitter = True
    
End Function

Private Sub Reconstituer()
    
    Dim zout As String, car As String, car2 As String
    Dim zout2 As String
    Dim s As String
    Dim DebutCar As Boolean
    Dim pos3 As Integer
    Dim pos1 As Integer, pos2 As Integer
    Dim pipeOP As String, i As Integer
    Dim j As Integer
    Dim bChamp As Integer, bOperateur As Boolean, bValeur As Boolean
    Dim sChamp As String, sOperateur As String, sValeur As String
    
    Exit Sub
    bChamp = True
    
    DebutCar = True
    zout = ""
    pipeOP = "|"
    
    
    'MsgBox "a virer"
    'Debug.Print txtchamp
    'Debug.Print txt(TXT_CONDPF)
    'MsgBox TxtChamp.Text
    For i = 1 To Len(TxtChamp.Text)
        car = Mid$(TxtChamp.Text, i, 1)
        'Debug.Print i & " " & car
        If car = "(" Then
            zout = zout & pipeOP & "OP:("
            DebutCar = True
        ElseIf car = ")" Then
            zout = zout & pipeOP & "OP:)"
            DebutCar = True
        ElseIf car = "=" Then
            zout = zout & pipeOP & "OP:EG"
            DebutCar = True
        ElseIf car = "<" And Mid$(TxtChamp.Text, i + 1, 1) = ">" Then
            zout = zout & pipeOP & "OP:NE"
            DebutCar = True
            i = i + 2
        ElseIf car = ">" Then
            zout = zout & pipeOP & "OP:SU"
            DebutCar = True
            i = i + 1
        ElseIf car = "<" Then
            zout = zout & pipeOP & "OP:IN"
            DebutCar = True
            i = i + 1
        ElseIf car = "¤" Then
            zout = zout & pipeOP & "OP:IN"
            DebutCar = True
            i = i + 1
        ElseIf car = "[" And Mid$(TxtChamp.Text, i + 1, 3) = "OU]" Then
            zout = zout & pipeOP & "OP:OU"
            DebutCar = True
            i = i + 3
        ElseIf car = "[" And Mid$(TxtChamp.Text, i + 1, 3) = "ET]" Then
            zout = zout & pipeOP & "OP:ET"
            DebutCar = True
            i = i + 3
        ElseIf Mid$(TxtChamp.Text, i, 4) = "VAL:" Then
            pos1 = InStr(i, TxtChamp.Text, ":")
            'pos2 = InStr(I, txtchamp.Text, "[")
            'pos3 = InStr(I, txtchamp.Text, "]")
            'I = pos3 + 1
            's = Mid$(txtchamp.Text, pos1 + 1, pos2 - pos1 - 1)
            s = Mid$(TxtChamp.Text, pos1 + 1)
            zout = zout & pipeOP
            For j = 0 To STR_GetNbchamp(s, ";")
                If STR_GetChamp(s, ";", j) <> "" Then
                    zout = zout & "VAL:" & STR_GetChamp(s, ";", j) & ";"
                End If
            Next j
            Exit For
'            zout = zout & pipeOP & "VAL:" & s
'            DebutCar = True
            'i = pos2 + 1
        ElseIf Mid$(TxtChamp.Text, i, 5) = "DATE:" Then
            pos1 = InStr(i, TxtChamp.Text, ":")
            pos2 = InStr(i, TxtChamp.Text, "[")
            pos3 = InStr(i, TxtChamp.Text, "]")
            i = pos3 + 1
            s = Mid$(TxtChamp.Text, pos2 + 1, pos3 - pos2 - 1)
            ' s = Mid$(txtchamp.Text, 5, pos1 - 5)
            zout = zout & pipeOP & "DATE:" & s
            DebutCar = True
            'i = pos2 + 1
        Else
            If DebutCar Then
                zout = zout & pipeOP & car
                DebutCar = False
            Else
                zout = zout & car
            End If
        End If
        'Debug.Print zout
    Next i
    
    'Debug.Print "zout=" & zout
    zout = zout & "|"
    zout = Right(zout, Len(zout) - 1)
    'Debug.Print "zout=" & zout
    zout2 = ""
    For i = 1 To Len(zout)
        car = Mid$(zout, i, 1)
        'Debug.Print i & " " & car
        If car <> " " Then
            zout2 = zout2 & car
        End If
    Next i
    'Debug.Print "zout2=" & zout2
    zout = ""
    For i = 1 To Len(zout2)
        car = Mid$(zout2, i, 1)
        'Debug.Print i & " " & car
        If car = "|" Then
            car2 = Mid$(zout2, i + 1, 1)
            If car2 <> "|" Then
                zout = zout & car
            End If
        Else
            zout = zout & car
        End If
    Next i
    'Debug.Print "zout=" & zout
    TxtChamp.Text = zout
    
    'Debug.Print txtchamp
    'Debug.Print txt(TXT_CONDPF)
End Sub

Private Sub valider()
    Dim sep As String
    Dim s As String
    Dim sql As String
    Dim sCond As String, stype As String, fctvalid As String, sNumLst As String
    Dim cr As Integer
    Dim Fjava_Fornum As Integer, Fjava_Chpnum As Integer, Fjava_Etpnum As Integer
    Dim Fjava_Trait As String, Fjava_Chp As String, Fjava_Declencheur As String
    Dim Fjava_Condition As String, Fjava_Condition_Fr As String
    Dim Fjava_Action As String, Fjava_Action_Fr As String, Fjava_Titre As String
    Dim op As String
    Dim i As Integer
    
    cr = verifier_tous_chp()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_NON Then
        Exit Sub
    End If
    
    Fjava_Fornum = g_numfor
    Fjava_Trait = g_Trait
    Fjava_Chp = g_strchp
    Fjava_Condition = "CHP:" & Me.TxtChamp.tag & ":" & FctNomChp(Me.TxtChamp.tag) & "¤"
    Fjava_Condition = Fjava_Condition & Me.TxtOperateur.tag & "¤"
    
    sql = "select FOREC_Label, FOREC_Type, FOREC_FctValid, FOREC_Valeurs_Possibles from FormEtapeChp" _
        & " where FOREC_Num=" & Me.TxtChamp.tag
    If Odbc_RecupVal(sql, sCond, stype, fctvalid, sNumLst) = P_ERREUR Then
        MsgBox "Erreur SQL " & sql
    End If
    If fctvalid = "%NUMSERVICE" Or stype = "HIERARCHIE" Then
        Fjava_Condition = Fjava_Condition & "VAL:" & Me.TxtValeur.tag & "_" & IIf(Me.ChkSrvDet.Value = 1, "DET", "") & "¤"
    Else
        If InStr(fctvalid, "%DATE") > 0 Then
            Me.TxtValeur.tag = Replace(Me.TxtValeur.tag, "DATE:", "")
            Fjava_Condition = Fjava_Condition & "DATE:" & Me.TxtValeur.tag & "¤"
        Else
            Fjava_Condition = Fjava_Condition & "VAL:" & Me.TxtValeur.tag & "¤"
        End If
    End If
    Fjava_Condition_Fr = "en français" ' txt(TXT_CONDF).Text
    op = ""
    Fjava_Chpnum = Me.TxtChamp.tag
    
    If g_boolListeFen Then
        Fjava_Condition = Fjava_Condition & cmd(CMD_CHOIX_FENETRE).tag & "¤"
    Else
        Fjava_Condition = Fjava_Condition & "¤"
    End If
    g_retour_PrmFctJS = Fjava_Condition     ' & "µ" & Fjava_Condition_Fr
    
    Unload Me
    Exit Sub
    
err_enreg:
    Unload Me
    
End Sub

Private Function verifier_tous_chp() As Integer
    
    Dim ret As Integer
        
    verifier_tous_chp = P_OUI
    
    If val(Me.TxtChamp.tag) > 0 Then
    Else
        verifier_tous_chp = P_NON
        MsgBox "Vous devez choisir un Champ"
    End If
    If Me.TxtOperateur.tag <> "" Then
    Else
        verifier_tous_chp = P_NON
        MsgBox "Vous devez choisir un Opérateur"
    End If
    If Me.TxtValeur.tag <> "" Then
    Else
        verifier_tous_chp = P_NON
        MsgBox "Vous devez choisir une Valeur"
    End If
    
End Function


Private Sub ChkSrvDet_Click()

    If Me.ChkSrvDet.Value = 1 Then
        Me.TxtOperateur.Text = "Fait partie de"
    Else
        Me.TxtOperateur.Text = "Strictement égal à"
    End If

End Sub

Private Sub cmd_Click(Index As Integer)

    Dim ret As Long
    Dim sValeur As String
    Dim sCond As String, stype As String, sNumLst As String
    Dim numlig As Integer
    Dim s As String
    Dim fctvalid As String
    Dim position As Integer
    Dim frm As Form
    Dim chpnum As Long
    Dim sret As String
    Dim numaction As Integer
    Dim strD As String, strG As String
    Dim strGPF As String, strDPF As String
    Dim strSel As String
    Dim str_action As String
    Dim iRound As Integer
    Dim i As Integer
    Dim newlen As Integer
    Dim TxtSaisieDate As String
    Dim Anc_Selstart As Integer
    Dim bok As Boolean
    Dim sql As String, rs As rdoResultset
    Dim iret As Long
    Dim d1 As String, d2 As String
    
    newlen = 0
    g_boolOper = False
    
    Select Case Index
    
    Case CMD_DATE_COMPRIS
        Me.TxtOperateur.tag = "OP:COMPRIS"
        Me.TxtOperateur.Text = "Compris"
        d1 = STR_GetChamp(TxtValeur.Text, " ", 0)
        d1 = SaisirDate(d1, g_Trait, "OP:SU", "Date de Début")
        d2 = STR_GetChamp(TxtValeur.Text, " ", 1)
        d2 = SaisirDate(d2, g_Trait, "OP:IN", "Date de Fin")
        If g_Trait = "Ajout" Or g_Trait = "Modif" Then
            Me.TxtValeur.Text = d1 & " " & d2
            Me.TxtValeur.tag = d1 & " " & d2
        End If
    Case CMD_CHOIX_FENETRE
        sret = ChoisirFenetres(cmd(CMD_CHOIX_FENETRE).tag)
        If sret <> "" Then
            cmd(CMD_CHOIX_FENETRE).tag = sret
            cmd(CMD_CHOIX_FENETRE).Caption = FaitListeFenetres(cmd(CMD_CHOIX_FENETRE).tag)
            If sret = "*" Then
                cmd(CMD_CHOIX_FENETRE).BackColor = &HE0E0E0
            Else
                cmd(CMD_CHOIX_FENETRE).BackColor = &H8080FF
            End If
        End If
    Case CMD_VIRGULE
        Me.TxtValeur.Text = Me.TxtValeur.Text & ","
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_0
        Me.TxtValeur.Text = Me.TxtValeur.Text & "0"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_1
        Me.TxtValeur.Text = Me.TxtValeur.Text & "1"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_2
        Me.TxtValeur.Text = Me.TxtValeur.Text & "2"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_3
        Me.TxtValeur.Text = Me.TxtValeur.Text & "3"
        Me.TxtValeur.tag = Me.TxtValeur.Text
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_4
        Me.TxtValeur.Text = Me.TxtValeur.Text & "4"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_5
        Me.TxtValeur.Text = Me.TxtValeur.Text & "5"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_6
        Me.TxtValeur.Text = Me.TxtValeur.Text & "6"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_7
        Me.TxtValeur.Text = Me.TxtValeur.Text & "7"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_8
        Me.TxtValeur.Text = Me.TxtValeur.Text & "8"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    Case CMD_BOUT_9
        Me.TxtValeur.Text = Me.TxtValeur.Text & "9"
        Me.TxtValeur.tag = Me.TxtValeur.Text
    
    Case CMD_SLASH
    Case CMD_SAISIR_DATE
        TxtSaisieDate = SaisirDate(TxtValeur.Text, g_Trait, Me.TxtOperateur.tag, "")
        Me.TxtValeur.Text = TxtSaisieDate
        Me.TxtValeur.tag = TxtSaisieDate
    Case CMD_VIDE
        Me.TxtValeur.Text = "[VIDE]"
        Me.TxtValeur.tag = "<NR>"
    Case CMD_OP_ET
    Case CMD_OP_OU
    Case CMD_OP_EGAL
        Me.TxtOperateur.tag = "OP:EG"
        Me.TxtOperateur.Text = "Egal"
        g_boolOper = True
    Case CMD_OP_SUPERIEUR
        Me.TxtOperateur.tag = "OP:SU"
        Me.TxtOperateur.Text = "Supérieur"
        g_boolOper = True
    Case CMD_OP_INFERIEUR
        Me.TxtOperateur.tag = "OP:IN"
        Me.TxtOperateur.Text = "Inférieur"
        g_boolOper = True
    Case CMD_OP_DIFF
        Me.TxtOperateur.tag = "OP:NE"
        FrmSrvDet.Visible = False
        ChkSrvDet.Value = 1
        Me.TxtOperateur.Text = "Différent"
        g_boolOper = True
    Case CMD_PAR_OUV
    Case CMD_PAR_FER
    Case CMD_CHOIX_CHP
        sret = ChoisirChamp(g_numfor)
        If IsNumeric(sret) And sret >= 0 Then
            sql = "select FOREC_Label, FOREC_Type, FOREC_FctValid, FOREC_Valeurs_Possibles from FormEtapeChp" _
                & " where FOREC_Num=" & sret
            If Odbc_RecupVal(sql, sCond, stype, fctvalid, sNumLst) = P_ERREUR Then
                MsgBox "Erreur SQL " & sql
            End If
            Me.TxtChamp.Text = sCond & " (" & STR_GetChamp(STR_GetChamp(g_straction, "¤", 0), ":", 2) & ")"
            If stype = "CHECK" Or stype = "RADIO" Or stype = "SELECT" Or stype = "HIERARCHIE" Then
                cmd_Click (CMD_OP_EGAL)
                Exit Sub
            ElseIf fctvalid = "%NUMSERVICE" Then
                p_BoolSaisieAutre = True
                cmd_Click (CMD_OP_EGAL)
                Exit Sub
            ElseIf fctvalid = "%NUMFCT" Then
                p_BoolSaisieAutre = True
                cmd_Click (CMD_OP_EGAL)
                Exit Sub
            End If
            p_derchamp = sret
        End If
    Case CMD_CHOIX_VAL
        If p_boolSaisieListe Then
            sret = choisir_valeur()
        ElseIf p_boolSaisieListeHierar Then
            sret = choisir_valeur_hierar()
        End If
    Case CMD_CHOIX_AUTRE
        sret = choisir_valeur_autre()
        If sret = "0" Then
            Exit Sub
        End If
        If p_TypeChamp = "SRV" Then
            Me.TxtValeur.Text = ""
            Me.TxtValeur.tag = ""
            For i = 0 To STR_GetNbchamp(sret, ";")
                s = STR_GetChamp(sret, ";", i)
                If s <> "" Then
                    iret = Replace(s, "S", "")
                    Call P_RecupSrvNom(iret, sValeur)
                    Me.TxtValeur.Text = IIf(Me.TxtValeur.Text = "", sValeur, Me.TxtValeur.Text & " OU " & sValeur)
                    Me.TxtValeur.tag = IIf(Me.TxtValeur.tag = "", s, Me.TxtValeur.tag & ";" & s)
                End If
            Next i
        ElseIf p_TypeChamp = "FCT" Then
            Me.Visible = True
        End If
    Case CMD_OK
        Call valider
        Exit Sub
    Case CMD_QUITTER
        Call quitter(False)
        Exit Sub
    Case CMD_SUPPRIMER
    End Select
    
    If g_boolOper And p_BoolSaisieDate Then
        cmd_Click (CMD_SAISIR_DATE)
        p_BoolSaisieDate = False
    End If
    
    If g_boolOper And p_boolSaisieListe Then
        cmd_Click (CMD_CHOIX_VAL)
        p_boolSaisieListe = False
    End If
    
    If g_boolOper And p_boolSaisieListeHierar Then
        cmd_Click (CMD_CHOIX_VAL)
        p_boolSaisieListeHierar = False
    End If
    
    If g_boolOper And p_BoolSaisieAutre Then
        cmd_Click (CMD_CHOIX_AUTRE)
        p_BoolSaisieAutre = False
    End If
    
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_QUITTER Then
        g_mode_saisie = False
    End If

End Sub

Private Function SaisirDate(v_strdate As String, v_Trait As String, v_operateur As String, v_titre As String)
    Dim TxtSaisieDate As String
    
SaisieDate:
    If v_titre = "" Then v_titre = "Saisir une date"
    If g_Trait = "Ajout" Then
        TxtSaisieDate = "01/01/" & Format(Date, "yyyy")
        If v_operateur = "OP:IN" Then
            TxtSaisieDate = "31/12/" & IIf(p_derannée = "", Format(Date, "yyyy"), p_derannée)
        ElseIf v_operateur = "OP:SU" Then
            TxtSaisieDate = "01/01/" & Format(Date, "yyyy")
        End If
        TxtSaisieDate = InputBox(v_titre, "Saisir une date", TxtSaisieDate)
        p_derannée = Mid(TxtSaisieDate, 7, 4)
    Else
        TxtSaisieDate = v_strdate
        p_derannée = Mid(v_strdate, 7, 4)
        If v_operateur = "OP:SU" Then
            TxtSaisieDate = Mid(TxtSaisieDate, 1, 6) & IIf(p_derannée = "", Format(Date, "yyyy"), p_derannée)
        ElseIf v_operateur = "OP:IN" Then
            TxtSaisieDate = Mid(TxtSaisieDate, 1, 6) & IIf(p_derannée = "", Format(Date, "yyyy"), p_derannée)
        End If
        TxtSaisieDate = InputBox(v_titre, "Saisir une date", TxtSaisieDate)
    End If
    If TxtSaisieDate = "" Then
        Exit Function
    End If
    If Not SAIS_CtrlChamp(TxtSaisieDate, SAIS_TYP_DATE) Then
        MsgBox "Le format d'une date est JJ/MM/AAAA"
        GoTo SaisieDate
    End If
    'Me.TxtValeur.Text = TxtSaisieDate
    'Me.TxtValeur.tag = TxtSaisieDate
    SaisirDate = TxtSaisieDate
End Function

Private Sub Form_Activate()

    If g_form_active Then
        Exit Sub
    End If
    
    g_form_active = True
    g_modif_val_directe = False
    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then
            Call valider
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "doc_a1_5_1_prmnature.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If TypeOf Me.ActiveControl Is TextBox Then
            Exit Sub
        End If
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If Not quitter(False) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub TxtValeur_Change()
    If g_modif_val_directe Then
        'Me.TxtValeur.tag = Me.TxtValeur.Text
    End If
End Sub
