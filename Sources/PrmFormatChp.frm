VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form PrmFormatChp 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix d'un Champ :"
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
      Height          =   9135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11595
      Begin VB.Frame FrmNivHier 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   7200
         Visible         =   0   'False
         Width           =   11235
         Begin VB.CheckBox chk_niveau_exact_H 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ce niveau seulement"
            Height          =   255
            Left            =   7200
            TabIndex        =   49
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox CmbNivHier 
            Height          =   315
            Left            =   3120
            TabIndex        =   47
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Niveau dans la Liste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame FrmNiveauStru 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         TabIndex        =   43
         Top             =   7080
         Visible         =   0   'False
         Width           =   11235
         Begin VB.ComboBox CmbNivStru 
            Height          =   315
            Left            =   6480
            TabIndex        =   52
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chk_niveau_exact_S 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ce niveau seulement"
            Height          =   255
            Left            =   9120
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox CmbTypeStru 
            Height          =   315
            Left            =   1920
            TabIndex        =   44
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblStruNiv 
            BackColor       =   &H00C0C0C0&
            Caption         =   " Niveau"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5520
            TabIndex        =   51
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblStruType 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Type de Structure"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame FrmRepartDates 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   7320
         Visible         =   0   'False
         Width           =   11235
         Begin VB.ComboBox CmbRepartDates 
            Height          =   315
            Left            =   2640
            TabIndex        =   42
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Répartition de date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Délier"
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
         Index           =   4
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Champ relié à : "
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
         Index           =   3
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Frame FrmSQL 
         BackColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   120
         TabIndex        =   35
         Top             =   8040
         Visible         =   0   'False
         Width           =   11235
         Begin VB.TextBox TxtSQL 
            Height          =   735
            Left            =   1560
            TabIndex        =   37
            Top             =   240
            Width           =   9495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Requête SQL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame FrmLibVal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Présentation"
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
         Height          =   765
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   6795
         Begin VB.OptionButton OptLibVal 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé + Valeur"
            Height          =   375
            Index           =   2
            Left            =   4440
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptLibVal 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton OptLibVal 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valeur"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame FrmGen 
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
         ForeColor       =   &H00800080&
         Height          =   7575
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   11445
         Begin VB.Frame FrmParRapport 
            BackColor       =   &H00C0C0C0&
            Caption         =   "par rapport à"
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
            Height          =   1125
            Left            =   120
            TabIndex        =   32
            Top             =   6240
            Visible         =   0   'False
            Width           =   11085
            Begin VB.OptionButton OptRapportVal 
               BackColor       =   &H00C0C0C0&
               Caption         =   "au nombre de FEI"
               Height          =   375
               Index           =   1
               Left            =   720
               TabIndex        =   34
               Top             =   720
               Width           =   10215
            End
            Begin VB.OptionButton OptRapportVal 
               BackColor       =   &H00C0C0C0&
               Caption         =   "FEI dont champ OUI NON renseigné "
               Height          =   375
               Index           =   0
               Left            =   720
               TabIndex        =   33
               Top             =   360
               Width           =   10215
            End
         End
         Begin VB.Frame FrmNbOccur 
            BackColor       =   &H00C0C0C0&
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Visible         =   0   'False
            Width           =   6795
            Begin VB.OptionButton OptNbOccur 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Toutes"
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
               Index           =   0
               Left            =   360
               TabIndex        =   22
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton OptNbOccur 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Choisir"
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
               Index           =   1
               Left            =   1560
               TabIndex        =   21
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtNbOccur 
               Height          =   285
               Left            =   4080
               TabIndex        =   20
               Text            =   "5"
               Top             =   690
               Width           =   855
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nombre d'occurences à afficher"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   240
               Width           =   3015
            End
            Begin VB.Label LblNbreOccur 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3240
               TabIndex        =   23
               Top             =   720
               Width           =   1815
            End
         End
         Begin VB.Frame FrmFormat 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Format des Valeurs"
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
            Height          =   2325
            Left            =   120
            TabIndex        =   13
            Top             =   3840
            Width           =   6795
            Begin VB.Frame FrmEtendu 
               BackColor       =   &H00C0C0C0&
               Caption         =   "étendue des Valeurs du champ OUI NON"
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
               Height          =   1125
               Left            =   120
               TabIndex        =   26
               Top             =   1080
               Width           =   6525
               Begin VB.OptionButton OptTypVal 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Toutes les valeurs significatives"
                  Height          =   375
                  Index           =   1
                  Left            =   360
                  TabIndex        =   31
                  Top             =   720
                  Width           =   3255
               End
               Begin VB.OptionButton OptTypVal 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Toutes les valeurs"
                  Height          =   375
                  Index           =   0
                  Left            =   360
                  TabIndex        =   29
                  Top             =   360
                  Width           =   3015
               End
               Begin VB.OptionButton OptTypVal 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Renseignées seulement"
                  Height          =   375
                  Index           =   2
                  Left            =   3720
                  TabIndex        =   28
                  Top             =   360
                  Width           =   2655
               End
               Begin VB.OptionButton OptTypVal 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Non renseignées seulement"
                  Height          =   375
                  Index           =   3
                  Left            =   3720
                  TabIndex        =   27
                  Top             =   720
                  Width           =   2655
               End
            End
            Begin VB.OptionButton OptValForme 
               BackColor       =   &H00C0C0C0&
               Caption         =   "écart type"
               Height          =   375
               Index           =   3
               Left            =   4560
               TabIndex        =   25
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton OptValForme 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Moyenne"
               Height          =   375
               Index           =   2
               Left            =   3360
               TabIndex        =   18
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton OptValForme 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Somme"
               Height          =   375
               Index           =   1
               Left            =   2280
               TabIndex        =   17
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton OptValForme 
               BackColor       =   &H00C0C0C0&
               Caption         =   "nombre"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   4
               Left            =   5760
               TabIndex        =   16
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton OptValForme 
               BackColor       =   &H00C0C0C0&
               Caption         =   "pourcentage"
               Height          =   375
               Index           =   0
               Left            =   840
               TabIndex        =   14
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Je veux les valeurs sous la forme de ..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   480
               TabIndex        =   30
               Top             =   360
               Width           =   5535
            End
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "choix d'une Valeur"
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
            Index           =   2
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptLigCol 
            BackColor       =   &H00C0C0C0&
            Caption         =   "en colonne"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
         Begin VB.ListBox ListVal 
            Columns         =   2
            Height          =   5685
            ItemData        =   "PrmFormatChp.frx":0000
            Left            =   7080
            List            =   "PrmFormatChp.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   360
            Width           =   4095
         End
         Begin VB.OptionButton OptLigCol 
            BackColor       =   &H00C0C0C0&
            Caption         =   "en ligne"
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
            Index           =   0
            Left            =   2400
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   975
            Left            =   4320
            Top             =   240
            Width           =   2535
         End
         Begin ComctlLib.ImageList imglst 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   62
            ImageHeight     =   31
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   2
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmFormatChp.frx":0004
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmFormatChp.frx":0C16
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label LabValChp 
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
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   1200
            Width           =   6615
         End
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   -15
      TabIndex        =   0
      Top             =   9135
      Width           =   11595
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmFormatChp.frx":18A4
         Enabled         =   0   'False
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
         Left            =   120
         Picture         =   "PrmFormatChp.frx":1E00
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   240
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
         Left            =   10920
         Picture         =   "PrmFormatChp.frx":2369
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmFormatChp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_SEL = &HC00000

' Index des objets cmd
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_CHOIX_VAL = 2
Private Const CMD_CHP_RELIER = 3
Private Const CMD_DELIER = 4

' nbre occurences
Private Const Opt_nboccur_Toutes = 0
Private Const Opt_nboccur_Choisir = 1

Private s_chp_nom As String
Private s_form_nom As String

' libelle - valeur ou les 2
Private Const Opt_LibVal_L = 0
Private Const Opt_LibVal_V = 1
Private Const Opt_LibVal_LV = 2

' forme du retour
Private Const Opt_ValForme_Pourcent = 0
Private Const Opt_ValForme_Somme = 1
Private Const Opt_ValForme_Moyenne = 2
Private Const Opt_ValForme_EcarType = 3
Private Const Opt_ValForme_Nombre = 4

' forme de la présentation
Private Const OptTypVal_Valeur = 0
Private Const OptTypVal_ValeurSignif = 1
Private Const OptTypVal_Valeur_Rens = 2
Private Const OptTypVal_Valeur_NonRens = 3

'par rapport à
Private Const OptRapportVal_Rens = 0
Private Const OptRapportVal_Toutes = 1

' Index des objets frm
Private Const FRM_PROP = 0

Private g_Trait As String
Private g_chpnum As Integer
Private g_fornum As Integer
Private g_filtrenum As Integer
Private g_i_tbExcel As Integer

Private g_MenForme As String
Private g_MenFormeListe As String
Private g_MenFormeNonListe As String

Private g_ValeurListe As String
Private g_boolLstSpécial As Boolean
Private g_strLstSpécial As String
Private FaireListeClick As Boolean

Private g_boolTexte As Boolean, g_BoolEntier As Boolean, g_BoolMTT As Boolean, g_boolListe As Boolean, g_boolDate As Boolean
Private g_boolCalcul As Boolean
Private g_forec_valeurs_possibles As String

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String

Public g_retour_PrmFormatChp As String

Public Function AppelFrm(ByVal v_Trait As String, _
                         ByVal V_ChpNum As Integer, _
                         ByVal v_i_tbExcel As Integer, _
                         ByVal v_fornum As Integer, _
                         ByVal v_filtrenum As Integer, _
                         ByVal v_MenForme As String, _
                         ByVal v_boolLstSpécial As Boolean, _
                         ByVal v_strLstSpécial As String) As String
    
    g_i_tbExcel = v_i_tbExcel
    g_filtrenum = v_filtrenum
    g_MenForme = v_MenForme
    g_fornum = v_fornum
    g_chpnum = V_ChpNum
    g_Trait = v_Trait
    g_ValeurListe = ""
    g_boolLstSpécial = v_boolLstSpécial
    g_strLstSpécial = v_strLstSpécial
    If g_chpnum <= -10 Then
        If Odbc_RecupVal("select ff_fornum from filtreform where ff_num =" & v_filtrenum, g_fornum) = P_ERREUR Then
            Exit Function
        End If
    End If
    If p_ModeAssistant Then
        MsgBox "Indiquez la présentation que vous voulez appliquer sur les champs du rapport" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "(la même présentation pour tous, que vous pourrez modifier ensuite)"
    End If
    
    Me.Show 1
    
    If g_ValeurListe = "" Then
        g_ValeurListe = "TOUTES"
    End If
    If g_Trait = "Ajout" Or g_Trait = "AjoutPlusieurs" Or g_Trait = "Modif" Then
        'MsgBox g_retour_PrmFctJS
        AppelFrm = g_retour_PrmFormatChp
        If Mid(g_retour_PrmFormatChp, 1, 4) = "SQL=" Then
        ElseIf g_retour_PrmFormatChp <> "QUITTER" Then
            If p_ModeAssistant Then
                If Mid(g_retour_PrmFormatChp, 1, 5) = "Ligne" Then
                    g_retour_PrmFormatChp = "Ligne_Lib_Val#NOMBRE#VALEUR#8#TOUTES#"
                ElseIf Mid(g_retour_PrmFormatChp, 1, 7) = "Colonne" Then
                    g_retour_PrmFormatChp = "Colonne_Lib_Val#NOMBRE#VALEUR#8#TOUTES#"
                Else
                    g_retour_PrmFormatChp = "Ligne_Lib_Val#NOMBRE#VALEUR#8#TOUTES#"
                End If
                AppelFrm = g_retour_PrmFormatChp
            End If
            If g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_boolListe Then
                p_Derniere_MenFormeListe = STR_GetChamp(g_retour_PrmFormatChp, "#", 0)
                p_Derniere_MenFormeListe = p_Derniere_MenFormeListe & "#" & STR_GetChamp(g_retour_PrmFormatChp, "#", 1)
                p_Derniere_MenFormeListe = p_Derniere_MenFormeListe & "#" & STR_GetChamp(g_retour_PrmFormatChp, "#", 2)
                p_Derniere_MenFormeListe = p_Derniere_MenFormeListe & "#*#TOUTES"
                p_Derniere_MenFormeListe = p_Derniere_MenFormeListe & "#" & STR_GetChamp(g_retour_PrmFormatChp, "#", 5)
            Else
                p_Derniere_MenFormeNonListe = g_retour_PrmFormatChp
            End If
        End If
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


Private Sub initialiser()

    Dim i As Integer
    Dim ret As String
    Dim sql As String
    Dim rs As rdoResultset
    Dim rs2 As rdoResultset
    Dim n As Integer
    Dim laS As String
    Dim sG As String, sD As String
    Dim strG As String, strD As String
    Dim LaStr As String
    Dim niveauRelier As String
    ' Dim boolTexte As Boolean, BoolEntier As Boolean, boolListe As Boolean, boolDate As Boolean
    Dim Anc_Forme As String, Forme As String
    Dim s As String
    Dim rstmp As rdoResultset
    Dim sFormule As String
    Dim sforlib As String
    Dim sF As String, sX As String, sY As String
    Dim strRelié_à As String
    
    g_boolTexte = False
    g_BoolEntier = False
    g_BoolMTT = False
    g_boolListe = False
    g_boolDate = False
    g_boolCalcul = False
        
    ' lire l'action
    g_mode_saisie = False
      
    If g_chpnum <= -10 Then
        ' C'est un nombre de fiches
        If Odbc_RecupVal("select for_lib from formulaire where for_num =" & g_fornum, sforlib) = P_ERREUR Then
            Exit Sub
        End If
        Me.Frame.Caption = "Champ : Nombre Total de fiches : " & sforlib
        If True Then   ' Or g_Trait = "Modif" And g_strLstSpécial <> "") Then ' requete SQL
            Me.FrmSQL.Visible = True
            If g_Trait = "Ajout" Then
                g_strLstSpécial = ""
            End If
            Me.TxtSQL.Text = g_strLstSpécial
        End If
        GoTo LabCasGeneral
    End If
    
    sql = "select * from formetapechp" _
        & " where forec_num=" & g_chpnum
    If Odbc_Select(sql, rs) = P_ERREUR Then
        MsgBox "PrmFormatChp : Champ " & g_chpnum & " introuvable"
        Exit Sub
    Else
        ' ce champ est il relié ?
        If tbl_fichExcel(g_i_tbExcel).CmdChpRelierà <> "" Then
            cmd(CMD_CHP_RELIER).Visible = True
            cmd(CMD_DELIER).Visible = True
            sF = STR_GetChamp(tbl_fichExcel(g_i_tbExcel).CmdChpRelierà, ";", 0)
            sX = STR_GetChamp(tbl_fichExcel(g_i_tbExcel).CmdChpRelierà, ";", 1)
            sY = STR_GetChamp(tbl_fichExcel(g_i_tbExcel).CmdChpRelierà, ";", 2)
            For i = 0 To UBound(tbl_fichExcel)
                If tbl_fichExcel(i).CmdFenNum = sF Then
                    If tbl_fichExcel(i).CmdX = sX Then
                        If tbl_fichExcel(i).CmdY = sY Then
                            Exit For
                        End If
                    End If
                End If
            Next i
            If i > UBound(tbl_fichExcel) Then
                MsgBox "Relié à " & tbl_fichExcel(g_i_tbExcel).CmdChpRelierà & " ???"
                tbl_fichExcel(g_i_tbExcel).CmdChpRelierà = ""
                cmd(CMD_CHP_RELIER).Visible = False
                Me.cmd(CMD_DELIER).Visible = False
            Else
                Call Odbc_RecupVal("select forec_label from formetapechp where forec_num=" & tbl_fichExcel(i).CmdChpNum, s)
                ' le combo si c'est un champ service ou hierar
                If rs("forec_type") = "HIERARCHIE" Or rs("forec_fctvalid") = "%NUMSERVICE" Then
                    'niveauRelier = tbl_fichExcel(g_i_tbExcel).CmdNiveauRelier
                    'If niveauRelier = "" Then niveauRelier = "Tous"
                    'If rs("forec_type") = "HIERARCHIE" Then
                    '    Me.CmbNiveau.AddItem "Tous"
                    '    Me.CmbNiveau.AddItem "1"
                    '    Me.CmbNiveau.AddItem "2"
                    '    Me.CmbNiveau.AddItem "3"
                    '    Me.CmbNiveau.AddItem "4"
                    'ElseIf rs("forec_fctvalid") = "%NUMSERVICE" Then
                    '    Me.CmbNiveau.AddItem "Tous"
                    '    Me.CmbNiveau.AddItem "1"
                    '    Me.CmbNiveau.AddItem "2"
                    '    Me.CmbNiveau.AddItem "3"
                    '    Me.CmbNiveau.AddItem "4"
                    'End If
                    'Me.CmbNiveau.Text = niveauRelier
                    'Me.CmbNiveau.tag = niveauRelier
                End If
                cmd(CMD_CHP_RELIER).Caption = "Champ relié à : " & s & " (F" & tbl_fichExcel(i).CmdFenNum & " - " & tbl_fichExcel(i).CmdX & tbl_fichExcel(i).CmdY & ")"
            End If
        Else
            ' Y a t - il des champs qui lui sont reliés ?
            strRelié_à = ""
            For i = 0 To UBound(tbl_fichExcel)
                If tbl_fichExcel(i).CmdChpRelierà <> "" Then
                    sF = STR_GetChamp(tbl_fichExcel(i).CmdChpRelierà, ";", 0)
                    sX = STR_GetChamp(tbl_fichExcel(i).CmdChpRelierà, ";", 1)
                    sY = STR_GetChamp(tbl_fichExcel(i).CmdChpRelierà, ";", 2)
                    If tbl_fichExcel(g_i_tbExcel).CmdFenNum = sF And tbl_fichExcel(g_i_tbExcel).CmdX = sX And tbl_fichExcel(g_i_tbExcel).CmdY = sY Then
                        Call Odbc_RecupVal("select forec_label from formetapechp where forec_num=" & tbl_fichExcel(i).CmdChpNum, s)
                        s = s & " (F" & tbl_fichExcel(i).CmdFenNum & " - " & tbl_fichExcel(i).CmdX & tbl_fichExcel(i).CmdY & ")"
                        strRelié_à = strRelié_à & IIf(strRelié_à <> "", " ET ", "") & s
                    End If
                End If
            Next i
            If strRelié_à <> "" Then
                cmd(CMD_CHP_RELIER).Caption = "Les Champs reliés : " & strRelié_à
                cmd(CMD_CHP_RELIER).Visible = True
            Else
                cmd(CMD_CHP_RELIER).Visible = False
            End If
            cmd(CMD_DELIER).Visible = False
        End If
        
        ' mettre le label du champ
        If Not rs.EOF Then
            Me.Frame.Caption = "Champ : " & rs("forec_label") & "  (" & rs("forec_nom") & ")"
            cmd(CMD_CHOIX_VAL).Visible = False
            If g_boolLstSpécial Then
                cmd(CMD_CHOIX_VAL).tag = g_strLstSpécial
                If g_strLstSpécial = "%NUMSERVICE" Then
                    Me.FrmNiveauStru.Visible = True
                    cmd(CMD_CHOIX_VAL).Visible = True
                    cmd(CMD_CHOIX_VAL).Caption = "Services"
                    
                    s = STR_GetChamp(g_MenForme, "#", 7)
                    Dim TypeNiv As String
                    TypeNiv = Mid(s, 3, 1)
                    If TypeNiv = "" Then
                        TypeNiv = "N"   ' par niveau N / par Type T
                    End If
                    Call Odbc_SelectV("select * from Niveau_Structure Order By Nivs_NivPere", rs2)
                    If rs2.EOF Then
                        Me.lblStruType.Visible = False
                        Me.CmbTypeStru.Visible = False
                    Else
                        ' Liste par structure
                        Me.CmbTypeStru.AddItem " "
                        Me.CmbTypeStru.AddItem "Tous"
                        Me.chk_niveau_exact_S.Visible = True
                        s = STR_GetChamp(g_MenForme, "#", 7)
                        s = Mid(s, 1, 1)
                        If s = "" Or s = "0" Then
                            Me.CmbTypeStru.ListIndex = 0
                        End If
                        While Not rs2.EOF
                            Me.CmbTypeStru.AddItem rs2("Nivs_Nom")
                            Me.CmbTypeStru.ItemData(Me.CmbTypeStru.ListCount - 1) = rs2("Nivs_Num")
                            If s <> "" Then
                                If s = rs2("Nivs_Num") Then
                                    Me.CmbTypeStru.ListIndex = Me.CmbTypeStru.ListCount - 1
                                End If
                            End If
                            rs2.MoveNext
                        Wend
                    End If
                    rs2.Close
                    ' Liste par Niveau
                    Me.CmbNivStru.AddItem " "
                    Me.CmbNivStru.AddItem "Tous"
                    s = STR_GetChamp(g_MenForme, "#", 7)
                    If s = "" Then s = "0"
                    s = Mid(s, 1, 1)
                    For i = 1 To 5
                        Me.CmbNivStru.AddItem i
                        Me.CmbNivStru.ItemData(Me.CmbNivStru.ListCount - 1) = i
                        If s <> "" Then
                            If i = s Then
                                Me.CmbNivStru.ListIndex = Me.CmbNivStru.ListCount - 1
                            End If
                        End If
                    Next i
                    If s = "" Or s = "0" Then
                        Me.CmbNivStru.ListIndex = 0
                    End If
                        
                    ' lequel des 2 ?
                    If TypeNiv = "N" Then
                        Me.CmbTypeStru.ListIndex = 0
                    Else
                        Me.CmbNivStru.ListIndex = 0
                    End If
                    
                    ' exact ?
                    s = STR_GetChamp(g_MenForme, "#", 7)
                    s = Mid(s, 2, 1)
                    If s <> "" Then
                        If s = "O" Then
                            Me.chk_niveau_exact_S.Value = 1
                        Else
                            Me.chk_niveau_exact_S.Value = 0
                        End If
                    End If
                ElseIf g_strLstSpécial = "HIERARCHIE" Then
                    Me.FrmNivHier.Visible = True
                    Me.chk_niveau_exact_H.Visible = True
                    Me.CmbNivHier.AddItem "Tous"
                    s = STR_GetChamp(g_MenForme, "#", 7)
                    If s = "" Then s = "0"
                    s = Mid(s, 1, 1)
                    For i = 1 To 5
                        Me.CmbNivHier.AddItem i
                        Me.CmbNivHier.ItemData(Me.CmbNivHier.ListCount - 1) = i
                        If s <> "" Then
                            If i = s Then
                                Me.CmbNivHier.ListIndex = Me.CmbNivHier.ListCount - 1
                            End If
                        End If
                    Next i
                    If s = "" Or s = "0" Then
                        Me.CmbNivHier.ListIndex = 0
                    End If
                    ' exact ?
                    s = STR_GetChamp(g_MenForme, "#", 7)
                    s = Mid(s, 2, 1)
                    If s <> "" Then
                        If s = "O" Then
                            Me.chk_niveau_exact_H.Value = 1
                        Else
                            Me.chk_niveau_exact_H.Value = 0
                        End If
                    End If
                    
                    cmd(CMD_CHOIX_VAL).tag = g_strLstSpécial & "%" & rs("forec_valeurs_possibles").Value
                    cmd(CMD_CHOIX_VAL).Visible = True
                    cmd(CMD_CHOIX_VAL).Caption = "Liste hiér."
                ElseIf g_strLstSpécial = "%NUMFCT" Then
                    cmd(CMD_CHOIX_VAL).Visible = True
                    cmd(CMD_CHOIX_VAL).Caption = "Fonctions"
                ElseIf g_strLstSpécial = "%ENTIER" Then
                    cmd(CMD_CHOIX_VAL).Caption = "Nombre Entier"
                ElseIf g_strLstSpécial = "calculer" Then
                    cmd(CMD_CHOIX_VAL).Caption = "Champ calculé"
                    g_boolCalcul = True
                ElseIf g_strLstSpécial = "%MTT" Then
                    cmd(CMD_CHOIX_VAL).Caption = "Montant"
                End If
            Else
                cmd(CMD_CHOIX_VAL).tag = rs("forec_valeurs_possibles")
                cmd(CMD_CHOIX_VAL).Caption = FctRecupNomListe(rs("forec_valeurs_possibles"))
            End If
        End If
    End If
    
    Me.cmd(2).Visible = False
    Me.FrmNbOccur.Visible = False
    Select Case rs("forec_type")
    Case "TEXT"
        g_boolTexte = True
        sFormule = IIf(IsNull(rs("forec_formule").Value), "", rs("forec_formule").Value)
        If InStr(rs("forec_fctvalid"), "DATE") > 0 Then
            g_boolDate = True
            Me.Frame.Caption = Me.Frame.Caption & "   (Date)"
        ElseIf InStr(rs("forec_fctvalid"), "ENTIER") > 0 Then
            g_BoolEntier = True
            Me.Frame.Caption = Me.Frame.Caption & "   (Entier)"
        ElseIf InStr(rs("forec_fctvalid"), "MTT") > 0 Then
            g_BoolMTT = True
            Me.Frame.Caption = Me.Frame.Caption & "   (Montant)"
        ElseIf g_strLstSpécial = "%NUMSERVICE" Then
            Me.FrmNbOccur.Visible = True
            Me.cmd(2).Visible = True
            Me.Frame.Caption = Me.Frame.Caption & "   (Services)"
        ElseIf g_strLstSpécial = "%NUMFCT" Then
            Me.FrmNbOccur.Visible = True
            Me.cmd(2).Visible = True
            Me.Frame.Caption = Me.Frame.Caption & "   (Fonctions)"
        ElseIf g_boolCalcul Then
            Me.Frame.Caption = Me.Frame.Caption & "   (Champ calculé)"
        End If
    Case "CHECK"
        Me.FrmNbOccur.Visible = True
        Me.cmd(2).Visible = True
        g_boolListe = True
        Me.Frame.Caption = Me.Frame.Caption & "   (Cases à cocher)"
    Case "RADIO"
        Me.FrmNbOccur.Visible = True
        Me.cmd(2).Visible = True
        g_boolListe = True
        Me.Frame.Caption = Me.Frame.Caption & "   (Boutons d'option)"
    Case "SELECT"
        Me.FrmNbOccur.Visible = True
        Me.cmd(2).Visible = True
        g_boolListe = True
        Me.Frame.Caption = Me.Frame.Caption & "   (Liste déroulante)"
    Case "HIERARCHIE"
        Me.FrmNbOccur.Visible = True
        Me.cmd(2).Visible = True
        g_boolListe = True
        Me.Frame.Caption = Me.Frame.Caption & "   (Liste hiér.)"
    End Select
    
    ' Affiner selon la mise en forme
    Me.FrmRepartDates.Visible = False
    If g_MenForme = "" Then
        If g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_strLstSpécial = "HIERARCHIE" Or g_boolListe Then
            OptLigCol(0).Value = True ' en ligne
            Call MettreModeLigCol("L")
            Me.OptLibVal(Opt_LibVal_LV).Value = True
            If g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_strLstSpécial = "HIERARCHIE" Then
                Me.OptTypVal(OptTypVal_ValeurSignif).Value = True
            ElseIf g_boolListe Then
                Me.OptTypVal(OptTypVal_Valeur).Value = True
            Else
                MsgBox "Case ?"
            End If
            Me.OptValForme(Opt_ValForme_Nombre).Value = True
            Me.FrmNbOccur.Visible = True
            Me.LblNbreOccur.Visible = True
            Me.TxtNbOccur.Visible = False
            Me.TxtNbOccur.Text = 0
            Me.OptNbOccur(Opt_nboccur_Toutes).Value = True
        Else    ' Entier ou MTT ou calcul
            OptLigCol(1).Value = True ' en colonne
            Call MettreModeLigCol("C")
            Me.OptLibVal(Opt_LibVal_LV).Value = True
            Me.OptTypVal(OptTypVal_Valeur_Rens).Value = True
            Me.OptValForme(Opt_ValForme_Nombre).Value = True
            Me.OptNbOccur(Opt_nboccur_Toutes).Value = True
            If g_boolDate Then
                Me.FrmRepartDates.Visible = True
                Me.CmbRepartDates.AddItem "par semaine"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 0
                Me.CmbRepartDates.AddItem "par mois"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 1
                Me.CmbRepartDates.AddItem "par année"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 2
                Me.CmbRepartDates.ListIndex = 1
            End If
        End If
    Else
LabCasGeneral:
        ' selon type de champs
        ' 0 Colonne_Lib_Val#
        ' 1 forme du retour ( somme, pourcent,...)
        ' 2 forme de la présentation (valeur,valeur signif, valeurs renseignées, ...)
        ' 3 par rapport à
        ' 4 * ou (nb occurences)

        ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
        ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
        s = STR_GetChamp(g_MenForme, "#", 0)    ' libelle ou valeur
        If g_MenForme = "" Or s = "Ligne_Lib_Val" Then
            Me.OptLigCol(0).Value = 1
            Me.OptLibVal(Opt_LibVal_LV).Value = True
        ElseIf s = "Ligne_Val" Then
            Me.OptLigCol(0).Value = 1
            Me.OptLibVal(Opt_LibVal_V).Value = True
        ElseIf s = "Ligne_Lib" Then
            Me.OptLigCol(0).Value = 1
            Me.OptLibVal(Opt_LibVal_L).Value = True
        ElseIf s = "Colonne_Lib_Val" Then
            Me.OptLigCol(1).Value = 1
            Me.OptLibVal(Opt_LibVal_LV).Value = True
        ElseIf s = "Colonne_Val" Then
            Me.OptLigCol(1).Value = 1
            Me.OptLibVal(Opt_LibVal_V).Value = True
        ElseIf s = "Colonne_Lib" Then
            Me.OptLigCol(1).Value = 1
            Me.OptLibVal(Opt_LibVal_L).Value = True
        Else
            Me.OptLibVal(Opt_LibVal_LV).Value = True
        End If
        Me.FrmLibVal.Visible = True
        
        Me.FrmRepartDates.Visible = False
        If STR_GetChamp(g_MenForme, "#", 2) = "NOMBRE_TOTAL" Then
            cmd(CMD_CHOIX_VAL).Visible = False
            FrmSQL.Visible = False
        Else
            If InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
                Me.FrmRepartDates.Visible = True
                Me.CmbRepartDates.AddItem "par Jour calendaire"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 0
                Me.CmbRepartDates.AddItem "par semaine"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 1
                Me.CmbRepartDates.AddItem "par mois"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 2
                Me.CmbRepartDates.AddItem "par trimestre"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 3
                Me.CmbRepartDates.AddItem "par année"
                Me.CmbRepartDates.ItemData(Me.CmbRepartDates.ListCount - 1) = 4
                s = STR_GetChamp(g_MenForme, "#", 6)
                If s = "J" Then
                    Me.CmbRepartDates.ListIndex = 0
                ElseIf s = "S" Then
                    Me.CmbRepartDates.ListIndex = 1
                ElseIf s = "M" Then
                    Me.CmbRepartDates.ListIndex = 2
                ElseIf s = "T" Then
                    Me.CmbRepartDates.ListIndex = 3
                ElseIf s = "A" Then
                    Me.CmbRepartDates.ListIndex = 4
                End If
            End If
        End If
        ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
        ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
        s = STR_GetChamp(g_MenForme, "#", 1)
        ' Pourcent, somme, ...
        Me.FrmParRapport.Visible = False
        If s = "POURCENT" Then
            OptValForme(Opt_ValForme_Pourcent).Value = True
            Me.FrmParRapport.Visible = True
        ElseIf s = "SOMME" Then
            OptValForme(Opt_ValForme_Somme).Value = True
        ElseIf s = "NOMBRE" Then
            OptValForme(Opt_ValForme_Nombre).Value = True
        ElseIf s = "ECART_TYPE" Then
            OptValForme(Opt_ValForme_EcarType).Value = True
        ElseIf s = "MOYENNE" Then
            Me.FrmParRapport.Visible = True
            OptValForme(Opt_ValForme_Moyenne).Value = True
        Else
            MsgBox "Case ?"
        End If

        ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
        ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
        s = STR_GetChamp(g_MenForme, "#", 2)
        ' valeur , valeurs significatives, ...
        If s = "VALEUR" Then
            Me.OptTypVal(OptTypVal_Valeur).Value = True
        ElseIf s = "VALEUR_SIGNIF" Then
            Me.OptTypVal(OptTypVal_ValeurSignif).Value = True
        ElseIf s = "NONVALEUR_R" Then
            Me.OptTypVal(OptTypVal_Valeur_Rens).Value = True
        ElseIf s = "NONVALEUR_NR" Then
            Me.OptTypVal(OptTypVal_Valeur_NonRens).Value = True
        ElseIf s = "NOMBRE_TOTAL" Then
            Me.OptTypVal(OptTypVal_Valeur_Rens).Value = True
        Else
            MsgBox "Case ?"
        End If
        
        ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
        ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
        s = STR_GetChamp(g_MenForme, "#", 5)
        ' par rapport à ?
        If s = "TOUTES" Or s = "" Then
            Me.OptRapportVal(OptRapportVal_Toutes).Value = True
        ElseIf s = "AUX_R" Then
            Me.OptRapportVal(OptRapportVal_Rens).Value = True
        Else
            MsgBox "Case ?"
        End If
 
        ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
        ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
        If g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_strLstSpécial = "HIERARCHIE" Or g_boolListe Then
            g_ValeurListe = STR_GetChamp(g_MenForme, "#", 4)
            Me.FrmLibVal.Visible = True
            Me.LabValChp.Visible = True
            s = STR_GetChamp(g_MenForme, "#", 4)
            If s = "TOUTES" Or s = "0" Then
                Me.LabValChp.Caption = "Toutes les Valeurs"
            Else
                Me.LabValChp.Caption = ""
                'If g_boolListe Then
                '    g_ValeurListe = Replace(g_ValeurListe, ";", "|")
                'End If
                For i = 0 To STR_GetNbchamp(g_ValeurListe, "|")
                    laS = STR_GetChamp(g_ValeurListe, "|", i)
                    If laS <> "" Then
                        If Not g_boolListe Then
                            laS = "S" & Mid$(STR_GetChamp(laS, ";", STR_GetNbchamp(laS, ";") - 1), 2)
                        End If
                        Me.LabValChp.Caption = Me.LabValChp.Caption & IIf(Me.LabValChp.Caption = "", "", " OU ") & FctRecupNomValeur(laS)
                    End If
                Next i
            End If
            ' nombre d'occurences
            s = STR_GetChamp(g_MenForme, "#", 3)
            If s = "*" Then
                Me.TxtNbOccur.Text = "0"
                Me.OptNbOccur(Opt_nboccur_Toutes).Value = True
            Else
                Me.TxtNbOccur.Text = s
                Me.OptNbOccur(Opt_nboccur_Choisir).Value = True
            End If
        Else
            Me.LabValChp.Visible = False
            Me.FrmLibVal.Visible = True
        End If
        
        If g_MenForme = "" Or Mid(g_MenForme, 1, 5) = "Ligne" Then
            Me.OptLigCol(0).Value = True
            Call MettreModeLigCol("L")
        Else
            Me.OptLigCol(1).Value = True
            Call MettreModeLigCol("C")
        End If
    End If
    
    If g_chpnum <= -10 Then
        g_BoolEntier = True
        GoTo LabSuite1
    End If
    ' mettre le nom du formulaire
    s_chp_nom = rs("forec_label")
    sql = "select * from formulaire where for_num=" & rs("forec_fornum")
    If Odbc_Select(sql, rs) = P_ERREUR Then
        MsgBox "PrmFormatChp : formulaire " & rs("forec_fornum") & " introuvable"
        Exit Sub
    Else
        s_form_nom = rs("for_lib")
    End If
    Me.OptRapportVal(OptRapportVal_Rens).Caption = "Les [" & s_form_nom & "] dont le champ [" & s_chp_nom & "] est renseigné"
    Me.OptRapportVal(OptRapportVal_Toutes).Caption = "Tous les " & s_form_nom

LabSuite1:
    Call MenForme("Init", "")
    

    ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
    ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
    If g_boolLstSpécial Then
        If cmd(CMD_CHOIX_VAL).tag <> "" Then
            ret = RemplirListe(cmd(CMD_CHOIX_VAL).tag, Replace(g_ValeurListe, "S", ""))
        End If
    ElseIf g_boolTexte Then
        Me.ListVal.Visible = False
        Me.LabValChp.Visible = False
    ElseIf g_boolListe Then
        If g_ValeurListe = "" Then g_ValeurListe = "TOUTES"
        If cmd(CMD_CHOIX_VAL).tag <> "" Then
            ret = RemplirListe(cmd(CMD_CHOIX_VAL).tag, g_ValeurListe)
            Me.ListVal.Visible = True
        End If
        If g_ValeurListe <> "" Then
            Me.LabValChp.Visible = True
            If g_ValeurListe = "" Or g_ValeurListe = "TOUTES" Then
                Me.LabValChp.Caption = "Toutes les Valeurs"
            Else
                Me.LabValChp.Caption = "Valeur : " & FctRecupNomValeur(g_ValeurListe)
            End If
        Else
            Me.LabValChp.Visible = True
            Me.LabValChp.Caption = "Toutes les Valeurs"
        End If
        'Me.FrmLibVal.Visible = True
        'Me.FrmTypeVal.Visible = True
        'Me.FrmVal.Visible = True
    End If
    
    If g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_strLstSpécial = "HIERARCHIE" Or g_boolListe Then
        '
    Else
        Me.FrmNbOccur.Visible = False
        Me.OptNbOccur(Opt_nboccur_Toutes).Value = True
        Me.TxtNbOccur.Text = 0
    End If
    '
Fin:
    If g_chpnum <= -10 Then
        Call TxtSQL_Change
    End If
    If STR_GetChamp(g_MenForme, "#", 2) = "NOMBRE_TOTAL" Then
        cmd(CMD_CHOIX_VAL).Visible = False
        FrmSQL.Visible = False
    End If
    g_mode_saisie = True
    cmd(CMD_OK).Visible = True
    If p_changement_de_champ Or g_Trait = "Ajout" Or g_Trait = "AjoutPlusieurs" Then
        cmd(CMD_OK).Enabled = True
    Else
        cmd(CMD_OK).Enabled = False
    End If
    
End Sub

Private Sub MenForme(ByVal v_type As String, ByVal v_Trait As String)
    ' OptVal_Click
    Dim s As String
    Dim ret As String

    
    If (g_boolDate) Then
        If OptTypVal(OptTypVal_Valeur_NonRens).Value Then  ' non renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = False
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = True
            OptValForme(Opt_ValForme_Moyenne).Visible = True
            OptTypVal(OptTypVal_Valeur).Visible = False
            OptTypVal(OptTypVal_ValeurSignif).Visible = False
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        ElseIf OptTypVal(OptTypVal_Valeur_Rens).Value Then  ' renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = True
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = True
            OptValForme(Opt_ValForme_Moyenne).Visible = True
            OptTypVal(OptTypVal_Valeur).Visible = True                  ' False
            OptTypVal(OptTypVal_ValeurSignif).Visible = True            ' False
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        ElseIf OptTypVal(OptTypVal_Valeur).Value Or OptTypVal(OptTypVal_ValeurSignif).Value Then  ' choix gauche (les valeurs)
            OptValForme(Opt_ValForme_EcarType).Visible = False
            OptValForme(Opt_ValForme_Moyenne).Visible = False
        End If
        Me.ListVal.Visible = False   ' liste des valeurs
        Me.LabValChp.Visible = False
    ElseIf (g_BoolEntier Or g_BoolMTT Or g_boolCalcul) Then
        If OptTypVal(OptTypVal_Valeur_NonRens).Value Then  ' non renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = False
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = True
            OptValForme(Opt_ValForme_Moyenne).Visible = True
            OptTypVal(OptTypVal_Valeur).Visible = False
            OptTypVal(OptTypVal_ValeurSignif).Visible = False
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        ElseIf OptTypVal(OptTypVal_Valeur_Rens).Value Then  ' renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = True
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = True
            OptValForme(Opt_ValForme_Moyenne).Visible = True
            OptTypVal(OptTypVal_Valeur).Visible = True                  ' False
            OptTypVal(OptTypVal_ValeurSignif).Visible = True            ' False
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        ElseIf OptTypVal(OptTypVal_Valeur).Value Or OptTypVal(OptTypVal_ValeurSignif).Value Then  ' choix gauche (les valeurs)
            OptValForme(Opt_ValForme_EcarType).Visible = False
            OptValForme(Opt_ValForme_Moyenne).Visible = False
        End If
        Me.ListVal.Visible = False   ' liste des valeurs
        Me.LabValChp.Visible = False
    ElseIf (g_strLstSpécial = "%NUMSERVICE" Or g_strLstSpécial = "%NUMFCT" Or g_strLstSpécial = "HIERARCHIE" Or g_boolListe) Then
        Me.FrmNbOccur.Visible = True
        If OptTypVal(OptTypVal_Valeur_NonRens).Value Then  ' non renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = False
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = False
            OptValForme(Opt_ValForme_Moyenne).Visible = False
            OptTypVal(OptTypVal_Valeur).Visible = True
            OptTypVal(OptTypVal_ValeurSignif).Visible = True
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        ElseIf OptTypVal(OptTypVal_Valeur_Rens).Value Then  ' renseignées
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = False
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = False
            OptValForme(Opt_ValForme_Moyenne).Visible = False
            OptTypVal(OptTypVal_Valeur).Visible = True
            OptTypVal(OptTypVal_ValeurSignif).Visible = True
            OptTypVal(OptTypVal_Valeur_Rens).Visible = True
            OptTypVal(OptTypVal_Valeur_NonRens).Visible = True
        Else
            OptValForme(Opt_ValForme_Pourcent).Visible = True
            OptValForme(Opt_ValForme_Somme).Visible = False
            OptValForme(Opt_ValForme_Nombre).Visible = True
            OptValForme(Opt_ValForme_EcarType).Visible = False
            OptValForme(Opt_ValForme_Moyenne).Visible = False
        End If
        Me.ListVal.Visible = True   ' liste des valeurs
        Me.LabValChp.Visible = True
    Else
        MsgBox "case ?"
    End If

End Sub

Function RemplirListe(v_numlst, g_ValeurListe)
    Dim sql As String, rs As rdoResultset
    Dim lig As Integer
    Dim j As Integer
    Dim strNom As String
    
    Me.ListVal.Visible = True
    If v_numlst = "%NUMSERVICE" Then
        If val(g_ValeurListe) > 0 Then
            Call P_RecupSrvNom(val(g_ValeurListe), strNom)
            Me.LabValChp.Caption = strNom
        Else
            Me.LabValChp.Visible = True
            Me.LabValChp.Caption = "Tous les Services"
        End If
    ElseIf left$(v_numlst, 10) = "HIERARCHIE" Then
        If val(Mid$(g_ValeurListe, 2)) > 0 Then
            strNom = FctRecupNomValeur(g_ValeurListe)
            Me.LabValChp.Caption = strNom
        Else
            Me.LabValChp.Visible = True
            Me.LabValChp.Caption = "Toutes les valeurs"
        End If
    ElseIf v_numlst = "%NUMFCT" Then
        If val(g_ValeurListe) > 0 Then
            strNom = ChercheNomFonction(g_ValeurListe)
            Me.LabValChp.Caption = strNom
        Else
            Me.LabValChp.Visible = True
            Me.LabValChp.Caption = "Toutes les Fonctions"
        End If
    ElseIf v_numlst = "%ENTIER" Then
        Me.LabValChp.Caption = "Nombre entier"
        Me.ListVal.Visible = False
    ElseIf v_numlst = "calculer" Then
        Me.LabValChp.Caption = "Champ calculé"
        Me.ListVal.Visible = False
    ElseIf v_numlst = "%MTT" Then
        Me.LabValChp.Caption = "Montant"
        Me.ListVal.Visible = False
    Else
        ' ******************************
        ' Cas des check - radio - select
        ' ******************************
        ' celles sélectionnés ont un X
        sql = "select * from valchp where vc_lvcnum=" & v_numlst & " order by vc_ordre"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        Me.ListVal.Clear
        Me.ListVal.Columns = 2
        Me.ListVal.AddItem "" & vbTab & "-->  Toutes  <--"
        'lig = 1
        Dim n As Integer, i As Integer, laVal As String, bok As Boolean
        FaireListeClick = False
        Do While Not rs.EOF
            Me.ListVal.AddItem " " & vbTab & rs("vc_lib")
            Me.ListVal.selected(Me.ListVal.ListCount - 1) = False
            Me.ListVal.ItemData(Me.ListVal.ListCount - 1) = rs("vc_num")
            rs.MoveNext
        Loop
        If g_ValeurListe = "TOUTES" Then
            Me.ListVal.selected(0) = True
        Else
            bok = False
            For i = 0 To Me.ListVal.ListCount - 1
                'bOK = False
                n = STR_GetNbchamp(g_ValeurListe, ";") - 1
                For j = 0 To n
                    laVal = STR_GetChamp(g_ValeurListe, ";", j)
                    If laVal = Me.ListVal.ItemData(i) Then
                        Me.ListVal.selected(i) = True
                        Me.ListVal.selected(0) = False
                        bok = True
                        Exit For
                    End If
                Next j
            Next i
        End If
        If Not bok Then     ' aucun => Toutes
            Me.ListVal.selected(0) = True
        End If
        FaireListeClick = True
    End If
End Function

Function ChoisirFonction(ByVal v_numfct)
    Dim sql As String, rs As rdoResultset
    Dim frm As Form, sret As String
    Dim n As Integer
    Dim i As Integer
    
    Call CL_Init
        
    Set frm = KS_PrmFonction
    sret = KS_PrmFonction.AppelFrm(v_numfct)
    Set frm = Nothing
    If sret = "-1" Then
        ChoisirFonction = "-1"
        Exit Function
    Else
        ChoisirFonction = sret
    End If

End Function

Function ChoisirService(ByVal v_numsrv)
    Dim sql As String, rs As rdoResultset
    Dim frm As Form, sret As String
    Dim s As String
    Dim n As Integer
    Dim i As Integer
    Dim laS As String
    Dim iret As Integer
    
    
    sql = "select EJ_Num from EntJuridique"
    If Odbc_RecupVal(sql, p_num_ent_juridique) = P_ERREUR Then
        ChoisirService = P_ERREUR
        Exit Function
    End If

    Call CL_Init
    If v_numsrv = "TOUTES" Then
        Call CL_AddLigne("L" & p_num_ent_juridique, p_num_ent_juridique, "", True)
    Else
        i = 0
        For n = 0 To STR_GetNbchamp(v_numsrv, "|")
            laS = STR_GetChamp(v_numsrv, "|", n)
            If laS <> "" And laS <> "<NR>" Then
                Call CL_AddLigne(laS, i, "", True)
                i = i + 1
            End If
        
            'laS = STR_GetChamp(v_numsrv, "|", n)
            'If laS <> "" And laS <> "<NR>" Then
            '    iret = Mid$(STR_GetChamp(laS, ";", STR_GetNbchamp(laS, ";") - 1), 2)
            '    Call CL_AddLigne("S" & iret, i, "", True)
            '    i = i + 1
            'End If
        Next n
    End If
    Set frm = KS_PrmService
    sret = KS_PrmService.AppelFrm("Choix des services", "C", True, "", "S", False)
    Set frm = Nothing
    
    If sret = "" Then
        ChoisirService = 0
        Exit Function
    ElseIf sret = "N0" Then
        ChoisirService = "L" & p_num_ent_juridique
    Else
        'ChoisirService = Replace(sret, "S", "")
        'ChoisirService = Replace(ChoisirService, ";", "")
        ChoisirService = ""
        For i = 0 To UBound(CL_liste.lignes)
            ChoisirService = ChoisirService & CL_liste.lignes(i).texte & "|"
        Next i
    End If
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Function

Private Function ChoisirValeurHierar(ByVal v_numlst_Valeur As String) As String

    Dim sval As String, sep As String, sep2 As String, s As String
    Dim n As Integer, i As Integer, n2 As Integer
    Dim numlst As Long
    Dim ValeurListe As String
    
    numlst = STR_GetChamp(v_numlst_Valeur, "|", 0)
    ValeurListe = STR_GetChamp(v_numlst_Valeur, "|", 1)
    
    ChoisirValeurHierar = ""
    
    Call CL_Init
    
    Call CL_AddLigne("<Non renseigné>", 0, 0, False)
    sval = ""
    sval = ValeurListe
    Call ajouter_hierar_fils(-numlst, sval, "", 0)
    If UBound(CL_liste.lignes()) = 1 Then
        Call MsgBox("Aucune valeur n'a été trouvée.", vbInformation + vbOKOnly, "")
        ChoisirValeurHierar = ""
        Exit Function
    End If
    Call CL_InitTitreHelp("Liste des valeurs", "")
    If UBound(CL_liste.lignes) < 20 Then
        n = UBound(CL_liste.lignes) + 1
    Else
        n = 20
    End If
    Call CL_InitTaille(0, -n)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        ChoisirValeurHierar = ""
        Exit Function
    End If
    
    If CL_liste.retour = 0 Then
        ChoisirValeurHierar = "M" & CL_liste.lignes(CL_liste.pointeur).num
    End If

End Function

Function FctRecupNomValeur(v_valchp)
    Dim sql As String, rs As rdoResultset
    Dim NumValeur As String, n As Integer
    Dim i As Integer
    Dim val As String
    Dim s As String
    Dim op As String
    
    If v_valchp = "TOUTES" Then
        FctRecupNomValeur = "TOUTES"
        Exit Function
    End If
    n = STR_GetNbchamp(v_valchp, ";")
    For i = 0 To n
        val = STR_GetChamp(v_valchp, ";", i)
        If val <> "" Then
            If left$(val, 1) = "S" Then
                sql = "select srv_nom from service where srv_num=" & Mid$(val, 2)
            ElseIf left$(val, 1) = "M" Then
                sql = "select hvc_nom from hierarvalchp where hvc_num=" & Mid$(val, 2)
            Else
                sql = "select * from valchp where vc_num=" & val
            End If
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Function
            End If
            If Not rs.EOF Then
                s = s & op & "[" & rs(0).Value & "]"
                op = " Ou "
            End If
        End If
    Next i
    FctRecupNomValeur = s
End Function

Function FctRecupNomListe(v_numliste)
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from lstvalchp where lvc_num=" & v_numliste
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If Not rs.EOF Then
        FctRecupNomListe = rs("lvc_nom")
    End If

End Function

Private Function quitter(ByVal v_bforce As Boolean) As Boolean

    Dim reponse As Integer
    
    If Not v_bforce Then
        If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
            reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            If reponse = vbNo Then
                quitter = False
                Exit Function
            End If
        End If
    End If
    
    ' retourner à l'appelant
    g_retour_PrmFormatChp = "QUITTER"
    
    Unload Me
    
    quitter = True
    
End Function



Private Function VerifSiChange()
        
    If g_MenForme <> CalCul_MenF() Then
        VerifSiChange = True
    Else
        VerifSiChange = False
    End If
        
End Function

Private Sub ajouter_hierar_fils(ByVal v_numval As Long, _
                                ByVal v_LstVal As String, _
                                ByVal v_shierar As String, _
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
            s = v_shierar & "M" & rs("HVC_Num").Value & ";"
            If InStr(s, v_LstVal & ";") > 0 Then
                trouve = True
                v_LstVal = "" ' on en prend qu'un seul (le père)
            End If
        End If
        Call CL_AddLigne(sdecal & rs("HVC_nom").Value, rs("HVC_Num").Value, s, trouve)
        Call ajouter_hierar_fils(rs("HVC_Num").Value, v_LstVal, s, v_niveau + 1)
        rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Function CalCul_MenF()

    Dim MenForme1 As String, MenForme2 As String, MenForme3 As String
    Dim MenForme4 As String, MenForme5 As String
    Dim MenForme6 As String
    Dim MenForme7 As String
    
    ' Calculer la mise en forme
    MenForme1 = ""  ' Ligne ou colonne ? / Libelle, valeur ou les 2 ?
    If Me.OptLigCol(0).Value = True Then   ' En ligne
        If Me.OptLibVal(Opt_LibVal_LV).Value = True Then
            MenForme1 = "Ligne_Lib_Val"
        ElseIf Me.OptLibVal(Opt_LibVal_L).Value = True Then
            MenForme1 = "Ligne_Lib"
        ElseIf Me.OptLibVal(Opt_LibVal_V).Value = True Then
            MenForme1 = "Ligne_Val"
        End If
    Else
        If Me.OptLibVal(Opt_LibVal_LV).Value = True Then
            MenForme1 = "Colonne_Lib_Val"
        ElseIf Me.OptLibVal(Opt_LibVal_L).Value = True Then
            MenForme1 = "Colonne_Lib"
        ElseIf Me.OptLibVal(Opt_LibVal_V).Value = True Then
            MenForme1 = "Colonne_Val"
        End If
    End If
        
    ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
    ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
    ' Format de valeurs
    MenForme2 = ""  ' Forme : Pourcent, Nombre, ...
    If Me.OptValForme(Opt_ValForme_Pourcent).Value = True Then
        MenForme2 = "POURCENT"
    ElseIf Me.OptValForme(Opt_ValForme_Somme).Value = True Then
        MenForme2 = "SOMME"
    ElseIf Me.OptValForme(Opt_ValForme_Nombre).Value = True Then
        MenForme2 = "NOMBRE"
    ElseIf Me.OptValForme(Opt_ValForme_EcarType).Value = True Then
        MenForme2 = "ECART_TYPE"
    ElseIf Me.OptValForme(Opt_ValForme_Moyenne).Value = True Then
        MenForme2 = "MOYENNE"
    End If
    
    ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
    ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
    ' Type de valeurs : ValeurSignif, Valeurs, Valeur_Rens, Valeur_nonRens
    MenForme3 = ""
    If Me.OptTypVal(OptTypVal_Valeur).Value = True Then
        MenForme3 = "VALEUR"
    ElseIf Me.OptTypVal(OptTypVal_ValeurSignif).Value = True Then
        MenForme3 = "VALEUR_SIGNIF"
    ElseIf Me.OptTypVal(OptTypVal_Valeur_Rens).Value = True Then
        MenForme3 = "NONVALEUR_R"
    ElseIf Me.OptTypVal(OptTypVal_Valeur_NonRens).Value = True Then
        MenForme3 = "NONVALEUR_NR"
    End If
    
    ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
    ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
    ' nombre d'occurence
    MenForme4 = ""
    If Me.OptNbOccur(Opt_nboccur_Toutes).Value Then  ' Toutes
        MenForme4 = "*"
        'g_ValeurListe = ""
    Else
        MenForme4 = Me.TxtNbOccur.Text
        If MenForme4 = "" Then MenForme4 = "*"
    End If
    
    ' Colonne_Val # POURCENT # NONVALEUR_R # * #             # TOUTES    ' pour Entier
    ' Colonne_Val # POURCENT # VALEUR_R    # 5 # 121;114;118 # AUX_R     ' pour Liste
    '   g_ValeurListe
    
    ' Par rapport à ?
    MenForme5 = ""
    If Me.FrmParRapport.Visible Then
        If Me.OptRapportVal(OptRapportVal_Rens).Value Then
            MenForme5 = "AUX_R"
        Else
            MenForme5 = "TOUTES"
        End If
    Else
        MenForme5 = ""
    End If
    
    If Me.FrmRepartDates.Visible Then
        If Me.CmbRepartDates.ListIndex = 0 Then
            MenForme6 = "J"
        ElseIf Me.CmbRepartDates.ListIndex = 1 Then
            MenForme6 = "S"
        ElseIf Me.CmbRepartDates.ListIndex = 2 Then
            MenForme6 = "M"
        ElseIf Me.CmbRepartDates.ListIndex = 3 Then
            MenForme6 = "T"
        ElseIf Me.CmbRepartDates.ListIndex = 4 Then
            MenForme6 = "A"
        Else
            MenForme6 = "M"
        End If
    End If
    
    If Me.FrmNiveauStru.Visible Then
        Dim TypeNiveau As String
        If Me.CmbTypeStru.ItemData(Me.CmbTypeStru.ListIndex) > 0 Then
            TypeNiveau = "S"
            MenForme7 = Me.CmbTypeStru.ItemData(Me.CmbTypeStru.ListIndex)
            MenForme7 = MenForme7 & IIf(Me.chk_niveau_exact_S.Value = 1, "O", "N") & TypeNiveau
        ElseIf Me.CmbNivStru.ItemData(Me.CmbNivStru.ListIndex) > 0 Then
            TypeNiveau = "N"
            MenForme7 = Me.CmbNivStru.ItemData(Me.CmbNivStru.ListIndex)
            MenForme7 = MenForme7 & IIf(Me.chk_niveau_exact_S.Value = 1, "O", "N") & TypeNiveau
        Else
            TypeNiveau = ""
        End If
    End If
    If Me.FrmNivHier.Visible Then
        MenForme7 = Me.CmbNivHier.ItemData(Me.CmbNivHier.ListIndex)
        MenForme7 = MenForme7 & IIf(Me.chk_niveau_exact_H.Value = 1, "O", "N")
        MenForme7 = MenForme7 & "S"
    End If
    
    If g_ValeurListe = "" Then g_ValeurListe = "TOUTES"
    If g_chpnum <= -10 Then
        CalCul_MenF = "Ligne_Val#NOMBRE#NOMBRE_TOTAL#*#TOUTES#"
    Else
        CalCul_MenF = MenForme1 & "#" & MenForme2 & "#" & MenForme3 & "#" & MenForme4 & "#" & g_ValeurListe & "#" & MenForme5 & "#" & MenForme6 & "#" & MenForme7
    End If

End Function

Private Sub valider()
    Dim cr As Integer
    Dim MenForme1 As String, MenForme2 As String, MenForme3 As String
    Dim MenForme4 As String, MenForme5 As String
    Dim MenForme As String
    Dim rs As rdoResultset
    
    cr = verifier_tous_chp()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_NON Then
        Exit Sub
    End If
    
    ' Calculer la mise en forme
    MenForme = CalCul_MenF()
    
    ' retourner à l'appelant
    g_retour_PrmFormatChp = MenForme

    'MsgBox g_retour_PrmFormatChp
    If g_chpnum <= -10 Then
        If Me.TxtSQL.Text <> "" Then
            Me.TxtSQL.Text = Replace(Me.TxtSQL.Text, "[FOR_NUM]", g_fornum)
            Me.TxtSQL.Text = Replace(Me.TxtSQL.Text, "SQL=", "")
            If Odbc_SelectV(Me.TxtSQL.Text, rs) = P_ERREUR Then
                Exit Sub
            End If
            g_retour_PrmFormatChp = "SQL=" & Me.TxtSQL.Text
        End If
    End If
    
    'If Me.CmbNiveau.tag <> "" Then
    '    tbl_fichExcel(g_i_tbExcel).CmdNiveauRelier = Me.CmbNiveau.tag
    'End If
    
    Unload Me
    Exit Sub
    
err_enreg:
    Unload Me
    
End Sub

Private Function verifier_tous_chp() As Integer
    
    Dim ret As Integer
        
    verifier_tous_chp = P_OUI
    
End Function

Private Sub chk_niveau_exact_H_Click()
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
End Sub

Private Sub chk_niveau_exact_S_Click()
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
End Sub

Private Sub CmbNivHier_Click()
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
End Sub

Private Sub CmbNivStru_Click()
    If g_mode_saisie Then
        If Me.CmbNivStru.ListIndex > 0 Then
            Me.CmbTypeStru.ListIndex = 0
        End If
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
End Sub

Private Sub CmbRepartDates_Click()
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub CmbTypeStru_Click()
    If g_mode_saisie Then
        If Me.CmbTypeStru.ListIndex > 0 Then
            Me.CmbNivStru.ListIndex = 0
        End If
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub cmd_Click(Index As Integer)

    Dim ret As String
    Dim numlig As Integer
    Dim position As Integer
    Dim frm As Form
    Dim chpnum As Long
    Dim sret As String
    Dim numaction As Integer
    Dim strD As String, strG As String
    Dim strGPF As String, strDPF As String
    Dim strSel As String
    Dim strserv As String
    Dim str_action As String
    Dim iRound As Integer
    Dim i As Integer
    Dim newlen As Integer
    Dim Anc_Selstart As Integer
    Dim strNom As String
    Dim iret As Integer
    
    Select Case Index
    Case CMD_DELIER
        iret = MsgBox("Supprimer le lien ?", vbYesNo)
        If iret = vbYes Then
            tbl_fichExcel(g_i_tbExcel).CmdChpRelierà = 0
            cmd(CMD_CHP_RELIER).Visible = False
            cmd(CMD_DELIER).Visible = False
        End If
    Case CMD_CHP_RELIER
        iret = MsgBox("Supprimer le lien ?", vbYesNo)
        If iret = vbYes Then
            tbl_fichExcel(g_i_tbExcel).CmdChpRelierà = 0
            cmd(CMD_CHP_RELIER).Visible = False
        End If
    Case CMD_CHOIX_VAL
        If cmd(CMD_CHOIX_VAL).tag = "%NUMSERVICE" Then
            ret = ChoisirService(g_ValeurListe)
            If ret = "0" Then
                Me.LabValChp.Visible = True
                Exit Sub
            End If
            If ret <> "L" & p_num_ent_juridique Then
                ' recupérer le dernier
                Me.LabValChp.Caption = ""
                For i = 0 To STR_GetNbchamp(ret, "|")
                    strserv = STR_GetChamp(ret, "|", i)
                    If strserv <> "" Then
                        iret = Mid$(STR_GetChamp(strserv, ";", STR_GetNbchamp(strserv, ";") - 1), 2)
                        Call P_RecupSrvNom(iret, strNom)
                        Me.LabValChp.Caption = Me.LabValChp.Caption & IIf(Me.LabValChp.Caption <> "", " et ", "") & strNom
                    End If
                Next i
                g_ValeurListe = ret
                Me.LabValChp.Visible = True
            Else
                Me.LabValChp.Visible = True
                Me.LabValChp.Caption = "Tous les Services"
                g_ValeurListe = 0
            End If
            If g_mode_saisie Then
                If VerifSiChange() Then
                    cmd(CMD_OK).Enabled = True
                Else
                    cmd(CMD_OK).Enabled = False
                End If
            End If
        ElseIf cmd(CMD_CHOIX_VAL).tag = "%NUMFCT" Then
            ret = ChoisirFonction(g_ValeurListe)
            If ret = "-1" Then
                ' rien choisi
            ElseIf ret <> "" And ret <> "0" Then
                strNom = ChercheNomFonction(ret)
                Me.LabValChp.Caption = strNom
                g_ValeurListe = ret
                Me.LabValChp.Visible = True
            Else
                Me.LabValChp.Visible = True
                Me.LabValChp.Caption = "Toutes les Fonctions"
                g_ValeurListe = "TOUTES"
            End If
            If g_mode_saisie Then
                If VerifSiChange() Then
                    cmd(CMD_OK).Enabled = True
                Else
                    cmd(CMD_OK).Enabled = False
                End If
            End If
        ElseIf left$(cmd(CMD_CHOIX_VAL).tag, 10) = "HIERARCHIE" Then
            ret = ChoisirValeurHierar(STR_GetChamp(cmd(CMD_CHOIX_VAL).tag, "%", 1) & "|" & g_ValeurListe)
            If ret <> "" Then
                strNom = FctRecupNomValeur(ret)
                Me.LabValChp.Caption = strNom
                g_ValeurListe = ret
                Me.LabValChp.Visible = True
            Else
                Me.LabValChp.Visible = True
                Me.LabValChp.Caption = "Toutes les valeurs"
                g_ValeurListe = 0
            End If
            If g_mode_saisie Then
                If VerifSiChange() Then
                    cmd(CMD_OK).Enabled = True
                Else
                    cmd(CMD_OK).Enabled = False
                End If
            End If
        End If
    Case CMD_OK
        Call valider
        Exit Sub
    Case CMD_QUITTER
        Call quitter(False)
        Exit Sub
    End Select
    
End Sub

Private Function ChercheNomFonction(ByVal v_fctnum)
    Dim sql As String, rs As rdoResultset
    Dim i As Integer, s As String
    Dim op As String
    
    op = ""
    For i = 0 To STR_GetNbchamp(v_fctnum, ";")
        s = STR_GetChamp(v_fctnum, ";", i)
        If s <> "" Then
            sql = "select * from fcttrav where ft_num = " & s
            If Odbc_Select(sql, rs) = P_ERREUR Then
                ChercheNomFonction = P_ERREUR
                Exit Function
            End If
            If Not rs.EOF Then
                ChercheNomFonction = ChercheNomFonction & op & rs("ft_libelle")
                op = " / "
            End If
        End If
    Next i
    
End Function

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_QUITTER Then
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

Private Sub ListVal_Click()
    Dim ret As String
    Dim Anc As String
    Dim i As Integer, op As String, bok As Boolean, sOp As String
    Dim sql As String, rs As rdoResultset
    
    If Not FaireListeClick Then Exit Sub

    Anc = g_ValeurListe
    
    If Me.ListVal.selected(0) Then
LabToutes:
        g_ValeurListe = "TOUTES"
        FaireListeClick = False
        For i = 1 To Me.ListVal.ListCount - 1
            Me.ListVal.selected(i) = False
        Next i
        FaireListeClick = True
    Else
        bok = False
        op = ""
        sOp = ""
        g_ValeurListe = ""
        For i = 1 To Me.ListVal.ListCount - 1
            If Me.ListVal.selected(i) Then
                g_ValeurListe = g_ValeurListe & op & Me.ListVal.ItemData(i)
                op = ";"
                bok = True
            End If
        Next i
    End If
    
    If g_ValeurListe = "TOUTES" Then
        Me.LabValChp.Caption = "Toutes les valeurs"
    Else
        Me.LabValChp.Caption = "Valeur : " & FctRecupNomValeur(g_ValeurListe)
    End If
    
    If Anc <> g_ValeurListe Then
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub OptLibVal_Click(Index As Integer)
    Dim i As Integer
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
    
    For i = 0 To Me.OptLibVal.Count - 1
        If Me.OptLibVal(i).Value Then
            Me.OptLibVal(i).ForeColor = COLOR_SEL
            Me.OptLibVal(i).FontBold = True
        Else
            Me.OptLibVal(i).ForeColor = &H80000012
            Me.OptLibVal(i).FontBold = False
        End If
    Next i
End Sub

Private Sub OptLigCol_Click(Index As Integer)
    Dim i As Integer
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
    
    For i = 0 To Me.OptLigCol.Count - 1
        If Me.OptLigCol(i).Value Then
            Me.OptLigCol(i).ForeColor = COLOR_SEL
            Me.OptLigCol(i).FontBold = True
        Else
            Me.OptLigCol(i).ForeColor = &H80000012
            Me.OptLigCol(i).FontBold = False
        End If
    Next i
    
    If g_mode_saisie Then       ' lignes ou colonnes
        If Index = 0 Then   ' ligne
            Call MettreModeLigCol("L")
        Else
            Call MettreModeLigCol("C")
        End If
    End If
End Sub

Private Function MettreModeLigCol(ByVal v_mode As String)
    
    If v_mode = "L" Then   ' ligne
        Set Me.Image1.Picture = Me.imglst.ListImages(2).Picture
    Else
        Set Me.Image1.Picture = Me.imglst.ListImages(1).Picture
    End If
    
End Function
Private Sub OptNbOccur_Click(Index As Integer)
    '
    Dim i As Integer
    
    
    For i = 0 To Me.OptNbOccur.Count - 1
        If Me.OptNbOccur(i).Value Then
            Me.OptNbOccur(i).ForeColor = COLOR_SEL
            Me.OptNbOccur(i).FontBold = True
        Else
            Me.OptNbOccur(i).ForeColor = &H80000012
            Me.OptNbOccur(i).FontBold = False
        End If
    Next i
    
    If Me.TxtNbOccur.Text = "" Then
        OptNbOccur(Opt_nboccur_Toutes).Value = True
        Index = Opt_nboccur_Toutes
    End If
    If Index = Opt_nboccur_Toutes Then   ' nombre d'occurences
        Me.LblNbreOccur.Visible = False
        Me.TxtNbOccur.Visible = False
        Me.TxtNbOccur.Text = 0
    Else
        Me.FrmNbOccur.Visible = True
        Me.LblNbreOccur.Visible = True
        Me.TxtNbOccur.Visible = True
        If Me.TxtNbOccur.Text = 0 Then Me.TxtNbOccur.Text = 5
        Me.TxtNbOccur.SetFocus
    End If

    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub OptRapportVal_Click(Index As Integer)
    Dim i As Integer
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
    
    For i = 0 To Me.OptRapportVal.Count - 1
        If Me.OptRapportVal(i).Value Then
            Me.OptRapportVal(i).ForeColor = COLOR_SEL
            Me.OptRapportVal(i).FontBold = True
        Else
            Me.OptRapportVal(i).ForeColor = &H80000012
            Me.OptRapportVal(i).FontBold = False
        End If
    Next i
    
End Sub

Private Sub OptTypVal_Click(Index As Integer)
    
    Dim i As Integer
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
    
    For i = 0 To Me.OptTypVal.Count - 1
        If Me.OptTypVal(i).Value Then
            Me.OptTypVal(i).ForeColor = COLOR_SEL
            Me.OptTypVal(i).FontBold = True
        Else
            Me.OptTypVal(i).ForeColor = &H80000012
            Me.OptTypVal(i).FontBold = False
        End If
    Next i
    
    If g_mode_saisie Then
        g_mode_saisie = False
        Call MenForme("OptTypVal", "")
        g_mode_saisie = True
    End If
End Sub

Private Sub OptVal_Click(Index As Integer)
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub OptValForme_Click(Index As Integer)
    Dim i As Integer
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If
    
    For i = 0 To Me.OptValForme.Count - 1
        If Me.OptValForme(i).Value Then
            Me.OptValForme(i).ForeColor = COLOR_SEL
            Me.OptValForme(i).FontBold = True
        Else
            Me.OptValForme(i).ForeColor = &H80000012
            Me.OptValForme(i).FontBold = False
        End If
    Next i
    
    If g_mode_saisie Then
        If Index = Opt_ValForme_Pourcent Or Index = Opt_ValForme_Moyenne Then
            ' par rapport à quoi ?
            Me.FrmParRapport.Visible = True
        Else
            Me.FrmParRapport.Visible = False
        End If
        g_mode_saisie = False
        Call MenForme("OptTypVal", "")
        g_mode_saisie = True
    End If
End Sub

Private Sub TxtNbOccur_Change()
    
    If g_mode_saisie Then
        If VerifSiChange() Then
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub TxtSQL_Change()
    cmd(CMD_OK).Enabled = True
    cmd(CMD_OK).Visible = True
    If TxtSQL.Text <> "" Then
        Me.Frame.Caption = "Champ : Requête SQL"
        Me.Image1.Visible = False
        Me.FrmLibVal.Visible = False
        Me.FrmNbOccur.Visible = False
        Me.FrmParRapport.Visible = False
        Me.FrmFormat.Visible = False
        Me.cmd(CMD_CHOIX_VAL).Visible = False
        Me.OptLigCol(0).Visible = False
        Me.OptLigCol(1).Visible = False
    Else
        Me.Frame.Caption = "Champ : nombre de fiches"
        Me.Image1.Visible = True
        Me.FrmLibVal.Visible = True
        Me.FrmNbOccur.Visible = True
        Me.FrmParRapport.Visible = True
        Me.FrmFormat.Visible = True
        Me.cmd(CMD_CHOIX_VAL).Visible = True
        Me.OptLigCol(0).Visible = True
        Me.OptLigCol(1).Visible = True
    End If
End Sub
