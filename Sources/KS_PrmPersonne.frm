VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form KS_PrmPersonne 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8175
   ClientLeft      =   1125
   ClientTop       =   1455
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8175
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmValid 
      Caption         =   "Validation en cours"
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
      Height          =   1665
      Left            =   960
      TabIndex        =   53
      Top             =   7890
      Visible         =   0   'False
      Width           =   8445
      Begin VB.Label lblValid 
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
         Height          =   405
         Left            =   390
         TabIndex        =   54
         Top             =   660
         Width           =   7785
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "    Personne"
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
      Height          =   7530
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10905
      Begin TabDlg.SSTab sst 
         Height          =   6915
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   12197
         _Version        =   393216
         Tabs            =   5
         Tab             =   3
         TabsPerRow      =   5
         TabHeight       =   520
         ForeColor       =   8388736
         TabCaption(0)   =   "&Général"
         TabPicture(0)   =   "KS_PrmPersonne.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lbl(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl(7)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lbl(8)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lbl(5)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl(16)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lbl(17)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl(18)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chk(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt(3)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txt(4)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "frmGenPlan"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txt(1)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "chk(2)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txt(6)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmd(6)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "chk(3)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txt(14)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "&Postes"
         TabPicture(1)   =   "KS_PrmPersonne.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbl(4)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "ImageListS"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lblLabo"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "grdLabo"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmd(5)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmd(4)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "tvSect"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "&Fonctions autorisées"
         TabPicture(2)   =   "KS_PrmPersonne.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ImageList"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "tvFct"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame2"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "H&oraires"
         TabPicture(3)   =   "KS_PrmPersonne.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "lbl(13)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "lbl(6)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "imglst"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "grdPoste"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "txt(10)"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "cmd(13)"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "cmd(14)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "frmNext"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "cbo(0)"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).ControlCount=   9
         TabCaption(4)   =   "P&articularités"
         TabPicture(4)   =   "KS_PrmPersonne.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chk(1)"
         Tab(4).ControlCount=   1
         Begin ComctlLib.TreeView tvSect 
            Height          =   2655
            Left            =   -73500
            TabIndex        =   15
            Top             =   3390
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   4683
            _Version        =   327682
            Indentation     =   0
            LabelEdit       =   1
            Style           =   1
            ImageList       =   "ImageListS"
            BorderStyle     =   1
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
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   -73380
            MaxLength       =   20
            TabIndex        =   2
            Top             =   1680
            Width           =   3855
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Code générique"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   -71520
            TabIndex        =   7
            Top             =   2640
            Width           =   1725
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
            Height          =   400
            Index           =   6
            Left            =   -74880
            Picture         =   "KS_PrmPersonne.frx":008C
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Envoyer une demande par mail au responsable de KaliBottin"
            Top             =   120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   -66840
            MaxLength       =   20
            TabIndex        =   9
            Top             =   3330
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Externe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   -73200
            TabIndex        =   6
            Top             =   2640
            Width           =   1005
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Doit accuser réception des documents dont il est destinataire"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   -74730
            TabIndex        =   23
            Top             =   810
            Width           =   5955
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   -68580
            MaxLength       =   50
            TabIndex        =   1
            Top             =   1140
            Width           =   3975
         End
         Begin VB.ComboBox cbo 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   0
            Left            =   2850
            TabIndex        =   18
            Top             =   1140
            Width           =   1755
         End
         Begin VB.Frame Frame2 
            Height          =   5835
            Left            =   -67230
            TabIndex        =   50
            Top             =   810
            Width           =   1725
            Begin VB.CommandButton cmd 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Recopier"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   330
               Style           =   1  'Graphical
               TabIndex        =   52
               TabStop         =   0   'False
               ToolTipText     =   "Recopier les fonctions autorisées d'une autre personne"
               Top             =   2700
               Width           =   975
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Toutes"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   330
               Style           =   1  'Graphical
               TabIndex        =   51
               TabStop         =   0   'False
               Tag             =   "T"
               Top             =   1740
               Width           =   990
            End
         End
         Begin VB.Frame frmGenPlan 
            BorderStyle     =   0  'None
            Height          =   2325
            Left            =   -74940
            TabIndex        =   45
            Top             =   4320
            Width           =   9825
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   2580
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   10
               Top             =   240
               Width           =   3615
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00C0C0C0&
               Caption         =   "..."
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
               Index           =   11
               Left            =   6180
               Picture         =   "KS_PrmPersonne.frx":09CA
               TabIndex        =   56
               TabStop         =   0   'False
               ToolTipText     =   "Lister les catégories"
               Top             =   240
               UseMaskColor    =   -1  'True
               Width           =   300
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   7
               Left            =   5400
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   11
               Top             =   900
               Width           =   2235
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00C0C0C0&
               Caption         =   "..."
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
               Index           =   12
               Left            =   7620
               Picture         =   "KS_PrmPersonne.frx":19DC
               TabIndex        =   46
               TabStop         =   0   'False
               ToolTipText     =   "Lister les types de contrat"
               Top             =   900
               UseMaskColor    =   -1  'True
               Width           =   300
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   8
               Left            =   1830
               MaxLength       =   10
               TabIndex        =   12
               Top             =   1350
               Width           =   1485
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   9
               Left            =   5400
               MaxLength       =   10
               TabIndex        =   13
               Top             =   1350
               Width           =   1485
            End
            Begin VB.Label lbl 
               Caption         =   "Catégorie professionnelle"
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
               Index           =   9
               Left            =   180
               TabIndex        =   57
               Top             =   270
               Width           =   2175
            End
            Begin VB.Label lbl 
               Caption         =   "Type de contrat"
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
               Index           =   10
               Left            =   3960
               TabIndex        =   49
               Top             =   930
               Width           =   1395
            End
            Begin VB.Label lbl 
               Caption         =   "Date d'embauche"
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
               Index           =   11
               Left            =   180
               TabIndex        =   48
               Top             =   1410
               Width           =   1605
            End
            Begin VB.Label lbl 
               Caption         =   "Fin du contrat"
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
               Index           =   12
               Left            =   3990
               TabIndex        =   47
               Top             =   1410
               Width           =   1275
            End
         End
         Begin VB.Frame frmNext 
            Height          =   1575
            Left            =   210
            TabIndex        =   41
            Top             =   5370
            Width           =   6975
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   12
               Left            =   4110
               MaxLength       =   3
               TabIndex        =   21
               Top             =   720
               Width           =   915
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00C0C0C0&
               Caption         =   "..."
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
               Index           =   9
               Left            =   3420
               Picture         =   "KS_PrmPersonne.frx":29EE
               TabIndex        =   43
               TabStop         =   0   'False
               ToolTipText     =   "Lister les postes"
               Top             =   720
               UseMaskColor    =   -1  'True
               Width           =   300
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   13
               Left            =   210
               MaxLength       =   5
               TabIndex        =   22
               Top             =   1080
               Width           =   1755
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   11
               Left            =   210
               MaxLength       =   5
               TabIndex        =   20
               Top             =   720
               Width           =   3225
            End
            Begin VB.Label lbl 
               Caption         =   "Prochaine semaine"
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
               Index           =   15
               Left            =   4110
               TabIndex        =   44
               Top             =   360
               Width           =   1905
            End
            Begin VB.Label lbl 
               Caption         =   "Prochain poste"
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
               Index           =   14
               Left            =   210
               TabIndex        =   42
               Top             =   360
               Width           =   1425
            End
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   14
            Left            =   10380
            Picture         =   "KS_PrmPersonne.frx":3A00
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter un poste"
            Top             =   2070
            UseMaskColor    =   -1  'True
            Width           =   285
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   13
            Left            =   10380
            Picture         =   "KS_PrmPersonne.frx":3EF2
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le poste"
            Top             =   5160
            UseMaskColor    =   -1  'True
            Width           =   285
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
            Index           =   4
            Left            =   -67620
            Picture         =   "KS_PrmPersonne.frx":43E4
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Accéder aux services"
            Top             =   3390
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   5
            Left            =   -67620
            Picture         =   "KS_PrmPersonne.frx":483B
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le service"
            Top             =   5730
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   570
            MaxLength       =   5
            TabIndex        =   17
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   -72360
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3330
            Width           =   3615
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   -74130
            MaxLength       =   15
            TabIndex        =   3
            Top             =   2220
            Width           =   1815
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   -74130
            MaxLength       =   50
            TabIndex        =   0
            Top             =   1140
            Width           =   3975
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   -70200
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   2250
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Actif"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   -74760
            TabIndex        =   5
            Top             =   2640
            Width           =   885
         End
         Begin ComctlLib.TreeView tvFct 
            Height          =   5745
            Left            =   -74670
            TabIndex        =   16
            Top             =   900
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   10134
            _Version        =   327682
            Indentation     =   0
            LabelEdit       =   1
            Style           =   1
            ImageList       =   "ImageList"
            BorderStyle     =   1
            Appearance      =   0
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
         Begin MSFlexGridLib.MSFlexGrid grdLabo 
            Height          =   765
            Left            =   -73500
            TabIndex        =   14
            Top             =   1290
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   1349
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
            GridColor       =   4194304
            GridColorFixed  =   4194304
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   2
            ScrollBars      =   2
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
         Begin MSFlexGridLib.MSFlexGrid grdPoste 
            Height          =   3435
            Left            =   180
            TabIndex        =   19
            Top             =   1920
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   6059
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   8388608
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   0
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   2
            ScrollBars      =   2
            Appearance      =   0
         End
         Begin VB.Label lbl 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   18
            Left            =   -71680
            TabIndex        =   64
            Top             =   2235
            Width           =   135
         End
         Begin VB.Label lbl 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   17
            Left            =   -74860
            TabIndex        =   63
            Top             =   2240
            Width           =   135
         End
         Begin VB.Label lbl 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   16
            Left            =   -74860
            TabIndex        =   62
            Top             =   1150
            Width           =   135
         End
         Begin VB.Label lbl 
            Caption         =   "Civilité / Titre"
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
            Index           =   5
            Left            =   -74790
            TabIndex        =   61
            Top             =   1690
            Width           =   1425
         End
         Begin VB.Label lbl 
            Caption         =   "Matricule"
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
            Index           =   8
            Left            =   -67770
            TabIndex        =   59
            Top             =   3360
            Width           =   885
         End
         Begin VB.Label lbl 
            Caption         =   "Prénom"
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
            Index           =   7
            Left            =   -69480
            TabIndex        =   58
            Top             =   1140
            Width           =   705
         End
         Begin ComctlLib.ImageList imglst 
            Left            =   6090
            Top             =   1170
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   15
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   1
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":4C82
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            Caption         =   "heures par"
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
            Index           =   6
            Left            =   1830
            TabIndex        =   55
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Postes de travail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   40
            Top             =   1620
            Width           =   2235
         End
         Begin VB.Label lblLabo 
            Caption         =   "Sites de travail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -74670
            TabIndex        =   37
            Top             =   1440
            Width           =   1035
         End
         Begin ComctlLib.ImageList ImageListS 
            Left            =   -67560
            Top             =   4230
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   24
            ImageHeight     =   21
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   3
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":4FD4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":5826
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":60F8
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            Caption         =   "Postes"
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
            Left            =   -74640
            TabIndex        =   36
            Top             =   4410
            Width           =   945
         End
         Begin ComctlLib.ImageList ImageList 
            Left            =   -67350
            Top             =   2490
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   36
            ImageHeight     =   20
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   4
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":69CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":70EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":780E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "KS_PrmPersonne.frx":7E58
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            Caption         =   "Adresse de messagerie"
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
            Left            =   -74760
            TabIndex        =   33
            Top             =   3330
            Width           =   2175
         End
         Begin VB.Label lbl 
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
            Index           =   0
            Left            =   -74670
            TabIndex        =   32
            Top             =   2250
            Width           =   555
         End
         Begin VB.Label lbl 
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
            Index           =   1
            Left            =   -74670
            TabIndex        =   31
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lbl 
            Caption         =   "Mot de passe"
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
            Index           =   3
            Left            =   -71490
            TabIndex        =   30
            Top             =   2265
            Width           =   1215
         End
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "KS_PrmPersonne.frx":84A2
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   0
      TabIndex        =   27
      Top             =   7410
      Width           =   10905
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KS_PrmPersonne.frx":8901
         Height          =   510
         Index           =   2
         Left            =   5100
         Picture         =   "KS_PrmPersonne.frx":8E90
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Supprimer la personne"
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
         Index           =   1
         Left            =   9750
         Picture         =   "KS_PrmPersonne.frx":9425
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KS_PrmPersonne.frx":99DE
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
         Left            =   600
         Picture         =   "KS_PrmPersonne.frx":9F3A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
   Begin VB.Menu mnuFct 
      Caption         =   "mnuFct"
      Visible         =   0   'False
      Begin VB.Menu mnuResp 
         Caption         =   "&Responsable du service"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQitter 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "KS_PrmPersonne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Images du treeview des fonctions
Private Const BOULE_VERTE = 1
Private Const BOULE_ROUGE = 2
Private Const CARRE_VERT = 3
Private Const CARRE_ROUGE = 4

' Images TreeView des services
Private Const IMGT_SERVICE = 1
Private Const IMGT_SERVICE_RESP = 3
Private Const IMGT_POSTE = 2

Private Const IMG_COCHE = 1

Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_SUPPRIMER = 2
Private Const CMD_RECOPIE = 3
Private Const CMD_DEM_KB = 6
Private Const CMD_CHOIX_CATEGPROF = 11
Private Const CMD_CHOIX_TYPECONTRAT = 12
Private Const CMD_ACCES_SPM = 4
Private Const CMD_MOINS_SPM = 5
Private Const CMD_BASCULE_AUTOR = 8
Private Const CMD_PLUS_POSTE = 14
Private Const CMD_MOINS_POSTE = 13

Private Const TXT_NOM = 0
Private Const TXT_PRENOM = 1
Private Const TXT_PREFIXE = 14
Private Const TXT_CODE = 2
Private Const TXT_MPASSE = 3
Private Const TXT_ADRNET = 4
Private Const TXT_CATEGPROF = 5
Private Const TXT_MATRICULE = 6
Private Const TXT_TYPECONTRAT = 7
Private Const TXT_DATEDEB_EMBAUCHE = 8
Private Const TXT_DATEFIN_EMBAUCHE = 9
Private Const TXT_NBHEURES = 10

Private Const CHK_ACTIF = 0
Private Const CHK_EXTERNE = 2
Private Const CHK_FICTIF = 3
Private Const CHK_AR = 1

Private Const CBO_BASEHEURES = 0

Private Const GRDL_NUMLABO = 0
Private Const GRDL_ESTLABO = 1
' Colonnes visibles
Private Const GRDL_CODLABO = 2
Private Const GRDL_IMG_ESTLABO = 3
Private Const GRDL_LABOPRINC = 4

Private Const GRDJT_NUMHOR = 0
' Colonnes visibles
Private Const GRDJT_JOUR = 1
Private Const GRDJT_HORAIRE = 2
Private Const GRDJT_PAUSE = 3

Private Const GRDP_NUMPOSTE = 0
Private Const GRDP_NUMLABO = 1
Private Const GRDP_NUMTITRE = 2
Private Const GRDP_ASTREINTE_POSSIBLE = 3
Private Const GRDP_ASTREINTE = 4
Private Const GRDP_GARDE = 5
Private Const GRDP_GARDE_POSSIBLE = 6
Private Const GRDP_NUMCYCLE = 7
Private Const GRDP_POSTRMNEXT_NUMTRM = 8
Private Const GRDP_TOURNANTE = 9
' Colonnes visibles
Private Const GRDP_NOMPOSTE = 10
Private Const GRDP_CODLABO = 11
Private Const GRDP_NOMTITRE = 12
Private Const GRDP_PIC_ASTREINTE = 13
Private Const GRDP_PIC_GARDE = 14
Private Const GRDP_NOMCYCLE = 15
Private Const GRDP_NOM_TRMNEXT_TRM = 16
Private Const GRDP_PIC_TOURNANTE = 17
Private Const GRDP_NBSEM_TOURNANTE = 18

Private Type SDOCDIFF
    numdoc As Long
    numvers As Long
    libvers As String
End Type

Private Type SPOSTECORR
    numposte_aremp As Long
    numposte_remp As Long
End Type

Private g_numutil As Long
Private g_sprm As String
Private g_numfct As Long
Private g_spm As String

Private g_crutil_autor As Boolean
Private g_modutil_autor As Boolean
Private g_modopt_autor As Boolean
Private g_crfct_autor As Boolean
Private g_crcateg_autor As Boolean
Private g_crcontrat_autor As Boolean
Private g_crcycle_autor As Boolean
Private g_crtrm_autor As Boolean

Private g_tbl_fctautor1() As Long
Private g_tbl_fctautor2() As Long

Private g_mode_saisie As Boolean
Private g_txt_avant As String
Private g_cbo_avant As Integer

Private g_button As Integer
Private g_form_active As Boolean
Private g_form_width As Integer, g_form_height As Integer

Public Sub AppelFrm(ByVal v_numutil As Long, _
                    ByVal v_sprm As String)

    g_numutil = v_numutil
    g_sprm = v_sprm
    
    Me.Show 1
    
End Sub

Private Sub activer_fctautor_fils(ByVal v_nddeb As Node)

    Dim s As String
    Dim i As Integer
    Dim ndf As Node, ndp As Node, ndn As Node, nd As Node
    
    If v_nddeb = tvFct.Nodes(1) Then
        For i = 2 To tvFct.Nodes.Count
            Set nd = tvFct.Nodes(i)
            If left$(nd.key, 1) = "M" Then
                nd.image = BOULE_VERTE
                nd.SelectedImage = BOULE_VERTE
            Else
                nd.image = CARRE_VERT
                nd.SelectedImage = CARRE_VERT
            End If
        Next i
        Exit Sub
    End If
    
    Set nd = v_nddeb
    If left$(nd.key, 1) = "M" Then
        nd.image = BOULE_VERTE
        nd.SelectedImage = BOULE_VERTE
    Else
        nd.image = CARRE_VERT
        nd.SelectedImage = CARRE_VERT
    End If
    
lab_deb:
    If nd.Children = 0 Then
        Set ndp = nd.Next
        On Error GoTo lab1
        s = ndp.key
        On Error GoTo 0
        If s = "" Then
            Set ndp = nd.Parent
            If ndp = tvFct.SelectedItem Then Exit Sub
            Set ndp = ndp.Next
            On Error GoTo lab_fin
            s = ndp.tag
            On Error GoTo 0
        End If
        Set ndn = ndp
    Else
        Set ndn = nd.Child
    End If
    If left$(ndn.key, 1) = "M" Then
        ndn.image = BOULE_VERTE
        ndn.SelectedImage = BOULE_VERTE
    End If
    Set nd = ndn
    GoTo lab_deb
    
lab_fin:
    Exit Sub

lab1:
    s = ""
    Resume Next
    
End Sub

Private Sub activer_fctautor_peres(ByVal v_nd As Node)

    Dim ndp As Node
    
    If v_nd = tvFct.Nodes(1) Then Exit Sub
    
    Set ndp = v_nd
    Do
        Set ndp = ndp.Parent
        If ndp = tvFct.Nodes(1) Then Exit Sub
        ndp.image = BOULE_VERTE
        ndp.SelectedImage = BOULE_VERTE
    Loop
   
End Sub

Private Function afficher_fct_autor() As Integer


End Function

Private Sub afficher_frm_valid()

    frmValid.Visible = True
    frmValid.ZOrder 0
    Me.Height = frmValid.Height
    Me.Width = frmValid.Width
    frmValid.Top = 0
    frmValid.left = 0
    Call FRM_CentrerForm(Me)
    Me.Refresh
    DoEvents
    
End Sub

Private Sub afficher_laboratoires(ByVal v_ssite As String, _
                                  ByVal v_lnumprinc As Long)

    Dim n As Integer, i As Integer, j As Integer
    Dim numl As Long
    
    n = STR_GetNbchamp(v_ssite, ";")
    For i = 0 To n - 1
        numl = Mid$(STR_GetChamp(v_ssite, ";", i), 2)
        For j = 0 To grdLabo.Rows - 1
            If grdLabo.TextMatrix(j, GRDL_NUMLABO) = numl Then
                grdLabo.TextMatrix(j, GRDL_ESTLABO) = True
                grdLabo.row = j
                grdLabo.col = GRDL_IMG_ESTLABO
                Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
                Exit For
            End If
        Next j
    Next i
    For i = 0 To grdLabo.Rows - 1
        If grdLabo.TextMatrix(i, GRDL_NUMLABO) = v_lnumprinc Then
            grdLabo.TextMatrix(i, GRDL_LABOPRINC) = "Princiapl"
            Exit For
        End If
    Next i
    
    If grdLabo.Rows > 0 Then
        grdLabo.row = 0
        grdLabo.RowSel = 0
    End If
    
End Sub

Private Sub afficher_menu()

    If tvSect.SelectedItem.tag = "1" Then
        mnuResp.Caption = "N'est pas responsable de ce service"
    Else
        mnuResp.Caption = "Est responsable de ce service"
    End If
    Call PopupMenu(mnuFct)
    
End Sub

Private Sub afficher_page(ByVal v_sens As Integer)

    If v_sens = 0 Then
        If sst.Tab > 0 Then
            If sst.Tab = 4 Then
                sst.Tab = 2
            Else
                sst.Tab = sst.Tab - 1
            End If
        Else
            sst.Tab = sst.Tabs - 1
        End If
    Else
        If sst.Tab < sst.Tabs - 1 Then
            If sst.Tab = 2 Then
                sst.Tab = 4
            Else
                sst.Tab = sst.Tab + 1
            End If
        Else
            sst.Tab = 0
        End If
    End If
    
    Call init_focus
    
End Sub

Private Function afficher_postetrav() As Integer

    Dim sql As String, nom As String, slst As String
    Dim bgarde As Boolean, bastreinte As Boolean
    Dim i As Integer, nbtournante As Integer
    Dim numtr As Long
    Dim rs As rdoResultset
    
    nbtournante = 0
    
    sql = "select UPOT_POTNum, UPOT_LNum, UPOT_TPONum, UPOT_GestAstreinte, UPOT_GestGarde" _
        & ", UPOT_TRMOrdreNext, UPOT_Tournante, UPOT_NbSemTournante, UPOT_CYTNum" _
        & ", UPOT_TRMNum" _
        & ", POT_Nom" _
        & ", L_Code" _
        & " from UtilPosteTrav, PosteTrav, Laboratoire" _
        & " where UPOT_UNum=" & g_numutil _
        & " and POT_Num=UPOT_POTNum" _
        & " and L_Num=UPOT_LNum" _
        & " order by UPOT_Ordre"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_postetrav = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        grdPoste.AddItem ""
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMPOSTE) = rs("UPOT_POTNum").Value
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMLABO) = rs("UPOT_LNum").Value
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMTITRE) = rs("UPOT_TPONum").Value
        If rs("UPOT_TPONum").Value > 0 Then
            If Odbc_RecupVal("select TPO_Nom from TitrePoste where TPO_Num=" & rs("UPOT_TPONum").Value, nom) = P_ERREUR Then
                afficher_postetrav = P_ERREUR
                Exit Function
            End If
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOMTITRE) = nom
        End If
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE_POSSIBLE) = True
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE) = rs("UPOT_GestAstreinte").Value
        If rs("UPOT_GestAstreinte").Value Then
            grdPoste.row = grdPoste.Rows - 1
            grdPoste.col = GRDP_PIC_ASTREINTE
            Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        End If
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE_POSSIBLE) = True
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE) = rs("UPOT_GestGarde").Value
        If rs("UPOT_GestGarde").Value Then
            grdPoste.row = grdPoste.Rows - 1
            grdPoste.col = GRDP_PIC_GARDE
            Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        End If
        sql = "select POTL_GestAstreinte, POTL_GestGarde" _
            & " from PosteTravLabo" _
            & " where POTL_POTNum=" & rs("UPOT_POTNum").Value _
            & " and POTL_LNum=" & rs("UPOT_LNum").Value
        If Odbc_RecupVal(sql, bastreinte, bgarde) = P_ERREUR Then
            afficher_postetrav = P_ERREUR
            Exit Function
        End If
        If Not bastreinte Then
            grdPoste.row = grdPoste.Rows - 1
            grdPoste.col = GRDP_PIC_ASTREINTE
            grdPoste.CellBackColor = P_GRIS
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE_POSSIBLE) = False
        End If
        If Not bgarde Then
            grdPoste.row = grdPoste.Rows - 1
            grdPoste.col = GRDP_PIC_GARDE
            grdPoste.CellBackColor = P_GRIS
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE_POSSIBLE) = False
        End If
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMCYCLE) = rs("UPOT_CYTNum").Value
        If rs("UPOT_CYTNum").Value > 0 Then
            If Odbc_RecupVal("select CYT_Nom from CycleTrameHebdo where CYT_Num=" & rs("UPOT_CYTNum").Value, nom) = P_ERREUR Then
                afficher_postetrav = P_ERREUR
                Exit Function
            End If
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOMCYCLE) = nom
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_POSTRMNEXT_NUMTRM) = rs("UPOT_TRMOrdreNext").Value
            If rs("UPOT_TRMOrdreNext").Value > 0 Then
                If Odbc_RecupVal("select CYT_LstTrame from CycleTrameHebdo where CYT_Num=" & rs("UPOT_CYTNum").Value, slst) = P_ERREUR Then
                    afficher_postetrav = P_ERREUR
                    Exit Function
                End If
                numtr = CLng(Mid$(STR_GetChamp(slst, ";", rs("UPOT_TRMOrdreNext").Value - 1), 2))
                If Odbc_RecupVal("select TRM_Nom from TrameHebdo where TRM_Num=" & numtr, nom) = P_ERREUR Then
                    afficher_postetrav = P_ERREUR
                    Exit Function
                End If
                grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOM_TRMNEXT_TRM) = rs("UPOT_TRMOrdreNext").Value & " - " & nom
            End If
        Else
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOMCYCLE) = "Horaires fixes"
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_POSTRMNEXT_NUMTRM) = rs("UPOT_TRMNum").Value
            If Odbc_RecupVal("select TRM_Nom from TrameHebdo where TRM_Num=" & rs("UPOT_TRMNum").Value, _
                             nom) = P_ERREUR Then
                afficher_postetrav = P_ERREUR
                Exit Function
            End If
            grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOM_TRMNEXT_TRM) = nom
        End If
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_TOURNANTE) = rs("UPOT_Tournante").Value
        If rs("UPOT_Tournante").Value Then
            grdPoste.row = grdPoste.Rows - 1
            grdPoste.col = GRDP_PIC_TOURNANTE
            Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
            nbtournante = nbtournante + 1
        End If
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOMPOSTE) = rs("POT_Nom").Value
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_CODLABO) = rs("L_Code").Value
        rs.MoveNext
    Wend
    rs.Close
    
    If nbtournante > 1 Then
        frmNext.Visible = True
    Else
        frmNext.Visible = False
    End If
    
End Function

Private Function afficher_services(ByVal v_spm As Variant) As Integer

    Dim s As String, s1 As String, lib As String, sql As String
    Dim n As Integer, i As Integer, j As Integer, n2 As Integer
    Dim num As Long
    Dim nd As Node
    
    n = STR_GetNbchamp(v_spm, "|")
    For i = 1 To n
        s = STR_GetChamp(v_spm, "|", i - 1)
        n2 = STR_GetNbchamp(s, ";")
        For j = 1 To n2
            s1 = STR_GetChamp(s, ";", j - 1)
            If TV_NodeExiste(tvSect, s1, nd) = P_OUI Then
                GoTo lab_sp_suiv
            End If
            num = CLng(Mid$(s1, 2))
            If left(s1, 1) = "S" Then
                If P_RecupSrvNom(num, lib) = P_ERREUR Then
                    afficher_services = P_ERREUR
                    Exit Function
                End If
                If j = 1 Then
                    Set nd = tvSect.Nodes.Add(, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                Else
                    Set nd = tvSect.Nodes.Add(nd, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                End If
                nd.Expanded = True
            Else
                If P_RecupPosteNom(num, lib) = P_ERREUR Then
                    afficher_services = P_ERREUR
                    Exit Function
                End If
                Call tvSect.Nodes.Add(nd, tvwChild, "P" & num, lib, IMGT_POSTE, IMGT_POSTE)
            End If
lab_sp_suiv:
        Next j
    Next i
        
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_MOINS_SPM).Visible = False
    End If
    
    afficher_services = P_OK
    
End Function

Private Function afficher_utilisateur() As Integer

    Dim sql As String, s As String
    Dim pos As Integer
    Dim lnb As Long
    Dim rs As rdoResultset
    
    g_mode_saisie = False
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    sst.TabVisible(3) = False
    frmGenPlan.Visible = False
    sst.TabVisible(4) = False
    chk(CHK_AR).Visible = False
    
    grdLabo.Rows = 0
    Call init_grD_Site
    
    tvSect.Nodes.Clear
    
    grdPoste.tag = ""
    grdPoste.Rows = grdPoste.FixedRows
    
    cmd(CMD_BASCULE_AUTOR).tag = "T"
    cmd(CMD_BASCULE_AUTOR).ToolTipText = "Autoriser toutes les fonctions"
    
    If g_numutil > 0 Then
        sql = "select * from Utilisateur" _
            & " where U_Num=" & g_numutil
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        ' Onglet Général
        txt(TXT_NOM).Text = rs("U_Nom").Value
        txt(TXT_PRENOM).Text = rs("U_Prenom").Value
        txt(TXT_PREFIXE).Text = rs("U_Prefixe").Value
        chk(CHK_ACTIF).Value = IIf(rs("U_Actif").Value, 1, 0)
        chk(CHK_EXTERNE).Value = IIf(rs("U_Externe").Value, 1, 0)
        chk(CHK_FICTIF).Value = IIf(rs("U_Fictif").Value, 1, 0)
        If chk(CHK_FICTIF).Value = 1 Then
            chk(CHK_AR).Value = 0
            chk(CHK_AR).Enabled = False
        Else
            chk(CHK_AR).Value = 1
            chk(CHK_AR).Enabled = True
        End If
        txt(TXT_CATEGPROF).tag = rs("U_CATPNum").Value
        If rs("U_CATPNum").Value > 0 Then
            sql = "select CATP_Nom from CategorieProf" _
                & " where CATP_Num=" & rs("U_CATPNum").Value
            If Odbc_RecupVal(sql, s) = P_ERREUR Then
                GoTo lab_err
            End If
            txt(TXT_CATEGPROF).Text = s
        Else
            txt(TXT_CATEGPROF).Text = ""
        End If
        txt(TXT_MATRICULE).Text = rs("U_Matricule").Value & ""
        txt(TXT_TYPECONTRAT).tag = rs("U_CTRAVNum").Value
        If rs("U_CTRAVNum").Value > 0 Then
            sql = "select CTRAV_Nom from ContratTravail" _
                & " where CTRAV_Num=" & rs("U_CTRAVNum").Value
            If Odbc_RecupVal(sql, s) = P_ERREUR Then
                GoTo lab_err
            End If
            txt(TXT_TYPECONTRAT).Text = s
        Else
            txt(TXT_TYPECONTRAT).Text = ""
        End If
        txt(TXT_DATEDEB_EMBAUCHE).Text = IIf(IsNull(rs("U_DateDebEmbauche").Value), "", Format(rs("U_DateDebEmbauche").Value, "dd/mm/yyyy"))
        txt(TXT_DATEFIN_EMBAUCHE).Text = IIf(IsNull(rs("U_DateFinEmbauche").Value), "", Format(rs("U_DateFinEmbauche").Value, "dd/mm/yyyy"))
        ' Onglet Poste
        Call afficher_laboratoires(rs("U_Labo").Value, rs("U_LNumPrinc").Value)
        If afficher_services(rs("U_SPM").Value & "") = P_ERREUR Then
            GoTo lab_err
        End If
        ' Onglet Particularités
        chk(CHK_AR).Value = IIf(rs("U_AR").Value, 1, 0)
        rs.Close
        ' Code et Mot de passe
        sql = "select UAPP_Code, UAPP_MotPasse from UtilAppli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_UNum=" & g_numutil
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        txt(TXT_CODE).Text = rs("UAPP_Code").Value
        txt(TXT_MPASSE).Text = STR_Decrypter(rs("UAPP_MotPasse").Value)
        ' Adr mail
        sql = "select UC_Valeur from ZoneUtil, UtilCoordonnee" _
            & " where ZU_Code='ADRMAIL'" _
            & " and UC_ZUNum=ZU_Num" _
            & " and UC_Type='U'" _
            & " and UC_TypeNum=" & g_numutil _
            & " and UC_Principal=true"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        If rs.EOF Then
            txt(TXT_ADRNET).Text = ""
        Else
            txt(TXT_ADRNET).Text = rs("UC_Valeur").Value & ""
        End If
        rs.Close
        txt(TXT_ADRNET).tag = txt(TXT_ADRNET).Text
        cmd(CMD_OK).Enabled = False
        cmd(CMD_SUPPRIMER).Enabled = True
    Else
        txt(TXT_NOM).Text = ""
        txt(TXT_PRENOM).Text = ""
        txt(TXT_PREFIXE).Text = ""
        txt(TXT_CODE).Text = ""
        txt(TXT_MPASSE).Text = ""
        chk(CHK_ACTIF).Value = 1
        chk(CHK_EXTERNE).Value = 0
        txt(TXT_ADRNET).Text = ""
        txt(TXT_CATEGPROF).tag = 0
        txt(TXT_CATEGPROF).Text = ""
        txt(TXT_MATRICULE).Text = ""
        txt(TXT_TYPECONTRAT).tag = 0
        txt(TXT_TYPECONTRAT).Text = ""
        txt(TXT_DATEDEB_EMBAUCHE).Text = ""
        txt(TXT_DATEFIN_EMBAUCHE).Text = ""
        cmd(CMD_MOINS_SPM).Visible = False
        chk(CHK_AR).Value = 1
        cmd(CMD_OK).Enabled = True
        cmd(CMD_SUPPRIMER).Enabled = False
        pos = InStr(g_sprm, "NOM=")
        If pos > 0 Then
            txt(TXT_NOM).Text = Mid$(g_sprm, pos + 4)
        End If
        pos = InStr(g_sprm, "POSTE=")
        If pos > 0 Then
            If build_services(Mid$(g_sprm, pos + 6), s) = P_ERREUR Then
                GoTo lab_err
            End If
            If afficher_services(s) = P_ERREUR Then
                GoTo lab_err
            End If
        End If
    End If
    
    ' Adr mail modifiable ?
    sql = "select count(*) from ZoneUtil, ZoneUtilAppli" _
        & " where ZU_Code='ADRMAIL'" _
        & " and ZUA_ZUNum=ZU_Num" _
        & " and ZUA_APPNum=" & p_appli_kalidoc
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        GoTo lab_err
    End If
    If lnb > 0 Then
        txt(TXT_ADRNET).Enabled = True
    Else
        txt(TXT_ADRNET).Enabled = False
    End If
    
    ' Onglet Fct autorisées
    If afficher_fct_autor() = P_ERREUR Then
        GoTo lab_err
    End If

    g_mode_saisie = True
    
    sst.Tab = 0
    txt(TXT_NOM).SetFocus
    
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    
    afficher_utilisateur = P_OK
    Exit Function
    
lab_err:
    afficher_utilisateur = P_ERREUR

End Function

' Ajouter dans la table des mouvements
Private Sub ajouter_mouvement(ByVal v_numutil As Long, _
                              ByVal v_stype As String, _
                              ByVal v_comm As Variant)

End Sub

Private Sub ajouter_postetrav()

    Dim sql As String, s As String
    Dim y_est As Boolean
    Dim n As Integer, ilig As Integer, ind As Integer
    Dim numposte As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    Call CL_InitTitreHelp("Postes de travail", "")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    n = 0
    sql = "select POT_Num, POT_Nom" _
        & ", POTL_LNum" _
        & ", L_Code" _
        & " from PosteTrav, PosteTravLabo, Laboratoire" _
        & " where POTL_POTNum=POT_Num" _
        & " and L_Num=POTL_LNum" _
        & " order by POT_Nom, L_Code"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    numposte = -1
    While Not rs.EOF
        If rs("POT_Num").Value <> numposte Then
            y_est = False
        End If
        For ilig = 0 To grdLabo.Rows - 1
            If grdLabo.TextMatrix(ilig, GRDL_NUMLABO) = rs("POTL_LNum").Value Then
                If Not y_est Then
                    y_est = True
                    s = rs("POT_Nom").Value & vbTab & rs("L_Code").Value
                Else
                    s = "" & vbTab & rs("L_Code").Value
                End If
                Call CL_AddLigne(s, _
                                  rs("POT_Num").Value, _
                                  rs("POTL_LNum").Value, _
                                  False)
                n = n + 1
            End If
        Next ilig
lab_pot_suiv:
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Call MsgBox("Aucun poste de travail n'est disponible.", vbExclamation + vbOKOnly, "")
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    ind = CL_liste.pointeur
    grdPoste.AddItem ""
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMPOSTE) = CL_liste.lignes(ind).num
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMLABO) = CL_liste.lignes(ind).tag
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMTITRE) = 0
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE_POSSIBLE) = True
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE) = False
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE_POSSIBLE) = True
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE) = False
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NUMCYCLE) = -1
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_POSTRMNEXT_NUMTRM) = -1
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_TOURNANTE) = False
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_NOMPOSTE) = STR_GetChamp(CL_liste.lignes(ind).texte, vbTab, 0)
    grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_CODLABO) = STR_GetChamp(CL_liste.lignes(ind).texte, vbTab, 1)
    
    sql = "select POTL_GestAstreinte, POTL_GestGarde" _
        & " from PosteTravLabo" _
        & " where POTL_POTNum=" & CL_liste.lignes(ind).num _
        & " and POTL_LNum=" & CL_liste.lignes(ind).tag
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs("POTL_GestAstreinte").Value Then
        grdPoste.row = grdPoste.Rows - 1
        grdPoste.col = GRDP_PIC_ASTREINTE
        grdPoste.CellBackColor = P_GRIS
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_ASTREINTE_POSSIBLE) = False
    End If
    If Not rs("POTL_GestGarde").Value Then
        grdPoste.row = grdPoste.Rows - 1
        grdPoste.col = GRDP_PIC_GARDE
        grdPoste.CellBackColor = P_GRIS
        grdPoste.TextMatrix(grdPoste.Rows - 1, GRDP_GARDE_POSSIBLE) = False
    End If
    rs.Close
    
    grdPoste.tag = "M"
    cmd(CMD_OK).Enabled = True
    cmd(CMD_MOINS_POSTE).Visible = True
    grdPoste.SetFocus
    
End Sub

Private Sub basculer_etat_resp()

    Dim nd As Node
    
    Set nd = tvSect.SelectedItem
    If nd.tag = "1" Then
        nd.tag = "0"
        nd.image = IMGT_SERVICE
        nd.SelectedImage = IMGT_SERVICE
    Else
        nd.tag = "1"
        nd.image = IMGT_SERVICE_RESP
        nd.SelectedImage = IMGT_SERVICE_RESP
    End If
    
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Function build_services(ByVal v_numposte As Long, _
                               ByRef r_sp As String) As Integer
                               
    Dim s As String, sql As String
    Dim numsrv As Long
    
    s = "P" & v_numposte & ";"
    sql = "select PO_SRVNum from Poste where PO_Num=" & v_numposte
    If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
        build_services = P_ERREUR
        Exit Function
    End If
    s = "S" & numsrv & ";" & s
    Do
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            build_services = P_ERREUR
            Exit Function
        End If
        If numsrv > 0 Then
            s = "S" & numsrv & ";" & s
        End If
    Loop Until numsrv = 0
    
    r_sp = s
    
    build_services = P_OK
    
End Function

Private Sub build_SPM_Fct(ByRef r_spm As Variant, _
                          ByRef r_sfct As String)

    Dim s As String, sp As String, sql As String
    Dim encore As Boolean
    Dim i As Integer, j As Integer, n As Integer
    Dim numfct As Long, num As Long
    Dim nd As Node, ndp As Node
    
    r_spm = ""
    r_sfct = ""
    
    For i = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(i)
        If left$(nd.key, 1) = "P" Then
            num = Mid$(nd.key, 2)
            sql = "select PO_FTNum from Poste" _
                & " where PO_Num=" & num
            Call Odbc_RecupVal(sql, numfct)
            If InStr(r_sfct, "F" & numfct & ";") = 0 Then
                r_sfct = r_sfct & "F" & numfct & ";"
            End If
            sp = nd.key & ";"
            Do
                Set ndp = nd.Parent
                encore = True
                On Error GoTo lab_no_prev
                s = ndp.key
                On Error GoTo 0
                If encore Then
                    sp = sp & ndp.key & ";"
                    Set nd = ndp
                End If
            Loop Until Not encore
            n = STR_GetNbchamp(sp, ";")
            For j = n - 1 To 0 Step -1
                r_spm = r_spm + STR_GetChamp(sp, ";", j) & ";"
            Next j
            r_spm = r_spm + "|"
        End If
    Next i
    
    If r_spm <> "" Then
        r_spm = IIf(Right$(r_spm, 1) = "|", r_spm, r_spm + "|")
    End If
    
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Sub basculer_autor()

    Dim i As Integer
    Dim img_boule As Long, img_carre As Long, img As Long
    Dim nd As Node
    
    If cmd(CMD_BASCULE_AUTOR).tag = "T" Then
        img_boule = BOULE_VERTE
        img_carre = CARRE_VERT
        cmd(CMD_BASCULE_AUTOR).tag = "A"
        cmd(CMD_BASCULE_AUTOR).Caption = "&Aucune"
        cmd(CMD_BASCULE_AUTOR).ToolTipText = "Interdire toutes les fonctions"
    Else
        img_boule = BOULE_ROUGE
        img_carre = CARRE_ROUGE
        cmd(CMD_BASCULE_AUTOR).tag = "T"
        cmd(CMD_BASCULE_AUTOR).Caption = "&Toutes"
        cmd(CMD_BASCULE_AUTOR).ToolTipText = "Autoriser toutes les fonctions"
    End If
    
    For i = 2 To tvFct.Nodes.Count
        Set nd = tvFct.Nodes(i)
        If left$(nd.key, 1) = "M" Then
            img = img_boule
        Else
            img = img_carre
        End If
        nd.image = img
        nd.SelectedImage = img
    Next i
    
    tvFct.tag = True
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_categprof()

    Dim libc As String, sql As String
    Dim n As Integer
    Dim numc As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    Call CL_InitTitreHelp("Catégories professionnelles", "")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    If g_crcateg_autor Then
        Call CL_AddLigne("<Nouvelle>", 0, "", False)
        n = 1
    Else
        n = 0
    End If
    
    sql = "select * from CategorieProf" _
        & " order by CATP_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("CATP_Nom").Value, rs("CATP_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        MsgBox "Aucune catégorie n'est disponible.", vbExclamation + vbOKOnly, ""
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    numc = CL_liste.lignes(CL_liste.pointeur).num
    If numc = 0 Then
        Call creer_categprof(numc, libc)
        If numc = 0 Then Exit Sub
    Else
        libc = CL_liste.lignes(CL_liste.pointeur).texte
    End If
    
    txt(TXT_CATEGPROF).tag = numc
    txt(TXT_CATEGPROF).Text = libc
        
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_contrat()

    Dim libc As String, sql As String
    Dim n As Integer
    Dim numc As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    Call CL_InitTitreHelp("Contrats de travail", "")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    If g_crcontrat_autor Then
        Call CL_AddLigne("<Nouveau>", 0, "", False)
        n = 1
    Else
        n = 0
    End If
    
    sql = "select * from ContratTravail" _
        & " order by CTRAV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("CTRAV_Nom").Value, rs("CTRAV_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        MsgBox "Aucun contrat n'est disponible.", vbExclamation + vbOKOnly, ""
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    numc = CL_liste.lignes(CL_liste.pointeur).num
    If numc = 0 Then
        Call creer_contrat(numc, libc)
        If numc = 0 Then Exit Sub
    Else
        libc = CL_liste.lignes(CL_liste.pointeur).num
    End If
    
    txt(TXT_TYPECONTRAT).tag = numc
    txt(TXT_TYPECONTRAT).Text = libc
        
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_cycletrame(ByVal v_lig As Integer)

    Dim sql As String, libc As String
    Dim n As Integer
    Dim numc As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    Call CL_InitTitreHelp("Cycles de trames", "")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    n = 0
    If g_crcycle_autor Then
        Call CL_AddLigne("<Nouveau>", -1, "", False)
        n = n + 1
    End If
    Call CL_AddLigne("Horaires fixes", 0, "", False)
    n = n + 1

    sql = "select * from CycleTrameHebdo" _
        & " order by CYT_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("CYT_Nom").Value, rs("CYT_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        MsgBox "Aucun cycle de trames n'est disponible.", vbExclamation + vbOKOnly, ""
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    numc = CL_liste.lignes(CL_liste.pointeur).num
    If numc = 0 Then
        Call creer_cycletrame(numc, libc)
        If numc = 0 Then Exit Sub
    Else
        libc = CL_liste.lignes(CL_liste.pointeur).texte
    End If
    
    grdPoste.TextMatrix(v_lig, GRDP_NUMCYCLE) = numc
    grdPoste.TextMatrix(v_lig, GRDP_NOMCYCLE) = libc
    grdPoste.tag = "M"
        
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_fonction()

    Dim n As Integer
    Dim rs As rdoResultset
    
    If Odbc_SelectV("select FT_Num, FT_Libelle from FctTrav order by FT_Libelle", rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
        
    Call CL_Init
    Call CL_InitTitreHelp("Choix d'une fonction", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = -1
    If g_numfct > 0 Then
        Call CL_AddLigne("Toutes les fonctions", 0, "", False)
        n = 0
    End If
    While Not rs.EOF
        n = n + 1
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        rs.MoveNext
    Wend
    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If

    g_numfct = CL_liste.lignes(CL_liste.pointeur).num
    
End Sub

Private Sub choisir_next_trame(ByVal v_lig As Integer)

    Dim sql As String, slst As String, nomtr As String
    Dim i As Integer, n As Integer, nbch As Integer
    Dim numtr As Long
    
    If grdPoste.TextMatrix(v_lig, GRDP_NUMCYCLE) <= 0 Then
        Exit Sub
    End If
    
    Call CL_Init
    Call CL_InitTitreHelp("Trames du cycle", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    n = 0
    sql = "select CYT_LstTrame" _
        & " from CycleTrameHebdo" _
        & " where CYT_Num=" & grdPoste.TextMatrix(v_lig, GRDP_NUMCYCLE)
    If Odbc_RecupVal(sql, slst) = P_ERREUR Then
        Exit Sub
    End If
    nbch = STR_GetNbchamp(slst, ";")
    For i = 0 To nbch - 1
        numtr = CLng(Mid$(STR_GetChamp(slst, ";", i), 2))
        sql = "select TRM_Nom" _
            & " from TrameHebdo" _
            & " where TRM_Num=" & numtr
        If Odbc_RecupVal(sql, nomtr) = P_ERREUR Then
            Exit Sub
        End If
        Call CL_AddLigne(n + 1 & vbTab & nomtr, numtr, "", False)
        n = n + 1
    Next i
    If n = 0 Then
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    grdPoste.TextMatrix(v_lig, GRDP_POSTRMNEXT_NUMTRM) = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 0)
    grdPoste.TextMatrix(v_lig, GRDP_NOM_TRMNEXT_TRM) = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 0) & " -" _
                                                       & STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 1)
    grdPoste.tag = "M"
        
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_tramehebdo(ByVal v_lig As Integer)

    Dim sql As String, libtr As String
    Dim n As Integer
    Dim numtr As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    Call CL_InitTitreHelp("Trames hebdomadaires", "")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    n = 0
    If g_crtrm_autor Then
        Call CL_AddLigne("<Nouvelle>", -1, "", False)
        n = n + 1
    End If

    sql = "select * from TrameHebdo" _
        & " order by TRM_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("TRM_Nom").Value, rs("TRM_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        MsgBox "Aucun trame n'est disponible.", vbExclamation + vbOKOnly, ""
        Exit Sub
    End If
    
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    numtr = CL_liste.lignes(CL_liste.pointeur).num
    If numtr = 0 Then
        Call creer_tramehebdo(numtr, libtr)
        If numtr = 0 Then Exit Sub
    Else
        libtr = CL_liste.lignes(CL_liste.pointeur).texte
    End If
    
    grdPoste.TextMatrix(v_lig, GRDP_POSTRMNEXT_NUMTRM) = numtr
    grdPoste.TextMatrix(v_lig, GRDP_NOM_TRMNEXT_TRM) = libtr
    grdPoste.tag = "M"
        
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub choisir_titre()

    Dim sql As String
    Dim n As Integer
    Dim rs As rdoResultset
    
    sql = "select * from TitrePoste" _
        & " where TPO_Num>1" _
        & " order by TPO_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("TPO_Nom").Value, rs("TPO_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    
    Call CL_Init
    Call CL_InitTitreHelp("Titres de poste", "")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    grdPoste.TextMatrix(grdPoste.row, GRDP_NUMTITRE) = CL_liste.lignes(CL_liste.pointeur).num
    grdPoste.TextMatrix(grdPoste.row, GRDP_NOMTITRE) = CL_liste.lignes(CL_liste.pointeur).texte
    grdPoste.tag = "M"

    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub creer_categprof(ByRef r_num As Long, _
                            ByRef r_lib As String)
                            
    MsgBox "Gérer la création d'une catégorie"
    
    r_num = 0
    
End Sub

Private Sub creer_contrat(ByRef r_num As Long, _
                          ByRef r_lib As String)
                            
    MsgBox "Gérer la création d'un contrat"
    
    r_num = 0
    
End Sub

Private Sub creer_cycletrame(ByRef r_num As Long, _
                             ByRef r_lib As String)
                            
    Dim sret As String
    
'    PrmCycle.Tag = "0|0"
'    PrmCycle.Show 1
'    sret = PrmCycle.Tag
'    Unload PrmCycle
    If sret = "" Then
        r_num = -1
        Exit Sub
    End If
    
    r_num = CLng(sret)
    If Odbc_RecupVal("select CYT_Nom from CycleTrameHebdo where CYT_Num=" & r_num, r_lib) = P_ERREUR Then
        r_num = -1
    End If
    
End Sub

Private Sub creer_tramehebdo(ByRef r_num As Long, _
                             ByRef r_lib As String)
                            
    Dim sret As String
    
'    PrmTrame.Tag = "0"
'    PrmTrame.Show 1
'    sret = PrmTrame.Tag
'    Unload PrmTrame
    If sret = "" Then
        r_num = 0
        Exit Sub
    End If
    
    r_num = CLng(sret)
    If Odbc_RecupVal("select TRM_Nom from TrameHebdo where TRM_Num=" & r_num, r_lib) = P_ERREUR Then
        r_num = 0
    End If
    
End Sub

Private Sub envoyer_demande_kb()

End Sub

' v_mode : -1=Suppression   0=Modif   1=Ajout
Private Function gerer_adrmail(ByVal v_numutil As Long, _
                               ByVal v_mode As Integer) As Integer

    Dim sql As String
    Dim num_zone As Long, lnb As Long, lbid As Long
    
    sql = "select ZU_Num from ZoneUtil where ZU_Code='ADRMAIL'"
    If Odbc_RecupVal(sql, num_zone) = P_ERREUR Then
        gerer_adrmail = P_ERREUR
        Exit Function
    End If
    Select Case v_mode
    Case -1
        If Odbc_Delete("UtilCoordonnee", _
                       "UC_Num", _
                        "where UC_ZUNum=" & num_zone _
                            & " and UC_Type='U'" _
                            & " and UC_TypeNum=" & v_numutil, _
                        lnb) = P_ERREUR Then
            gerer_adrmail = P_ERREUR
            Exit Function
        End If
    Case 0
        If Odbc_Update("UtilCoordonnee", _
                        "UC_Num", _
                        "where UC_ZUNum=" & num_zone _
                            & " and UC_Type='U'" _
                            & " and UC_TypeNum=" & v_numutil, _
                        "UC_Valeur", txt(TXT_ADRNET).Text) = P_ERREUR Then
            gerer_adrmail = P_ERREUR
            Exit Function
        End If
    Case 1
        If Odbc_AddNew("UtilCoordonnee", _
                        "UC_Num", _
                        "uc_seq", _
                        False, _
                        lbid, _
                        "UC_ZUNum", num_zone, _
                        "UC_Type", "U", _
                        "UC_TypeNum", v_numutil, _
                        "UC_Valeur", txt(TXT_ADRNET).Text, _
                        "UC_Principal", True, _
                        "UC_Niveau", 0, _
                        "UC_Comm", "") = P_ERREUR Then
            gerer_adrmail = P_ERREUR
            Exit Function
        End If
    End Select
    
    gerer_adrmail = P_OK
    
End Function

Private Function gerer_ajout_dest(ByVal v_numutil As Long, _
                                  ByVal v_ar As Boolean, _
                                  ByVal v_ssite As String, _
                                  ByVal v_sfct As String, _
                                  ByVal v_spm As String) As Integer
                                  
End Function

Public Function gerer_chgt_poste_act(ByVal v_numutil As Long, _
                                     ByVal v_spm As Variant) As Integer

End Function

Public Function gerer_chgt_prmutil(ByVal v_numutil As Long, _
                                    ByVal v_ar As Boolean, _
                                    ByVal v_ossite As String, _
                                    ByVal v_ssite As String, _
                                    ByVal v_osfct As String, _
                                    ByVal v_sfct As String, _
                                    ByVal v_ospm As Variant, _
                                    ByVal v_spm As Variant) As Integer

    Dim s As String, s_ajout As String, s_suppr As String, s_comm As String
    Dim fsupp As Boolean, fsupp_poste As Boolean, fajout As Boolean
    Dim n As Integer, i As Integer
    
    ' On regarde si on a supprimé labo/fct/SPM
    fsupp = False
    fsupp_poste = False
    ' Labo
    n = STR_GetNbchamp(v_ossite, ";")
    For i = 0 To n - 1
        s = STR_GetChamp(v_ossite, ";", i) & ";"
        If InStr(v_ssite, s) = 0 Then
            fsupp = True
            Exit For
        End If
    Next i
    ' Fct
    If Not fsupp Then
        n = STR_GetNbchamp(v_osfct, ";")
        For i = 0 To n - 1
            s = STR_GetChamp(v_osfct, ";", i) & ";"
            If InStr(v_sfct, s) = 0 Then
                fsupp = True
                Exit For
            End If
        Next i
    End If
    ' SPM
    n = STR_GetNbchamp(v_ospm, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(v_ospm, "|", i)
        If InStr(v_spm, s) = 0 Then
            fsupp = True
            fsupp_poste = True
            Exit For
        End If
    Next i
    ' Suppression
    If fsupp Then
        lblValid.Caption = "Suppression de cette personne dans les destinataires des documents concernés"
        Me.Refresh
        If gerer_suppr_dest(v_numutil, _
                            v_ssite, _
                            v_sfct, _
                            v_spm) = P_ERREUR Then
            gerer_chgt_prmutil = P_ERREUR
            Exit Function
        End If
    End If
    ' Suppression de postes : Si acteur dans DoXX alors que le poste n'existe plus -> demande de remplacement
    If fsupp_poste Then
        lblValid.Caption = "Vérification du poste si la personne est paramétrée comme acteur"
        Me.Refresh
        If gerer_chgt_poste_act(v_numutil, v_spm) = P_ERREUR Then
            gerer_chgt_prmutil = P_ERREUR
            Exit Function
        End If
    End If
    
    ' On regarde si on a ajouté labo/fct/SPM
    fajout = False
    ' Labo
    n = STR_GetNbchamp(v_ssite, ";")
    For i = 0 To n - 1
        s = STR_GetChamp(v_ssite, ";", i) & ";"
        If InStr(v_ossite, s) = 0 Then
            fajout = True
            Exit For
        End If
    Next i
    ' Fct
    If Not fajout Then
        n = STR_GetNbchamp(v_sfct, ";")
        For i = 0 To n - 1
            s = STR_GetChamp(v_sfct, ";", i) & ";"
            If InStr(v_osfct, s) = 0 Then
                fajout = True
                Exit For
            End If
        Next i
    End If
    ' SPM
    If Not fajout Then
        n = STR_GetNbchamp(v_spm, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(v_spm, "|", i)
            If InStr(v_ospm, s) = 0 Then
                fajout = True
                Exit For
            End If
        Next i
    End If
    
    If fajout Then
        lblValid.Caption = "Ajout de cette personne dans les destinataires des documents concernés"
        Me.Refresh
        If gerer_ajout_dest(v_numutil, _
                            v_ar, _
                            v_ssite, _
                            v_sfct, _
                            v_spm) = P_ERREUR Then
            gerer_chgt_prmutil = P_ERREUR
            Exit Function
        End If
    End If
    
    gerer_chgt_prmutil = P_OK

End Function

Public Function gerer_diffusion(ByVal v_numdoc As Long, _
                                 ByVal v_numutil As Long, _
                                 ByVal v_ar As Boolean, _
                                 ByVal v_bdiffpapier As Boolean) As Integer
                          
End Function

Public Function gerer_nouvel_utilisateur(ByVal v_numutil As Long) As Integer

End Function

Private Function gerer_suppr_dest(ByVal v_numutil As Long, _
                                 ByVal v_ssite As String, _
                                 ByVal v_sfct As String, _
                                 ByVal v_spm As String) As Integer

End Function

Private Function gerer_suppr_diffusion(ByVal v_numutil As Long, _
                                       ByVal v_ulabo As String) As Integer

End Function

Public Function gerer_suppr_utilisateur(ByVal v_numutil As Long) As Integer

End Function

Private Sub inhiber_fctautor_fils(ByVal v_nddeb As Node)

    Dim i As Integer
    Dim nd As Node, ndp As Node
    
    If v_nddeb = tvFct.Nodes(1) Then
        For i = 2 To tvFct.Nodes.Count
            Set nd = tvFct.Nodes(i)
            If left$(nd.key, 1) = "M" Then
                nd.image = BOULE_ROUGE
                nd.SelectedImage = BOULE_ROUGE
            Else
                nd.image = CARRE_ROUGE
                nd.SelectedImage = CARRE_ROUGE
            End If
        Next i
        Exit Sub
    End If
    
    For i = 2 To tvFct.Nodes.Count
        Set nd = tvFct.Nodes(i)
        Set ndp = nd.Parent
        While ndp.Index <> tvFct.Nodes(1).Index
            If ndp.Index = v_nddeb.Index Then
                If left$(nd.key, 1) = "M" Then
                    nd.image = BOULE_ROUGE
                    nd.SelectedImage = BOULE_ROUGE
                Else
                    nd.image = CARRE_ROUGE
                    nd.SelectedImage = CARRE_ROUGE
                End If
            End If
            Set ndp = ndp.Parent
        Wend
    Next i

End Sub

Private Sub inhiber_frm_valid()

    frmValid.Visible = False
    Me.Height = g_form_height
    Me.Width = g_form_width
    Me.Refresh
    DoEvents
    
End Sub

Private Sub init_focus()

    Select Case sst.Tab
    Case 0
        txt(TXT_CODE).SetFocus
    Case 1
        If grdLabo.Visible Then
            grdLabo.SetFocus
        Else
            tvSect.SetFocus
        End If
    Case 2
        tvFct.SetFocus
    Case 3
        txt(TXT_NBHEURES).SetFocus
    Case 4
        If chk(CHK_AR).Enabled Then
            chk(CHK_AR).SetFocus
        End If
    End Select

End Sub

Private Function init_grD_Site() As Integer

    Dim sql As String
    Dim i As Integer
    Dim rs As rdoResultset
    
    If p_NbLabo > 1 Then
        sql = "select * from Laboratoire" _
            & " order by L_Code"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            init_grD_Site = P_ERREUR
            Exit Function
        End If
        i = 0
        While Not rs.EOF
            grdLabo.AddItem rs("L_Num").Value & vbTab _
                            & False & vbTab _
                            & rs("L_Code").Value
            i = i + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        sql = "select * from Laboratoire"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            init_grD_Site = P_ERREUR
            Exit Function
        End If
        grdLabo.AddItem rs("L_Num").Value & vbTab _
                        & True & vbTab _
                        & rs("L_Code").Value & vbTab _
                        & "" & vbTab _
                        & "Principal"
        lblLabo.Visible = False
        grdLabo.Visible = False
    End If

    init_grD_Site = P_OK
    
End Function

Private Sub initialiser()

    Dim col As Integer
    
    g_numfct = 0
    g_spm = ""
    
    g_crcateg_autor = P_UtilEstAutorFct("CR_CATEG")
    g_crcontrat_autor = P_UtilEstAutorFct("CR_CONTRATTRAV")
    g_crcycle_autor = P_UtilEstAutorFct("CR_CYCLESEM")
    g_crtrm_autor = P_UtilEstAutorFct("CR_TRAME")
    
    grdLabo.Cols = 5
    grdLabo.ColWidth(0) = 0
    grdLabo.ColWidth(1) = 0
    grdLabo.ColWidth(2) = 2800
    grdLabo.ColWidth(3) = 500
    grdLabo.ColWidth(4) = 1000
    
    cmd(CMD_BASCULE_AUTOR).tag = "T"
    cmd(CMD_BASCULE_AUTOR).Caption = "&Toutes"
    cmd(CMD_BASCULE_AUTOR).ToolTipText = "Autoriser toutes les fonctions"
    
    grdPoste.FormatString = "||||||||||Poste|" _
                            & "Site|Titre|A.|G.|Cycle|" _
                            & "Prochaine trame hebdo dans le cycle / Trame si horaires fixes|" _
                            & "Tournante|" _
                            & "Cycle de tournante en semaines"
    grdPoste.Cols = 19
    grdPoste.RowHeight(0) = 1240
    grdPoste.TextMatrix(0, GRDP_NBSEM_TOURNANTE) = "Cycle de tournante" & vbCr & vbLf & "en semaines"

    col = 0
    While col < GRDP_NOMPOSTE
        grdPoste.ColWidth(col) = 0
        col = col + 1
    Wend
    grdPoste.ColAlignment(GRDP_NOMPOSTE) = 1
    grdPoste.ColWidth(GRDP_NOMPOSTE) = 1500
    grdPoste.ColAlignment(GRDP_CODLABO) = 1
    grdPoste.ColWidth(GRDP_CODLABO) = 1000
    grdPoste.ColAlignment(GRDP_NOMTITRE) = 1
    grdPoste.ColWidth(GRDP_NOMTITRE) = 1500
    grdPoste.ColAlignment(GRDP_PIC_ASTREINTE) = 1
    grdPoste.ColAlignment(GRDP_PIC_GARDE) = 1
    grdPoste.ColAlignment(GRDP_NOMCYCLE) = 1
    grdPoste.ColWidth(GRDP_NOMCYCLE) = 1500
    grdPoste.ColAlignment(GRDP_NOM_TRMNEXT_TRM) = 0
    grdPoste.ColWidth(GRDP_NOM_TRMNEXT_TRM) = 1500
    grdPoste.ColAlignment(GRDP_PIC_TOURNANTE) = 1
    grdPoste.ColAlignment(GRDP_NOM_TRMNEXT_TRM) = 1
    grdPoste.ColAlignment(GRDP_NBSEM_TOURNANTE) = 1
    
    cbo(CBO_BASEHEURES).AddItem "semaine"
    cbo(CBO_BASEHEURES).AddItem "mois"
    
    Call maj_droits
    
    If afficher_utilisateur() = P_ERREUR Then
        Unload Me
        Exit Sub
    End If
    
End Sub

Private Sub inverser_autor()

    Dim img_test As Long, img_aff As Long
    Dim nd As Node
    
    Set nd = tvFct.SelectedItem
    If nd = tvFct.Nodes(1) Then
        If nd.Children = 0 Then Exit Sub
        Set nd = nd.Child
    End If
    If left$(nd.key, 1) = "M" Then
        img_test = BOULE_VERTE
        img_aff = BOULE_ROUGE
    Else
        img_test = CARRE_VERT
        img_aff = CARRE_ROUGE
    End If
    If nd.image = img_test Then
        nd.image = img_aff
        nd.SelectedImage = img_aff
        If tvFct.SelectedItem.Children > 0 Then Call inhiber_fctautor_fils(tvFct.SelectedItem)
    Else
        nd.image = img_test
        nd.SelectedImage = img_test
        Call activer_fctautor_peres(tvFct.SelectedItem)
        If tvFct.SelectedItem.Children > 0 Then Call activer_fctautor_fils(tvFct.SelectedItem)
    End If
    cmd(CMD_OK).Enabled = True
    tvFct.tag = True
    
End Sub

Private Sub inverser_etat_labo()

    g_mode_saisie = False
    
    If grdLabo.TextMatrix(grdLabo.row, GRDL_ESTLABO) = True Then
        grdLabo.TextMatrix(grdLabo.row, GRDL_ESTLABO) = False
        grdLabo.col = GRDL_IMG_ESTLABO
        Set grdLabo.CellPicture = CM_LoadPicture("")
    Else
        grdLabo.TextMatrix(grdLabo.row, GRDL_ESTLABO) = True
        grdLabo.col = GRDL_IMG_ESTLABO
        Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
    End If
    cmd(CMD_OK).Enabled = True

    g_mode_saisie = True

End Sub

Private Sub inverser_laboprinc()

    Dim i As Integer
    
    g_mode_saisie = False
    
    If grdLabo.TextMatrix(grdLabo.row, GRDL_LABOPRINC) = "" Then
        For i = 0 To grdLabo.Rows - 1
            grdLabo.TextMatrix(i, GRDL_LABOPRINC) = ""
        Next i
        grdLabo.TextMatrix(grdLabo.row, GRDL_LABOPRINC) = "Principal"
        If grdLabo.TextMatrix(grdLabo.row, GRDL_ESTLABO) = False Then
            grdLabo.TextMatrix(grdLabo.row, GRDL_ESTLABO) = True
            grdLabo.col = GRDL_IMG_ESTLABO
            Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        End If
    Else
        grdLabo.TextMatrix(grdLabo.row, GRDL_LABOPRINC) = ""
    End If
    
    grdLabo.col = GRDL_LABOPRINC
    grdLabo.ColSel = GRDL_LABOPRINC
    cmd(CMD_OK).Enabled = True

    g_mode_saisie = True

End Sub

Private Function labo_concerne(ByVal v_dlabo As String, _
                               ByVal v_ulabo As String) As Boolean
                               
    Dim s As String, s_u As String, s_d As String
    Dim i As Integer, j As Integer, n As Integer, m As Integer
    
    n = STR_GetNbchamp(v_ulabo, ";")
    m = STR_GetNbchamp(v_dlabo, ";")
    For i = 0 To n - 1
        s_u = STR_GetChamp(v_ulabo, ";", i) & ";"
        For j = 0 To m - 1
            s_d = STR_GetChamp(v_dlabo, ";", j) & ";"
            If s_u = s_d Then
                labo_concerne = True
                Exit Function
            End If
        Next j
    Next i
    labo_concerne = False
                               
End Function

Private Sub maj_droits()

    g_crfct_autor = P_UtilEstAutorFct("CR_FCTTRAV")
    g_modopt_autor = P_UtilEstAutorFct("MOD_FCT")
    cmd(CMD_OK).Visible = P_UtilEstAutorFct("MOD_UTIL")
    cmd(CMD_SUPPRIMER).Visible = P_UtilEstAutorFct("SUPP_UTIL")
    cmd(CMD_RECOPIE).Visible = P_UtilEstAutorFct("COPIE_FCT")
    
End Sub

Private Sub prm_service()

    Dim s As String, s1 As String, sql As String
    Dim lib As String, sret As String, ssite As String
    Dim au_moins_un As Boolean
    Dim i As Integer, j As Integer, nbch As Integer, n As Integer
    Dim numlabo As Long, num As Long
    Dim spm As Variant
    Dim nd As Node
    Dim frm As Form
    
    Call CL_Init
    Call build_SPM_Fct(spm, s)
    nbch = STR_GetNbchamp(spm, "|")
    n = 0
    For i = 1 To nbch
        s = STR_GetChamp(spm, "|", i - 1)
        ReDim Preserve CL_liste.lignes(n)
        CL_liste.lignes(n).texte = s
        CL_liste.lignes(n).fmodif = True
        n = n + 1
    Next i
    
    ssite = ""
    au_moins_un = False
    If p_NbLabo > 1 Then
        For i = 0 To grdLabo.Rows - 1
            If grdLabo.TextMatrix(i, GRDL_ESTLABO) = True Then
                au_moins_un = True
                ssite = ssite & grdLabo.TextMatrix(i, GRDL_NUMLABO) & ";"
            End If
        Next i
        If Not au_moins_un Then
            Call MsgBox("Aucun site n'est indiqué pour cette personne.", vbExclamation + vbOKOnly, "")
            Exit Sub
        End If
    Else
        ssite = ssite + "1;"
    End If
    
    Set frm = KS_PrmService
    sret = KS_PrmService.AppelFrm("Choix des postes", "S", True, ssite, "P", False)
    Set frm = Nothing
    p_numlabo = numlabo
    If sret = "" Then
        Exit Sub
    End If
    
    cmd(CMD_OK).Enabled = True
    
    tvSect.Nodes.Clear
    n = CLng(Mid$(sret, 2))
    For i = 0 To n - 1
        s = CL_liste.lignes(i).texte
        nbch = STR_GetNbchamp(s, ";")
        For j = 1 To nbch
            s1 = STR_GetChamp(s, ";", j - 1)
            If TV_NodeExiste(tvSect, s1, nd) = P_NON Then
                num = Mid$(s1, 2)
                If left$(s1, 1) = "S" Then
                    If P_RecupSrvNom(num, lib) = P_ERREUR Then
                        Exit Sub
                    End If
                    If j = 1 Then
                        Set nd = tvSect.Nodes.Add(, tvwChild, s1, lib, IMGT_SERVICE, IMGT_SERVICE)
                    Else
                        Set nd = tvSect.Nodes.Add(nd, tvwChild, s1, lib, IMGT_SERVICE, IMGT_SERVICE)
                    End If
                Else
                    If P_RecupPosteNom(num, lib) = P_ERREUR Then
                        Exit Sub
                    End If
                    Set nd = tvSect.Nodes.Add(nd, tvwChild, s1, lib, IMGT_POSTE, IMGT_POSTE)
                End If
                nd.Expanded = True
            End If
        Next j
    Next i
    
    tvSect.SetFocus
    If tvSect.Nodes.Count > 0 Then
        cmd(CMD_MOINS_SPM).Visible = True
        Set tvSect.SelectedItem = tvSect.Nodes(1)
        SendKeys "{PGDN}"
        SendKeys "{HOME}"
        DoEvents
    End If
    
End Sub

Private Function quitter(ByVal v_bforce As Boolean) As Boolean

    Dim reponse As Integer
    
    If Not v_bforce Then
        If cmd(CMD_OK).Enabled Then
            If g_numutil = 0 Then
                reponse = MsgBox("La création de cette personne ne s'effectuera pas !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            Else
                reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            End If
            If reponse = vbNo Then
                quitter = False
                Exit Function
            End If
        End If
    End If
    
    Unload Me
    
    quitter = True
    
End Function

Private Function recopier_fct_autor() As Integer

    Dim sql As String
    Dim n As Integer
    Dim num As Long, sav_num As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    'Choix de l'utilisateur
    sql = "select U_num, U_Nom, U_Prenom from Utilisateur" _
        & " where U_Num<>" & g_numutil _
        & " and U_Num>" & P_SUPER_UTIL _
        & " order by U_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        recopier_fct_autor = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("U_Nom").Value & " " & rs("U_Prenom").Value, rs("U_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        recopier_fct_autor = P_OK
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Personne à recopier", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        recopier_fct_autor = P_OK
        Exit Function
    End If
    
    num = CL_liste.lignes(CL_liste.pointeur).num
    sav_num = g_numutil
    g_numutil = num
    If afficher_fct_autor() = P_ERREUR Then
        recopier_fct_autor = P_ERREUR
        Exit Function
    End If
    g_numutil = sav_num
    tvFct.tag = True
    
    cmd(CMD_OK).Enabled = True
    
    recopier_fct_autor = P_OK

End Function

Private Function recup_autorisation(ByVal v_numfct As Long, _
                                    ByRef r_iauto As Integer) As Integer

    Dim i As Integer
    
    If p_NumUtil <> P_SUPER_UTIL Then
        r_iauto = -1
        For i = 0 To CM_UboundL(g_tbl_fctautor1)
            If g_tbl_fctautor1(i) = v_numfct Then
                r_iauto = 1
                Exit For
            End If
        Next i
        If r_iauto = -1 Then
            recup_autorisation = P_OK
            Exit Function
        End If
    End If
    
    For i = 0 To CM_UboundL(g_tbl_fctautor2)
        If g_tbl_fctautor2(i) = v_numfct Then
            r_iauto = 1
            recup_autorisation = P_OK
            Exit Function
        End If
    Next i
    r_iauto = 0
    
    recup_autorisation = P_OK
    
End Function

Private Function recup_poste_corr(ByVal v_numutil As Long, _
                                  ByVal v_numposte As Long, _
                                  ByRef v_tbl_poste() As Long, _
                                  ByRef vr_nposte_corr As Integer, _
                                  ByRef vr_tbl_poste_corr() As SPOSTECORR, _
                                  ByRef r_numposte_corr As Long) As Integer
    Dim libposte As String
    Dim i As Integer
    
    If CM_UboundL(v_tbl_poste()) = 0 Then
        r_numposte_corr = v_tbl_poste(0)
        recup_poste_corr = P_OK
        Exit Function
    End If
    
    For i = 0 To CM_UboundL(v_tbl_poste())
        If v_tbl_poste(i) = v_numposte Then
            r_numposte_corr = v_numposte
            recup_poste_corr = P_OK
            Exit Function
        End If
    Next i
        
    For i = 0 To vr_nposte_corr
        If vr_tbl_poste_corr(i).numposte_aremp = v_numposte Then
            r_numposte_corr = vr_tbl_poste_corr(i).numposte_remp
            recup_poste_corr = P_OK
            Exit Function
        End If
    Next i
    
    If Odbc_RecupVal("select PO_Libelle from Poste where PO_Num=" & v_numposte, _
                     libposte) = P_ERREUR Then
        recup_poste_corr = P_ERREUR
        Exit Function
    End If
    If P_UtilAPoste(v_numutil, v_numposte) Then
        GoTo lab_affecte
    End If
    If P_UtilAPlusieursPostes(v_numutil) = P_OUI Then
        Call MsgBox("Vous devez choisir le poste de la nouvelle personne qui remplacera '" & libposte & "'", vbInformation + vbOKOnly, "")
    End If
lab_choix_fct:
    If P_ChoisirPosteUtilisateur(v_numutil, r_numposte_corr, libposte) <> P_OUI Then
        GoTo lab_choix_fct
    End If
    
lab_affecte:
    vr_nposte_corr = vr_nposte_corr + 1
    ReDim Preserve vr_tbl_poste_corr(vr_nposte_corr) As SPOSTECORR
    vr_tbl_poste_corr(vr_nposte_corr).numposte_aremp = v_numposte
    vr_tbl_poste_corr(vr_nposte_corr).numposte_remp = r_numposte_corr

    recup_poste_corr = P_OK
    
End Function

Private Sub saisir_nbsem(ByVal v_lig As Integer)

End Sub

Private Function supprimer() As Integer

    Dim sql As String
    Dim reponse As Integer, cr As Integer
    Dim lnb As Long
    Dim rs As rdoResultset
    
    reponse = vbNo
    
    If p_appli_kalidoc > 0 Then
        ' Util est paramétré comme acteur dans docsutil, dosutil ou docutil ?
        cr = util_dans_do()
        If cr = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        If cr = P_OUI Then
            MsgBox "Cette personne est un acteur paramétré dans certains documents." & vbLf & vbCr & "Elle ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, ""
            supprimer = P_OK
            Exit Function
        End If
        ' Util avec actions en cours ?
        cr = util_dans_docaction()
        If cr = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        If cr = P_OUI Then
            MsgBox "Cette personne est acteur en cours de certains documents." & vbLf & vbCr & "Elle ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, ""
            supprimer = P_OK
            Exit Function
        End If
        ' Util ayant effectué des actions ?
        cr = util_dans_docetapeversion()
        If cr = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        If cr = P_OUI Then
            MsgBox "Cette personne a été acteur dans certains documents." & vbLf & vbCr & "Elle ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, ""
            supprimer = P_OK
            Exit Function
        End If
        ' Util avec diffusions ?
        cr = util_dans_diff()
        If cr = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        If cr = P_OUI Then
            If p_NumUtil = P_SUPER_UTIL Then
                reponse = MsgBox("ATTENTION : Certains documents ont été diffusés à cette personne." & vbLf & vbCr & "Confirmez-vous quand même la suppression de cette personne ?", vbQuestion + vbYesNo, "")
                If reponse = vbNo Then
                    supprimer = P_OK
                    Exit Function
                End If
            Else
                MsgBox "Certains documents ont été diffusés à cette personne." & vbLf & vbCr & "Elle ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, ""
                supprimer = P_OK
                Exit Function
            End If
        End If
    End If
    
    If reponse = vbNo Then
        reponse = MsgBox("Confirmez-vous la suppression de cette personne ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
        If reponse = vbNo Then
            supprimer = P_OK
            Exit Function
        End If
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    ' Effacement de l'utilisateur
    If Odbc_Delete("Utilisateur", _
                   "U_Num", _
                   "where U_Num=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    ' Effacement de ses coordonnées
    If Odbc_Delete("UtilCoordonnee", _
                   "UC_Num", _
                   "where UC_Type='U' and UC_TypeNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    ' Effacement de ses codes dans les applications
    If Odbc_Delete("UtilAppli", _
                   "UAPP_Num", _
                   "where UAPP_UNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    If Odbc_Delete("UtilADIM", _
                   "UA_Num", _
                   "where UA_UNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    If Odbc_Delete("DocPrmDiffusion", _
                   "DPD_Num", _
                   "where DPD_UNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    If Odbc_Delete("DocDiffusion", _
                   "DD_Num", _
                   "where DD_UNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    ' Effacement de ses droits aux fonctions
    If Odbc_Delete("FctOK_Util", _
                   "FU_Num", _
                   "where FU_UNum=" & g_numutil, _
                   lnb) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    If Odbc_CommitTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    supprimer = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    supprimer = P_ERREUR
    Exit Function
    
End Function

Private Sub supprimer_postetrav()

    If grdPoste.Rows = grdPoste.FixedRows + 1 Then
        grdPoste.Rows = grdPoste.FixedRows
        cmd(CMD_MOINS_POSTE).Visible = False
    Else
        grdPoste.RemoveItem (grdPoste.row)
        grdPoste.row = grdPoste.FixedRows
    End If
    
    grdPoste.tag = "M"
    grdPoste.SetFocus
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub supprimer_poste()

    Dim encore As Boolean
    Dim nd As Node, ndp As Node
    
    If tvSect.Nodes.Count = 0 Then
        Exit Sub
    End If
    
    On Error GoTo err_tv
    Set nd = tvSect.SelectedItem
    On Error GoTo 0
    
    Do
        encore = True
        Set ndp = nd
        If TV_NodeParent(ndp) Then
            If ndp.Children > 1 Then
                encore = False
            Else
                Set nd = ndp
            End If
        Else
            encore = False
        End If
    Loop Until Not encore
        
    tvSect.Nodes.Remove (nd.Index)
        
    tvSect.Refresh
    cmd(CMD_OK).Enabled = True
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_MOINS_SPM).Visible = False
    End If
    
    Exit Sub
    
err_tv:
    MsgBox "Vous devez sélectionner l'élément à supprimer", vbOKOnly, ""
    On Error GoTo 0
    
End Sub

Private Function util_dans_do() As Integer

    Dim sql As String
    Dim nb As Long
    
    sql = "select count(*) from DocsUtil" _
        & " where DOU_UNum=" & g_numutil _
        & " and DOU_CYOrdre<>" & P_DESTINATAIRE
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Documentation" _
        & " where DO_LstResp like '%U" & g_numutil & ";%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from DosUtil" _
        & " where DSU_UNum=" & g_numutil _
        & " and DSU_CYOrdre<>" & P_DESTINATAIRE
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Dossier" _
        & " where DS_LstResp like '%U" & g_numutil & ";%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from DocUtil" _
        & " where DU_UNum=" & g_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Document" _
        & " where D_LstResp like '%U" & g_numutil & ";%'" _
        & " or D_UNumResp=" & g_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_do = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_do = P_OUI
        Exit Function
    End If
    
    util_dans_do = P_NON
    
End Function
 
Private Function util_dans_diff() As Integer

    Dim sql As String
    Dim nb As Long
    
    sql = "select count(*) from DocDiffusion" _
        & " where DD_UNum=" & g_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_diff = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_diff = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like '%U" & g_numutil & "|%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_diff = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_diff = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like '%U" & g_numutil & "|%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_diff = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_diff = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Document" _
        & " where D_Dest like '%U" & g_numutil & "|%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_diff = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_diff = P_OUI
        Exit Function
    End If
    
    util_dans_diff = P_NON

End Function
    
Private Function util_dans_docaction() As Integer

    Dim sql As String
    Dim nb As Long
    
    sql = "select count(*) from DocAction" _
        & " where DAC_UNum=" & g_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_docaction = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_docaction = P_OUI
        Exit Function
    End If
    
    util_dans_docaction = P_NON

End Function

Private Function util_dans_docetapeversion() As Integer

    Dim sql As String
    Dim nb As Long
    
    sql = "select count(*) from DocEtapeVersion" _
        & " where DEV_UNum=" & g_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_docetapeversion = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_docetapeversion = P_OUI
        Exit Function
    End If
    
    util_dans_docetapeversion = P_NON

End Function

Private Function util_dans_groupeutil() As Integer

    Dim sql As String
    Dim nb As Long
    
    sql = "select count(*) from GroupeUtil" _
        & " where GU_Lst like '%U" & g_numutil & "|%'"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        util_dans_groupeutil = P_ERREUR
        Exit Function
    End If
    If nb > 0 Then
        util_dans_groupeutil = P_OUI
        Exit Function
    End If
    
    util_dans_groupeutil = P_NON

End Function


Private Function valider() As Integer

    Dim sql As String, ssite As String, old_ssite As String
    Dim sfct As String, old_sfct As String
    Dim nom_avt As String, prenom_avt As String, matricule_avt As String
    Dim bactif As Boolean
    Dim i As Integer, fmaj_dest As Integer, n As Integer, nb As Integer, ilig As Integer
    Dim reponse As Integer, cr As Integer, nbheures As Integer, base_heures As Integer, uheures As Integer
    Dim nbsem As Integer
    Dim num_util As Long, image As Long, numlabop As Long, lbid As Long, lnb As Long
    Dim datdeb As Date, datfin As Date
    Dim spm As Variant, old_spm As Variant, spm_avt As Variant
    Dim rs As rdoResultset
    
    If verifier_tous_champs(numlabop, ssite, sfct, spm) = P_NON Then
        valider = P_NON
        Exit Function
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        valider = P_ERREUR
        Exit Function
    End If
    
    fmaj_dest = 0
    If p_appli_kaliress > 0 Then
        nbheures = txt(TXT_NBHEURES).Text
        base_heures = cbo(CBO_BASEHEURES).ListIndex
    Else
        nbheures = 0
        base_heures = 0
    End If
    If txt(TXT_DATEDEB_EMBAUCHE).Text <> "" Then
        datdeb = CDate(txt(TXT_DATEDEB_EMBAUCHE).Text)
    End If
    If txt(TXT_DATEFIN_EMBAUCHE).Text <> "" Then
        datfin = CDate(txt(TXT_DATEFIN_EMBAUCHE).Text)
    End If
    If g_numutil = 0 Then
        If Odbc_AddNew("Utilisateur", "U_Num", "u_seq", True, num_util, _
                        "U_Nom", txt(TXT_NOM).Text, "U_Prenom", txt(TXT_PRENOM).Text, _
                        "U_Prefixe", txt(TXT_PREFIXE).Text, _
                        "U_Actif", IIf(chk(CHK_ACTIF).Value = 1, True, False), _
                        "U_Externe", IIf(chk(CHK_EXTERNE).Value = 1, True, False), _
                        "U_ExterneFich", False, "U_Importe", False, _
                        "U_Fictif", IIf(chk(CHK_FICTIF).Value = 1, True, False), _
                        "U_AR", IIf(chk(CHK_AR).Value = 1, True, False), _
                        "U_FctTrav", sfct, "U_SPM", spm, _
                        "U_FctTrav_kb", IIf(p_appli_kalibottin > 0, "", sfct), _
                        "U_SPM_kb", IIf(p_appli_kalibottin > 0, "", spm), _
                        "U_POB_Princ", 0, _
                        "U_Labo", ssite, "U_LNumPrinc", numlabop, _
                        "U_Matricule", txt(TXT_MATRICULE).Text, _
                        "U_DONumLast", 0, "U_LstDocs", "", _
                        "U_CATPNum", txt(TXT_CATEGPROF).tag, _
                        "U_DateDebEmbauche", IIf(txt(TXT_DATEDEB_EMBAUCHE).Text = "", Null, datdeb), _
                        "U_DateFinEmbauche", IIf(txt(TXT_DATEDEB_EMBAUCHE).Text = "", Null, datfin), _
                        "U_CTRAVNum", txt(TXT_TYPECONTRAT).tag, _
                        "U_NbHeures", nbheures, _
                        "U_BaseHeures", base_heures, _
                        "U_POTNumNext", 0, _
                        "U_LNumNext", 0, _
                        "U_NoSemNext", 0, _
                        "U_kw_mailauth", True) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' Code + Mot de Passe
        If Odbc_AddNew("UtilAppli", _
                        "UAPP_Num", _
                        "uapp_seq", _
                        False, _
                        lbid, _
                        "UAPP_APPNum", p_appli_kalidoc, _
                        "UAPP_UNum", num_util, _
                        "UAPP_Code", UCase(txt(TXT_CODE).Text), _
                        "UAPP_MotPasse", STR_Crypter(UCase(txt(TXT_MPASSE).Text))) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' Adr Mail
        If txt(TXT_ADRNET).Enabled And txt(TXT_ADRNET).Text <> "" Then
            If gerer_adrmail(num_util, 1) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
        If chk(CHK_ACTIF).Value = 1 Then
            fmaj_dest = 2
            Call ajouter_mouvement(num_util, "C", "")
        End If
    Else
        num_util = g_numutil
        If Odbc_RecupVal("select U_Actif, U_Labo, U_FctTrav, U_SPM from Utilisateur where U_Num=" & g_numutil, _
                          bactif, _
                          old_ssite, _
                          old_sfct, _
                          old_spm) = P_ERREUR Then
            GoTo err_enreg
        End If
        If chk(CHK_ACTIF).Value = 0 Then
            If bactif Then
                fmaj_dest = 1
                Call ajouter_mouvement(num_util, "I", "")
            End If
        Else
            If Not bactif Then
                fmaj_dest = 2
                Call ajouter_mouvement(num_util, "A", "")
            Else
                If Not ya_meme_labo_fct_spm(old_ssite, ssite, old_sfct, sfct, old_spm, spm) Then
                    fmaj_dest = 3
                End If
            End If
        End If
        sql = "select U_Nom, U_Prenom, U_Matricule, U_SPM from Utilisateur" _
            & " where U_Num=" & g_numutil
        If Odbc_RecupVal(sql, nom_avt, prenom_avt, matricule_avt, spm_avt) = P_ERREUR Then
            GoTo err_enreg
        End If
        If Odbc_Update("Utilisateur", _
                       "U_Num", _
                       "where U_Num=" & g_numutil, _
                        "U_Nom", txt(TXT_NOM).Text, _
                        "U_Prenom", txt(TXT_PRENOM).Text, "U_Prefixe", txt(TXT_PREFIXE).Text, _
                        "U_Actif", IIf(chk(CHK_ACTIF).Value = 1, True, False), _
                        "U_Externe", IIf(chk(CHK_EXTERNE).Value = 1, True, False), _
                        "U_Fictif", IIf(chk(CHK_FICTIF).Value = 1, True, False), _
                        "U_AR", IIf(chk(CHK_AR).Value = 1, True, False), _
                        "U_FctTrav", sfct, _
                        "U_SPM", spm, _
                        "U_Labo", ssite, _
                        "U_LNumPrinc", numlabop, _
                        "U_Matricule", txt(TXT_MATRICULE).Text, _
                        "U_CATPNum", txt(TXT_CATEGPROF).tag, _
                        "U_DateDebEmbauche", IIf(txt(TXT_DATEDEB_EMBAUCHE).Text = "", Null, datdeb), _
                        "U_DateFinEmbauche", IIf(txt(TXT_DATEFIN_EMBAUCHE).Text = "", Null, datfin), _
                        "U_CTRAVNum", txt(TXT_TYPECONTRAT).tag, _
                        "U_NbHeures", nbheures, _
                        "U_BaseHeures", base_heures, _
                        "U_POTNumNext", 0, _
                        "U_LNumNext", 0, _
                        "U_NoSemNext", 0) = P_ERREUR Then
            GoTo err_enreg
        End If
        If p_appli_kalibottin = 0 Then
            If Odbc_Update("Utilisateur", _
                           "U_Num", _
                           "where U_Num=" & g_numutil, _
                            "U_FctTrav_kb", sfct, _
                            "U_SPM_kb", spm) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
        ' Code + Mot de Passe
        If Odbc_Update("UtilAppli", _
                       "UAPP_Num", _
                       "where UAPP_UNum=" & g_numutil _
                            & " and UAPP_APPNum=" & p_appli_kalidoc, _
                        "UAPP_Code", UCase(txt(TXT_CODE).Text), _
                        "UAPP_MotPasse", STR_Crypter(UCase(txt(TXT_MPASSE).Text))) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' Adr Mail
        If txt(TXT_ADRNET).Enabled And txt(TXT_ADRNET).Text <> txt(TXT_ADRNET).tag Then
            If txt(TXT_ADRNET).Text = "" Then
                If gerer_adrmail(num_util, -1) = P_ERREUR Then
                    GoTo err_enreg
                End If
            ElseIf txt(TXT_ADRNET).tag = "" Then
                If gerer_adrmail(num_util, 1) = P_ERREUR Then
                    GoTo err_enreg
                End If
            Else
                If gerer_adrmail(num_util, 0) = P_ERREUR Then
                    GoTo err_enreg
                End If
            End If
        End If
        ' Mouvements
        If nom_avt <> txt(TXT_NOM).Text Then
            Call ajouter_mouvement(num_util, "M", "NOM=" & nom_avt)
        End If
        If prenom_avt <> txt(TXT_PRENOM).Text Then
            Call ajouter_mouvement(num_util, "M", "PRENOM=" & prenom_avt)
        End If
        If matricule_avt <> txt(TXT_MATRICULE).Text Then
            Call ajouter_mouvement(num_util, "M", "MATRICULE=" & matricule_avt)
        End If
    End If
    
    If p_appli_kalidoc > 0 Then
        ' Suppression de l'utilisateur
        If fmaj_dest = 1 Then
            Call afficher_frm_valid
            If gerer_suppr_utilisateur(g_numutil) = P_ERREUR Then
                GoTo err_enreg
            End If
        ' Nouvel utilisateur
        ElseIf fmaj_dest = 2 Then
            Call afficher_frm_valid
            If gerer_nouvel_utilisateur(num_util) = P_ERREUR Then
                GoTo err_enreg
            End If
        ' Changement de labo/fct/SPM
        ElseIf fmaj_dest = 3 Then
            Call afficher_frm_valid
            If gerer_chgt_prmutil(g_numutil, _
                                  IIf(chk(CHK_AR).Value = 1, True, False), _
                                  old_ssite, _
                                  ssite, _
                                  old_sfct, _
                                  sfct, _
                                  old_spm, _
                                  spm) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
        Call inhiber_frm_valid
    End If
    
    If p_appli_kaliress > 0 Then
        If grdPoste.tag <> "" Then
            If g_numutil > 0 Then
                If Odbc_Delete("UtilPosteTrav", _
                               "UPOT_Num", _
                                "where UPOT_UNum=" & g_numutil, _
                                lnb) = P_ERREUR Then
                    GoTo err_enreg
                End If
            End If
            For ilig = grdPoste.FixedRows To grdPoste.Rows - 1
                If grdPoste.TextMatrix(ilig, GRDP_NBSEM_TOURNANTE) = "" Then
                    nbsem = 0
                Else
                    nbsem = CInt(grdPoste.TextMatrix(ilig, GRDP_NBSEM_TOURNANTE))
                End If
                If Odbc_AddNew("UtilPosteTrav", _
                               "UPOT_Num", _
                               "upopt_seq", _
                               False, _
                               lbid, _
                               "UPOT_UNum", num_util, _
                               "UPOT_LNum", grdPoste.TextMatrix(ilig, GRDP_NUMLABO), _
                               "UPOT_POTNum", grdPoste.TextMatrix(ilig, GRDP_NUMPOSTE), _
                               "UPOT_TPONum", grdPoste.TextMatrix(ilig, GRDP_NUMTITRE), _
                               "UPOT_Ordre", ilig, _
                               "UPOT_GestAstreinte", grdPoste.TextMatrix(ilig, GRDP_ASTREINTE), _
                               "UPOT_GestGarde", grdPoste.TextMatrix(ilig, GRDP_GARDE), _
                               "UPOT_CYTNum", grdPoste.TextMatrix(ilig, GRDP_NUMCYCLE), _
                               "UPOT_TRMOrdreNext", IIf(grdPoste.TextMatrix(ilig, GRDP_NUMCYCLE) > 0, grdPoste.TextMatrix(ilig, GRDP_POSTRMNEXT_NUMTRM), 0), _
                               "UPOT_TRMNum", IIf(grdPoste.TextMatrix(ilig, GRDP_NUMCYCLE) = 0, grdPoste.TextMatrix(ilig, GRDP_POSTRMNEXT_NUMTRM), 0), _
                               "UPOT_Tournante", grdPoste.TextMatrix(ilig, GRDP_TOURNANTE), _
                               "UPOT_NbSemTournante", nbsem) = P_ERREUR Then
                    GoTo err_enreg
                End If
            Next ilig
        End If
    End If
    
    If tvFct.tag = False Then GoTo lab_commit
    
    If g_numutil > 0 Then
        If Odbc_Delete("FctOK_Util", _
                       "FU_Num", _
                       "where FU_UNum=" & g_numutil, _
                       lnb) = P_ERREUR Then
            GoTo err_enreg
        End If
    End If
    
    For i = 1 To tvFct.Nodes.Count
        image = tvFct.Nodes(i).image
        If image = BOULE_VERTE Or image = CARRE_VERT Then
            If Odbc_AddNew("FctOK_Util", _
                           "FU_Num", _
                           "fu_seq", _
                           False, _
                           lbid, _
                           "FU_UNum", num_util, _
                           "FU_FCTNum", Mid$(tvFct.Nodes(i).key, 2)) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
    Next i
    If Odbc_AddNew("FctOK_Util", _
                   "FU_Num", _
                   "fu_seq", _
                   False, _
                   lbid, _
                   "FU_UNum", num_util, _
                   "FU_FCTNum", P_FCT_CHGUTIL) = P_ERREUR Then
        GoTo err_enreg
    End If

lab_commit:
    If Odbc_CommitTrans() = P_ERREUR Then
        valider = P_ERREUR
        Exit Function
    End If
    
    If p_NumUtil = g_numutil Then
        Call P_ChargerFctAutor
        Call maj_droits
    End If
    
    valider = P_OUI
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    valider = P_ERREUR
    Exit Function
    
    
End Function

Private Function verifier_champ(ByVal v_indtxt As Integer) As Integer

    Dim sql As String, s As String
    Dim reponse As Integer
    Dim mess As Variant
    Dim rs As rdoResultset
    
    Select Case v_indtxt
    Case TXT_NOM
        If txt(v_indtxt).Text <> "" Then
            sql = "select U_Nom, U_Prenom from Utilisateur" _
                & " where U_Nom=" & Odbc_String(UCase(txt(v_indtxt).Text)) _
                & " and U_Num <> " & g_numutil
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                verifier_champ = P_NON
                Exit Function
            End If
            mess = "Il existe déjà des personnes avec ce nom :" & vbCrLf
            If Not rs.EOF Then
                While Not rs.EOF
                    mess = mess & "     " & rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbCrLf
                    rs.MoveNext
                Wend
                rs.Close
                reponse = MsgBox(mess & vbCrLf & vbCrLf & "Voulez-vous poursuivre la création ?", vbQuestion + vbYesNo, "")
                If reponse = vbNo Then
                    Call quitter(True)
                    verifier_champ = P_ERREUR
                    Exit Function
                End If
            Else
                rs.Close
            End If
        End If
        verifier_champ = P_OUI
        Exit Function
    Case TXT_CODE
        If txt(v_indtxt).Text = "" Then
            verifier_champ = P_OUI
            Exit Function
        End If
        If txt(v_indtxt).Text = "ROOT" Then
            MsgBox "Code d'accès réservé.", vbOKOnly + vbExclamation, ""
            verifier_champ = P_NON
            Exit Function
        End If
        s = UCase(txt(v_indtxt).Text)
        sql = "select UAPP_UNum from UtilAppli" _
            & " where UAPP_Code=" & Odbc_String(txt(v_indtxt).Text) _
            & " and UAPP_APPNum=" & p_appli_kalidoc
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            verifier_champ = P_NON
            Exit Function
        End If
        If Not rs.EOF Then
            If rs("UAPP_UNum").Value <> g_numutil Then
                rs.Close
                MsgBox "Code d'accès déjà attribué.", vbOKOnly + vbExclamation, ""
                verifier_champ = P_NON
                Exit Function
            End If
        End If
        rs.Close
        txt(v_indtxt).Text = s
        verifier_champ = P_OUI
        Exit Function
    Case TXT_DATEDEB_EMBAUCHE, TXT_DATEFIN_EMBAUCHE
        If txt(v_indtxt).Text = "" Then
            verifier_champ = P_OUI
            Exit Function
        End If
        s = txt(v_indtxt).Text
        If Not SAIS_CtrlChamp(s, SAIS_TYP_DATE) Then
            verifier_champ = P_NON
        Else
            txt(v_indtxt).Text = s
            verifier_champ = P_OUI
        End If
        Exit Function
    Case TXT_NBHEURES
        If txt(v_indtxt).Text = "" Then
            verifier_champ = P_OUI
            Exit Function
        End If
        If Not SAIS_CtrlChamp(txt(v_indtxt).Text, SAIS_TYP_ENTIER) Then
            verifier_champ = P_NON
        Else
            verifier_champ = P_OUI
        End If
        Exit Function
    End Select
    
    verifier_champ = P_OUI
    
End Function

Private Function verifier_tous_champs(ByRef r_numlabop As Long, _
                                      ByRef r_ssite As String, _
                                      ByRef r_sfct As String, _
                                      ByRef r_spm As Variant) As Integer

    Dim ilig As Integer
    
    If txt(TXT_CODE).Text = "" Then
        MsgBox "Le CODE de la personne est une rubrique obligatoire.", vbOKOnly + vbExclamation, ""
        sst.Tab = 0
        txt(TXT_CODE).SetFocus
        verifier_tous_champs = P_NON
        Exit Function
    End If
    
    If p_appli_kaliress > 0 Then
        If txt(TXT_DATEDEB_EMBAUCHE).Text = "" Then
            Call MsgBox("Veuillez indiquer la date d'embauche.", vbOKOnly + vbExclamation, "")
            sst.Tab = 3
            txt(TXT_DATEDEB_EMBAUCHE).SetFocus
            verifier_tous_champs = P_NON
            Exit Function
        ElseIf verifier_champ(TXT_DATEDEB_EMBAUCHE) = P_NON Then
            sst.Tab = 3
            txt(TXT_DATEDEB_EMBAUCHE).SetFocus
            verifier_tous_champs = P_NON
            Exit Function
        End If
        If verifier_champ(TXT_DATEFIN_EMBAUCHE) = P_NON Then
            sst.Tab = 3
            txt(TXT_DATEDEB_EMBAUCHE).SetFocus
            verifier_tous_champs = P_NON
            Exit Function
        End If
        If txt(TXT_NBHEURES).Text = "" Then txt(TXT_NBHEURES).Text = "0"
        If verifier_champ(TXT_NBHEURES) = P_NON Then
            sst.Tab = 3
            txt(TXT_NBHEURES).SetFocus
            verifier_tous_champs = P_NON
            Exit Function
        End If
        If cbo(CBO_BASEHEURES).ListIndex = -1 Then
            Call MsgBox("Veuillez indiquer sur quelle période s'effectue le calcul des heures effectuées.", vbOKOnly + vbExclamation, "")
            sst.Tab = 3
            cbo(CBO_BASEHEURES).SetFocus
            verifier_tous_champs = P_NON
            Exit Function
        End If
    End If
    
    ' Construction U_SPM et U_FctTRav
    If tvSect.Nodes.Count = 0 Then
        Call MsgBox("Veuillez indiquer le poste affectée à la personne.", vbOKOnly + vbExclamation, "")
        sst.Tab = 1
        tvSect.SetFocus
        verifier_tous_champs = P_NON
        Exit Function
    End If
    Call build_SPM_Fct(r_spm, r_sfct)
    
    ' Construction U_Labo
    r_ssite = ""
    r_numlabop = 0
    For ilig = 0 To grdLabo.Rows - 1
        If grdLabo.TextMatrix(ilig, GRDL_ESTLABO) = True Then
            r_ssite = r_ssite + "L" + grdLabo.TextMatrix(ilig, GRDL_NUMLABO) + ";"
        End If
        If grdLabo.TextMatrix(ilig, GRDL_LABOPRINC) <> "" Then
            r_numlabop = grdLabo.TextMatrix(ilig, GRDL_NUMLABO)
        End If
    Next ilig
    If r_numlabop = 0 Then
        MsgBox "Indiquez le site principal.", vbOKOnly + vbExclamation, ""
        grdLabo.SetFocus
        verifier_tous_champs = P_NON
        Exit Function
    End If
    
    verifier_tous_champs = P_OUI
    
End Function

Private Function ya_meme_labo_fct_spm(ByVal v_o_ssite As String, _
                                      ByVal v_ssite As String, _
                                      ByVal v_o_sfct As String, _
                                      ByVal v_sfct As String, _
                                      ByVal v_o_spm As Variant, _
                                      ByVal v_spm As Variant) As Boolean
                                      
    Dim s1 As String, s2 As String
    Dim trouve As Boolean
    Dim n1 As Integer, n2 As Integer, i As Integer, j As Integer
    Dim num1 As Long, num2 As Long
    
    n1 = STR_GetNbchamp(v_o_ssite, ";")
    n2 = STR_GetNbchamp(v_ssite, ";")
    For i = 0 To n1 - 1
        num1 = CLng(Mid$(STR_GetChamp(v_o_ssite, ";", i), 2))
        trouve = False
        For j = 0 To n2 - 1
            num2 = CLng(Mid$(STR_GetChamp(v_ssite, ";", j), 2))
            If num2 = num1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    For i = 0 To n2 - 1
        num2 = CLng(Mid$(STR_GetChamp(v_ssite, ";", i), 2))
        trouve = False
        For j = 0 To n1 - 1
            num1 = CLng(Mid$(STR_GetChamp(v_o_ssite, ";", j), 2))
            If num2 = num1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    
    n1 = STR_GetNbchamp(v_o_sfct, ";")
    n2 = STR_GetNbchamp(v_sfct, ";")
    For i = 0 To n1 - 1
        num1 = CLng(Mid$(STR_GetChamp(v_o_sfct, ";", i), 2))
        trouve = False
        For j = 0 To n2 - 1
            num2 = CLng(Mid$(STR_GetChamp(v_sfct, ";", j), 2))
            If num2 = num1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    For i = 0 To n2 - 1
        num2 = CLng(Mid$(STR_GetChamp(v_sfct, ";", i), 2))
        trouve = False
        For j = 0 To n1 - 1
            num1 = CLng(Mid$(STR_GetChamp(v_o_sfct, ";", j), 2))
            If num2 = num1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    
    n1 = STR_GetNbchamp(v_o_spm, "|")
    n2 = STR_GetNbchamp(v_spm, "|")
    For i = 0 To n1 - 1
        s1 = STR_GetChamp(v_o_spm, "|", i)
        trouve = False
        For j = 0 To n2 - 1
            s2 = STR_GetChamp(v_spm, "|", j)
            If s2 = s1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    For i = 0 To n2 - 1
        s1 = STR_GetChamp(v_spm, "|", i)
        trouve = False
        For j = 0 To n1 - 1
            s2 = STR_GetChamp(v_o_spm, "|", j)
            If s2 = s1 Then
                trouve = True
                Exit For
            End If
        Next j
        If Not trouve Then
            ya_meme_labo_fct_spm = False
            Exit Function
        End If
    Next i
    
    ya_meme_labo_fct_spm = True
    
End Function

Private Sub cbo_GotFocus(Index As Integer)

    g_cbo_avant = cbo(Index).ListIndex
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub cbo_LostFocus(Index As Integer)

    Dim i As Integer
    
    If Not g_mode_saisie Then Exit Sub
    
    If cbo(Index).ListIndex = -1 And cbo(Index).Text <> "" Then
        g_mode_saisie = False
        For i = 0 To cbo(Index).ListCount - 1
            If left$(cbo(Index).List(i), Len(cbo(Index).Text)) = cbo(Index).Text Then
                cbo(Index).ListIndex = i
                Exit For
            End If
        Next i
        g_mode_saisie = True
        If cbo(Index).ListIndex = -1 Then
            cbo(Index).Text = ""
            cbo(Index).SetFocus
            Exit Sub
        End If
    End If
        
    If cbo(Index).ListIndex <> g_cbo_avant Then
        cmd(CMD_OK).Enabled = True
    End If
    
End Sub

Private Sub chk_Click(Index As Integer)

    If g_mode_saisie Then
        If Index = CHK_FICTIF Then
            If chk(CHK_FICTIF).Value = 1 Then
                chk(CHK_AR).Value = 0
                chk(CHK_AR).Enabled = False
            Else
                chk(CHK_AR).Value = 1
                chk(CHK_AR).Enabled = True
            End If
        End If
        cmd(CMD_OK).Enabled = True
    End If
    
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        If valider() <> P_NON Then
            Unload Me
            Exit Sub
        End If
    Case CMD_QUITTER
        Call quitter(False)
    Case CMD_SUPPRIMER
        Call supprimer
        Unload Me
    Case CMD_RECOPIE
        If recopier_fct_autor() = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
    Case CMD_DEM_KB
        Call envoyer_demande_kb
    Case CMD_CHOIX_CATEGPROF
        Call choisir_categprof
    Case CMD_CHOIX_TYPECONTRAT
        Call choisir_contrat
    Case CMD_ACCES_SPM
        Call prm_service
    Case CMD_MOINS_SPM
        Call supprimer_poste
    Case CMD_BASCULE_AUTOR
        Call basculer_autor
    Case CMD_PLUS_POSTE
        Call ajouter_postetrav
    Case CMD_MOINS_POSTE
        Call supprimer_postetrav
    End Select
    
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = CMD_QUITTER Then g_mode_saisie = False

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub
    
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then
            If valider() <> P_NON Then
                Unload Me
                Exit Sub
            End If
        End If
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_SUPPRIMER).Enabled Then
            Call supprimer
            Call quitter(True)
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_e_utilisateur.htm")
    ElseIf KeyCode = vbKeyPageUp Then
        Call afficher_page(0)
    ElseIf KeyCode = vbKeyPageDown Then
        Call afficher_page(1)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False
    
    g_form_width = Me.Width
    g_form_height = Me.Height
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If Not quitter(False) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub grdLabo_DblClick()

    If grdLabo.col = GRDL_IMG_ESTLABO Then
        Call inverser_etat_labo
    ElseIf grdLabo.col = GRDL_LABOPRINC Then
        Call inverser_laboprinc
    End If
    
End Sub

Private Sub grdLabo_GotFocus()

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    grdLabo.col = GRDL_CODLABO
    grdLabo.ColSel = GRDL_CODLABO

End Sub

Private Sub grdLabo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        KeyCode = 0
        If grdLabo.col = GRDL_IMG_ESTLABO Then
            Call inverser_etat_labo
        ElseIf grdLabo.col = GRDL_LABOPRINC Then
            Call inverser_laboprinc
        End If
    End If
    
End Sub

Private Sub grdLabo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub grdPoste_DblClick()

    Dim i As Integer, n As Integer
    
    Select Case grdPoste.col
    Case GRDP_NOMTITRE
        Call choisir_titre
    Case GRDP_PIC_ASTREINTE
        If grdPoste.TextMatrix(grdPoste.row, GRDP_ASTREINTE_POSSIBLE) = True Then
            If grdPoste.TextMatrix(grdPoste.row, GRDP_ASTREINTE) = True Then
                grdPoste.TextMatrix(grdPoste.row, GRDP_ASTREINTE) = False
                Set grdPoste.CellPicture = CM_LoadPicture("")
            Else
                grdPoste.TextMatrix(grdPoste.row, GRDP_ASTREINTE) = True
                Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
            End If
        End If
        grdPoste.tag = "M"
        cmd(CMD_OK).Enabled = True
    Case GRDP_PIC_GARDE
        If grdPoste.TextMatrix(grdPoste.row, GRDP_GARDE_POSSIBLE) = True Then
            If grdPoste.TextMatrix(grdPoste.row, GRDP_GARDE) = True Then
                grdPoste.TextMatrix(grdPoste.row, GRDP_GARDE) = False
                Set grdPoste.CellPicture = CM_LoadPicture("")
            Else
                grdPoste.TextMatrix(grdPoste.row, GRDP_GARDE) = True
                Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
            End If
        End If
        grdPoste.tag = "M"
        cmd(CMD_OK).Enabled = True
    Case GRDP_NOMCYCLE
        Call choisir_cycletrame(grdPoste.row)
    Case GRDP_NOM_TRMNEXT_TRM
        If grdPoste.TextMatrix(grdPoste.row, GRDP_NUMCYCLE) = 0 Then
            Call choisir_tramehebdo(grdPoste.row)
        ElseIf grdPoste.TextMatrix(grdPoste.row, GRDP_NUMCYCLE) > 0 Then
            Call choisir_next_trame(grdPoste.row)
        End If
    Case GRDP_PIC_TOURNANTE
        If grdPoste.TextMatrix(grdPoste.row, GRDP_TOURNANTE) = True Then
            grdPoste.TextMatrix(grdPoste.row, GRDP_TOURNANTE) = False
            Set grdPoste.CellPicture = CM_LoadPicture("")
        Else
            grdPoste.TextMatrix(grdPoste.row, GRDP_TOURNANTE) = True
            Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        End If
        n = 0
        For i = 1 To grdPoste.Rows - 1
            If grdPoste.TextMatrix(i, GRDP_TOURNANTE) = True Then
                n = n + 1
            End If
        Next i
        If n > 1 Then
            frmNext.Visible = True
        Else
            frmNext.Visible = False
        End If
        grdPoste.tag = "M"
        cmd(CMD_OK).Enabled = True
    Case GRDP_NBSEM_TOURNANTE
        Call saisir_nbsem(grdPoste.row)
    End Select
    
End Sub

Private Sub grdposte_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyAdd Or KeyCode = 187 Then
        KeyCode = 0
        Call ajouter_postetrav
    ElseIf KeyCode = vbKeySubtract Or KeyCode = 54 Then
        KeyCode = 0
        Call supprimer_postetrav
    End If
    
End Sub

Private Sub mnuResp_Click()

    Call basculer_etat_resp
    
End Sub

Private Sub sst_Click(PreviousTab As Integer)

    If Not g_mode_saisie Then
        Exit Sub
    End If

    Call init_focus
    
End Sub

Private Sub tvfct_Collapse(ByVal Node As ComctlLib.Node)

    Node.Expanded = True
    
End Sub

Private Sub tvfct_DblClick()

    If g_modopt_autor Then
        Call inverser_autor
    End If
    
End Sub

Private Sub tvFct_GotFocus()

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
End Sub

Private Sub tvfct_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If p_appli_kaliress > 0 Then
            sst.Tab = 3
        Else
            sst.Tab = 4
        End If
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
        If g_modopt_autor Then Call inverser_autor
    End If
    
End Sub

Private Sub tvSect_BeforeLabelEdit(Cancel As Integer)

    Cancel = True
    
End Sub

Private Sub tvSect_dblClick()

    tvSect.SelectedItem.Expanded = True

End Sub

Private Sub tvSect_GotFocus()

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    If tvSect.Nodes.Count > 0 Then
        Set tvSect.SelectedItem = tvSect.Nodes(1)
    End If
    
End Sub

Private Sub tvSect_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call prm_service
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_poste
    End If
    
End Sub

Private Sub tvSect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        sst.Tab = 2
    End If
    
End Sub

Private Sub tvSect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    g_button = Button
    
End Sub

Private Sub txt_Change(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)

    If Not g_mode_saisie Then
        Exit Sub
    End If

    g_txt_avant = txt(Index).Text
    
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        If Index = TXT_CATEGPROF Then
            Call choisir_categprof
        ElseIf Index = TXT_TYPECONTRAT Then
            Call choisir_contrat
        End If
    End If
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If p_appli_kaliress > 0 Then
            If Index = TXT_DATEFIN_EMBAUCHE Then
                sst.Tab = 1
                Exit Sub
            End If
        Else
            If Index = TXT_MATRICULE Then
                sst.Tab = 1
                Exit Sub
            End If
        End If
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txt_lostfocus(Index As Integer)

    Dim cr As Integer
    
    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            cr = verifier_champ(Index)
            If cr = P_ERREUR Then
                Exit Sub
            ElseIf cr = P_NON Then
                txt(Index).Text = ""
                txt(Index).SetFocus
                Exit Sub
            End If
        End If
    End If

End Sub
