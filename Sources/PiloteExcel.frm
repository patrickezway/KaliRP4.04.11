VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PiloteExcel 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix des relecteurs"
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
      Height          =   8190
      Left            =   0
      TabIndex        =   3
      Top             =   -15
      Width           =   15015
      Begin VB.Frame Frame2 
         Height          =   7935
         Left            =   120
         TabIndex        =   7
         Top             =   -120
         Width           =   14775
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
            Index           =   12
            Left            =   3360
            Picture         =   "PiloteExcel.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Choisir un formulaire"
            Top             =   240
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
            Height          =   510
            Index           =   11
            Left            =   4440
            Picture         =   "PiloteExcel.frx":0457
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Effectuer une Simulation"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   550
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   6
            Left            =   14400
            Picture         =   "PiloteExcel.frx":07D6
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer la personne"
            Top             =   6720
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
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
            Index           =   7
            Left            =   14400
            Picture         =   "PiloteExcel.frx":0C1D
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter une personne"
            Top             =   2160
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
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
            Index           =   4
            Left            =   14400
            Picture         =   "PiloteExcel.frx":1074
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter une règle"
            Top             =   1200
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   5
            Left            =   14400
            Picture         =   "PiloteExcel.frx":14CB
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer la règle"
            Top             =   1800
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
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
            Index           =   2
            Left            =   14400
            Picture         =   "PiloteExcel.frx":1912
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Choisir un formulaire"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   3
            Left            =   14400
            Picture         =   "PiloteExcel.frx":1D69
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le formulaire"
            Top             =   840
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
            Height          =   510
            Index           =   10
            Left            =   3720
            Picture         =   "PiloteExcel.frx":21B0
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Voir le fichier Excel"
            Top             =   600
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
            Index           =   9
            Left            =   5160
            Picture         =   "PiloteExcel.frx":252F
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Rafraichir"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   550
         End
         Begin MSFlexGridLib.MSFlexGrid grdFeuille 
            Height          =   1725
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   3043
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
         Begin MSFlexGridLib.MSFlexGrid grdForm 
            Height          =   885
            Left            =   7560
            TabIndex        =   11
            Top             =   240
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   1561
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
         Begin MSFlexGridLib.MSFlexGrid grdCond 
            Height          =   885
            Index           =   0
            Left            =   4440
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   10005
            _ExtentX        =   17648
            _ExtentY        =   1561
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
         Begin MSFlexGridLib.MSFlexGrid grdCell 
            Height          =   5655
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   2160
            Visible         =   0   'False
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   9975
            _Version        =   393216
            AllowUserResizing=   1
         End
         Begin ComctlLib.ImageList imglst 
            Left            =   3960
            Top             =   1440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   20
            ImageHeight     =   20
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   9
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":2923
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":2F05
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":34D3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":3899
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":3DEB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":63D9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":69BB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":6F9D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PiloteExcel.frx":741F
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   14895
      Begin ComctlLib.ProgressBar PgBarChp 
         Height          =   255
         Left            =   8760
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgBarFeuille 
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
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
         Index           =   8
         Left            =   1320
         Picture         =   "PiloteExcel.frx":7921
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "PiloteExcel.frx":7D2B
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Valider"
         Top             =   210
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
         Left            =   14160
         Picture         =   "PiloteExcel.frx":8184
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin ComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label LblSimulFeuille 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   11655
      End
   End
End
Attribute VB_Name = "PiloteExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bfaire_Click As Boolean
Private bfaire_RowColChange As Boolean

Private Const Ligne_Lib_Val = 1
Private Const Ligne_Lib = 2
Private Const Ligne_Val = 3
Private Const Colonne_Lib_Val = 4
Private Const Colonne_Lib = 5
Private Const Colonne_Val = 6

Private Const IMG_SOMME = 2
Private Const IMG_CHAMP = 5
Private Const IMG_BOULE = 8
Private Const IMG_BOULER = 9

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_FERMER = 1
Private Const CMD_PRM_FORM = 2
Private Const CMD_SUPPR_FORM = 3
Private Const CMD_PRM_COND = 4
Private Const CMD_SUPPR_COND = 5
Private Const CMD_ENREGISTRER = 8
Private Const CMD_RAFRAICHIR = 9
Private Const CMD_EXCEL_CELL = 10
Private Const CMD_SIMULATION = 11
Private Const CMD_AJOUT_FENETRE = 12

Private Const LAB_FEUILLE = 1
Private Const LAB_FORM = 0
Private Const LAB_REGLE = 2

Private g_nummodele As Long
Private g_numfiltre_encours As Integer
Private g_mode_saisie As Boolean
Private g_txt_avant As String
Private g_form_active As Boolean
Private g_numfeuille As Integer
Private g_CheminModele As String

' Tableau des Champs et conditions
Dim NomFichierParam As String

Private Type CELL
    CellFeuille As Integer
    CellX As Integer
    CellY As Integer
    CellTag As String
End Type
Dim tbl_cell() As CELL

Private Type SFICH_PARAM
    CmdType As String
    CmdFenNum As String
    CmdX As String
    CmdY As String
    CmdForNum As String
    CmdChpNum As String
    CmdCondition As String
    CmdTypeChp As String
    CmdMenFormeChp As String
End Type
Dim tbl_fich() As SFICH_PARAM

Private Type SCOND_PARAM
    CondNumFiltre As Integer
    CondString As String
    CondOper As String
    CondType As String
    CondFrancais As String
End Type
Dim tbl_cond() As SCOND_PARAM

Private Type SFEN_EXCEL
    FenNum As Integer
    FenNom As String
    FenLoad As Boolean
End Type
Dim tbl_fen() As SFEN_EXCEL

Private Type RDOF
    RDOF_num As Integer
    RDOF_rdoresultset As rdoResultset
    RDOF_sql As String
    RDOF_fornum As String
    RDOF_etat As String
End Type
Dim tbl_rdoF() As RDOF

Private Type RDOL
    RDOL_num As Integer
    RDOL_sql As String
    RDOL_sqlFrancais As String
    RDOL_fornum As String
End Type
Dim tbl_rdoL() As RDOL

Private Const Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const ColMax = 26
Private Const RowMax = 20


Public Function AppelFrm(ByVal v_nummodele As String) As Boolean

    g_nummodele = v_nummodele
    
    Me.Show 1
 
End Function

Private Function ajouter_form()

    Dim sql As String, sret As String, sfct As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    Dim LaDim As Integer
    Dim numfor As Integer
    Dim numfiltre As Integer
    Dim bajout As Boolean
    Dim selected As Boolean
    Dim trouve As Boolean
    
    Call CL_Init
    n = 0
    ' ceux qui y sont déja
    For i = 0 To grdForm.Rows - 1
        Call CL_AddLigne(grdForm.TextMatrix(i, 1) & " " & grdForm.TextMatrix(i, 2), grdForm.TextMatrix(i, 0), "", True)
        n = n + 1
    Next i
 
    sql = "select * from formulaire,filtreform where formulaire.for_num = filtreform.ff_fornum " _
            & " order by for_num,For_Lib"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    ' les autres
    While Not rs.EOF
        If rs("For_Lib").Value <> "" Then
            trouve = False
            For i = 0 To grdForm.Rows - 1
                If grdForm.TextMatrix(i, 0) = rs("FF_Num").Value Then
                    trouve = True
                    Exit For
                End If
            Next i
            If Not trouve Then
                Call CL_AddLigne(rs("For_Lib").Value & " " & rs("FF_Titre").Value, rs("FF_Num").Value, "", selected)
                n = n + 1
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Formulaires disponibles", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        ajouter_form = False
        Exit Function
    End If
    ' si ce formulaire n'est pas encore dans le tableau, l'y mettre
    For i = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(i).selected Then
            numfiltre = CL_liste.lignes(i).num
            If ajouter_form_grd(numfiltre, True, False) = P_OUI Then
                bajout = True
                ' charger les conditions de ce formulaire dans tblCond
                RecupereConditionF (numfiltre)
                ' charger le grid des regles pour ce formulaire
                Call ChargerGridRegle(numfiltre)
            End If
        End If
    Next i
    
    If bajout Then
        If grdForm.Rows > 0 Then
            cmd(CMD_SUPPR_FORM).Visible = True
        End If
        cmd(CMD_OK).Enabled = True
        ajouter_form = True
    Else
        ajouter_form = False
    End If
    
End Function

Private Function ChargerGridRegle(ByVal v_numfiltre As Integer)
    Dim i As Integer, lig As Integer, i_grdregle As Integer
    Dim j As Integer
    Dim sCnd As String
    Dim sOp As String
    Dim k As Integer
    Dim kdeja As Boolean
    Dim leUbound As Integer
    
    'Label(LAB_REGLE).Visible = True
    cmd(CMD_PRM_COND).Visible = True
    cmd(CMD_SUPPR_COND).Visible = True
    Load grdCond(v_numfiltre)
    For i = 0 To grdForm.Rows - 1
        i_grdregle = grdForm.TextMatrix(i, 0)
        grdCond(i_grdregle).Visible = False
    Next i
    grdCond(v_numfiltre).Visible = True
    grdCond(v_numfiltre).Cols = 4
    grdCond(v_numfiltre).Rows = 0
    grdCond(v_numfiltre).ColWidth(0) = 0    ' cachée
    grdCond(v_numfiltre).ColWidth(1) = 0    ' en kalitech
    grdCond(v_numfiltre).ColWidth(2) = grdCond(v_numfiltre).Width - 1000   ' en francais
    grdCond(v_numfiltre).ColWidth(3) = 1000     ' Type
    
    For i = 0 To UBound(tbl_cond())
        If tbl_cond(i).CondNumFiltre = v_numfiltre Then
            If i > 0 Then
                grdCond(v_numfiltre).AddItem "", lig
            Else
                grdCond(v_numfiltre).AddItem ""
                lig = grdCond(v_numfiltre).Rows - 1
            End If
            
            If tbl_cond(i).CondType = "CONDF" Then
                grdCond(v_numfiltre).TextMatrix(lig, 0) = tbl_cond(i).CondNumFiltre
                grdCond(v_numfiltre).TextMatrix(lig, 1) = tbl_cond(i).CondOper
                grdCond(v_numfiltre).TextMatrix(lig, 2) = tbl_cond(i).CondFrancais
                grdCond(v_numfiltre).TextMatrix(lig, 3) = "F"
                ' voir s'il y a des conditions locales (dans tbl_rdoL)
                On Error GoTo Err_Tab
                leUbound = UBound(tbl_rdoL())
                GoTo Suite_Tab
Err_Tab:
                Resume Apres_Tab
Suite_Tab:
                For j = 0 To UBound(tbl_rdoL())
                    If v_numfiltre = tbl_rdoL(j).RDOL_fornum Then
                        ' c'est le même filtre
                        ' on ajoute si y a pas déjà le meme num
                        kdeja = False
                        For k = 0 To grdCond(v_numfiltre).Rows - 1
                            If grdCond(v_numfiltre).TextMatrix(k, 3) = "L" Then
                                If grdCond(v_numfiltre).TextMatrix(k, 1) = tbl_rdoL(j).RDOL_num Then
                                    ' y est déjà
                                    kdeja = True
                                End If
                            End If
                        Next k
                        If Not kdeja Then
                            lig = lig + 1
                            grdCond(v_numfiltre).AddItem "", lig
                            grdCond(v_numfiltre).TextMatrix(lig, 0) = tbl_rdoL(j).RDOL_fornum
                            grdCond(v_numfiltre).TextMatrix(lig, 1) = tbl_rdoL(j).RDOL_num
                            grdCond(v_numfiltre).TextMatrix(lig, 2) = tbl_rdoL(j).RDOL_sqlFrancais
                            grdCond(v_numfiltre).TextMatrix(lig, 3) = "L"
                        End If
                    End If
                Next j
Apres_Tab:
                On Error GoTo 0
            End If
        End If
    Next i
End Function

Private Function ajouter_cond()
    Dim frm As Form
    Dim bcr As String
    Dim numcond As Integer, numfor As Integer
    Dim indTbCond As Integer, nb As Integer, lig As Integer
    Dim LaCond As String, LaCondFrancais As String
    Dim s As String
    Dim droite As String
    Dim gauche As String
    Dim pos As Integer
    Dim i As Integer
    Dim laS As String, laSF As String
    Dim encore As Boolean
    Dim itbl_rdoL As Integer, indRDO As Integer
    
    numcond = 0
    numfor = grdForm.TextMatrix(grdForm.RowSel, 0)
    Set frm = PrmFctJSChp
    bcr = PrmFctJSChp.AppelFrm("Ajout", 0, numcond, numfor)
    Set frm = Nothing
    If bcr <> "" Then
        laS = STR_GetChamp(bcr, "µ", 0)
        laSF = STR_GetChamp(bcr, "µ", 1)
        encore = True
        itbl_rdoL = -1
        While encore
            pos = InStr(laS, "OP:OU")
            If pos = 0 Then
                encore = False
                gauche = laS
            Else
                gauche = Mid(laS, 1, pos - 1)
                laS = Mid(laS, pos + 6)
            End If
            On Error GoTo Err_Tab_RDOL
            indRDO = UBound(tbl_rdoL()) + 1
            GoTo Suite_Tab_RDO
Err_Tab_RDOL:
            indRDO = 0
            Resume Suite_Tab_RDO
Suite_Tab_RDO:
            On Error GoTo 0
            ReDim Preserve tbl_rdoL(indRDO)
            tbl_rdoL(indRDO).RDOL_fornum = numfor
            If itbl_rdoL = -1 Then
                ' toutes les conditions d'une même requête locale portent le même RDOL_num
                itbl_rdoL = indRDO
            End If
            tbl_rdoL(indRDO).RDOL_num = itbl_rdoL
            tbl_rdoL(indRDO).RDOL_sql = Replace(gauche, "\", "|")
            tbl_rdoL(indRDO).RDOL_sqlFrancais = laSF
        Wend
        On Error GoTo 0
        
        cmd(CMD_ENREGISTRER).Visible = True
        indTbCond = UBound(tbl_cond(), 1) + 1
        ReDim Preserve tbl_cond(indTbCond) As SCOND_PARAM
        tbl_cond(indTbCond).CondNumFiltre = numfor
        tbl_cond(indTbCond).CondOper = laS
        tbl_cond(indTbCond).CondFrancais = laSF
        tbl_cond(indTbCond).CondType = "CONDL"  ' condition locale (ajoutée)
                
        grdCond(numfor).AddItem ""
        ' ajouter dans grdcond
        lig = grdCond(numfor).Rows - 1
        grdCond(numfor).AddItem "", lig
        grdCond(numfor).TextMatrix(lig, 0) = numfor
        grdCond(numfor).TextMatrix(lig, 1) = itbl_rdoL
        grdCond(numfor).TextMatrix(lig, 2) = laSF
        grdCond(numfor).TextMatrix(lig, 3) = "L"
    End If
End Function

Private Function ajouter_form_grd(ByVal v_numfiltre As Long, _
                                  ByVal v_nomi As Boolean, _
                                  ByVal v_mess_y_est As Boolean) As Integer

    Dim for_lib As String
    Dim ff_titre As String
    Dim lig As Integer, j As Integer
    
    If Odbc_RecupVal("select For_Lib, FF_Titre from Formulaire, FiltreForm where FF_Num=" & v_numfiltre & " and formulaire.for_num = filtreform.ff_fornum ", _
                     for_lib, ff_titre) = P_ERREUR Then
        ajouter_form_grd = P_ERREUR
        Exit Function
    End If
    lig = -1
    For j = 0 To grdForm.Rows - 1
        If grdForm.TextMatrix(j, 0) = v_numfiltre Then
            If v_mess_y_est Then
                Call MsgBox("'" & for_lib & "' est déjà dans la liste.", vbInformation + vbOKOnly, "")
            End If
            ajouter_form_grd = P_NON
            Exit Function
        End If
    Next j
    If lig >= 0 Then
        grdForm.AddItem "", lig
    Else
        grdForm.AddItem ""
        lig = grdForm.Rows - 1
    End If
    grdForm.TextMatrix(lig, 0) = v_numfiltre
    grdForm.TextMatrix(lig, 1) = for_lib
    grdForm.TextMatrix(lig, 2) = ff_titre
    grdForm.col = grdForm.Cols - 1
    grdForm.ColSel = grdForm.col
    grdForm.RowSel = grdForm.Rows - 1
    g_numfiltre_encours = v_numfiltre
    
    ajouter_form_grd = P_OUI
    
End Function

Private Function ajouter_feuille_grd(ByVal v_i As Integer)
    Dim lig As Integer
    
    lig = v_i - 1
    If lig >= 0 Then
        grdFeuille.AddItem "", lig
    Else
        grdFeuille.AddItem ""
        lig = grdFeuille.Rows - 1
    End If
    grdFeuille.TextMatrix(lig, 0) = tbl_fen(v_i).FenNum
    grdFeuille.TextMatrix(lig, 1) = tbl_fen(v_i).FenNom
    grdFeuille.TextMatrix(lig, 2) = ""
    grdFeuille.col = grdFeuille.Cols - 1
    grdFeuille.ColSel = grdFeuille.col
    grdFeuille.RowSel = grdFeuille.Rows - 1
    
    ajouter_feuille_grd = P_OUI
    
End Function

Private Sub ajouter_champ(ByVal v_idgrid As Integer, ByVal v_numfiltre As Long, ByVal v_rowsel As Long, ByVal v_colsel As Long)

    Dim sql As String, sret As String, sfct As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    Dim LaDim As Integer
    Dim numfor As Integer
    Dim numchp As Integer
    Dim numfiltre As Integer
    Dim bajout As Boolean
    Dim selected As Boolean
    Dim trouve As Boolean
    Dim chpnum As Integer
    Dim boolTexte As Boolean, boolDate As Boolean, boolEntier As Boolean
    Dim boolListe As Boolean
    Dim lib As String
    Dim Forme As String, MenForme As String
    Dim leX As Integer, leY As Integer
    Dim sX As Integer
    Dim strX As String
    
    chpnum = 0

    Call CL_Init
    n = 0
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = v_idgrid Then
                'Debug.Print Mid(Alpha, grdCell(index).ColSel, 1)
                If tbl_fich(i).CmdX = Mid(Alpha, v_colsel, 1) And tbl_fich(i).CmdY = v_rowsel Then
                    chpnum = tbl_fich(i).CmdChpNum
                    Exit For
                End If
            End If
        End If
    Next i
    ' afficher tous les champs de ce filtre
    sql = "select * from filtreform where ff_num =" & v_numfiltre
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    '
    numfor = rs("ff_fornum")
    sql = "select * from formetapechp where forec_fornum = " & numfor _
            & " order by forec_nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    ' les autres
    While Not rs.EOF
        If rs("Forec_nom").Value <> "" Then
            lib = rs("Forec_Nom").Value & "  (" & rs("Forec_Type").Value & " " & rs("Forec_FctValid").Value & ")"
            If chpnum = rs("Forec_Num") Then
                Call CL_AddLigne("  ===>" & lib, rs("Forec_Num").Value, "", selected)
            Else
                Call CL_AddLigne(lib, rs("Forec_Num").Value, "", selected)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Champs disponibles", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    cmd(CMD_ENREGISTRER).Visible = True
    ' Mettre le champ
    sql = "select * from formetapechp where forec_num = " & CL_liste.lignes(CL_liste.pointeur).num
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    '
    ' selon type de champs
    ' mettre dans tbl_fich
    Select Case rs("forec_type")
    Case "TEXT"
        boolTexte = True
        If InStr(rs("forec_fctvalid"), "DATE") > 0 Then
            boolDate = True
        ElseIf InStr(rs("forec_fctvalid"), "ENTIER") > 0 Then
            boolEntier = True
        End If
    Case "CHECK"
        boolListe = True
    Case "RADIO"
        boolListe = True
    Case "SELECT"
        boolListe = True
    End Select
    If boolTexte Or boolListe Then
        grdCell(v_idgrid).TextMatrix(v_rowsel, v_colsel) = rs("forec_nom") & " " & rs("forec_type")
        Dim newDim As Integer
        newDim = UBound(tbl_fich()) + 1
        ReDim Preserve tbl_fich(newDim) As SFICH_PARAM
        tbl_fich(newDim).CmdType = "CHP"
        tbl_fich(newDim).CmdForNum = v_numfiltre
        tbl_fich(newDim).CmdChpNum = rs("forec_num")
        tbl_fich(newDim).CmdFenNum = g_numfeuille
        tbl_fich(newDim).CmdX = Mid(Alpha, v_colsel, 1)
        tbl_fich(newDim).CmdY = v_rowsel
        ' choisir la forme à donner au résultat
        Forme = ChoisirForme(Forme, rs("forec_num"))
        If Forme = "" Then
            Forme = "Ligne_Lib_Val"
        End If
        tbl_fich(newDim).CmdMenFormeChp = Forme
        ' finaliser l'affichage
        'MsgBox Forme
        strX = tbl_fich(newDim).CmdX
        leX = InStr(Alpha, strX)
        leY = tbl_fich(newDim).CmdY
        ' Mettre le champ
        sql = "select * from formetapechp where forec_num = " & tbl_fich(newDim).CmdChpNum
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        grdCell(v_idgrid).TextMatrix(leX, leY) = "    " & rs("forec_nom")
        bfaire_RowColChange = False
        grdCell(v_idgrid).row = leY
        grdCell(v_idgrid).col = leX
        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_CHAMP).Picture
        ' Mise en forme
        Call MettreChamp("Mettre", leX, leY, Forme, rs("forec_num"), v_idgrid)
        bfaire_RowColChange = True
    End If
End Sub

Private Function ChoisirForme(v_Forme As String, v_forecNum As Integer)
    Dim sql As String
    Dim i As Integer
    Dim rs As rdoResultset
    Dim selected As Boolean
    
    Call CL_Init
    sql = "select * from formetapechp where forec_num = " & v_forecNum
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Champ " & rs("forec_nom"), "")
    rs.Close
            
    Call CL_AddLigne("En Ligne : Libellé", Ligne_Lib, "", selected)
    Call CL_AddLigne("En Ligne : Valeur", Ligne_Val, "", selected)
    Call CL_AddLigne("En Ligne : Libellé + Valeur", Ligne_Lib_Val, "", selected)
    Call CL_AddLigne("En Colonne : Libellé", Colonne_Lib, "", selected)
    Call CL_AddLigne("En Colonne : Valeur", Colonne_Val, "", selected)
    Call CL_AddLigne("En Colonne : Libellé + Valeur", Colonne_Lib_Val, "", selected)
    
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Function
    End If
    For i = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(i).selected Then
            Select Case CL_liste.lignes(i).num
            Case 1
                ChoisirForme = "Ligne_Lib_Val"
            Case 2
                ChoisirForme = "Ligne_Lib"
            Case 3
                ChoisirForme = "Ligne_Val"
            Case 4
                ChoisirForme = "Colonne_Lib_Val"
            Case 5
                ChoisirForme = "Colonne_Lib"
            Case 6
                ChoisirForme = "Colonne_Val"
            End Select
        End If
    Next i
End Function

Private Sub ajouter_modifier_regle(ByVal v_fajout As Boolean)

    Dim oper As String, valeur As String, Cond As String, libetape As String
    Dim sql As String
    Dim cr As Integer, numetape As Integer, lig As Integer
    Dim numchp As Long
    Dim frm As Form
    
    If v_fajout Then
        numchp = 0
        oper = ""
        valeur = ""
        numetape = -1
    Else
        If grdCond(lig).row < 0 Then
            Exit Sub
        End If
        numchp = grdCond(lig).TextMatrix(grdCond(lig).row, 1)
        oper = grdCond(lig).TextMatrix(grdCond(lig).row, 1)
        valeur = grdCond(lig).TextMatrix(grdCond(lig).row, 1)
        numetape = grdCond(lig).TextMatrix(grdCond(lig).row, 1)
    End If
    
    'Set Frm = ChoixFormEtapeSuiv
    'cr = ChoixFormEtapeSuiv.AppelFrm(g_numfor, g_numetape, numchp, oper, valeur, numetape)
    'Set Frm = Nothing
    If cr = 0 Then
        Exit Sub
    End If
    
    If v_fajout Then
        grdCond(lig).AddItem ""
        lig = grdCond(lig).Rows - 1
    Else
        lig = grdCond(lig).row
    End If
    If numchp = 0 Or oper = "" Then
        numchp = 0
        oper = ""
        valeur = ""
        Cond = ""
    Else
        Cond = build_cond(numchp, oper, valeur)
    End If
    If numetape = 0 Then
        libetape = "FIN DU PROCESSUS"
    Else
        sql = "select FORE_Libcourt from FormEtape where FORE_FORNum=" & 1 _
            & " and FORE_Numetape=" & numetape
        If Odbc_RecupVal(sql, libetape) = P_ERREUR Then
            Exit Sub
        End If
        libetape = "E" & numetape & " - " & libetape
    End If
    'grdCond(lig).TextMatrix(lig, GRDE_CHAMP) = numchp
    'grdCond(lig).TextMatrix(lig, GRDE_OPER) = oper
    'grdCond(lig).TextMatrix(lig, GRDE_VALEUR) = valeur
    'grdCond(lig).TextMatrix(lig, GRDE_ETAPE) = numetape
    'grdCond(lig).TextMatrix(lig, GRDE_COND) = Cond
    'grdCond(lig).TextMatrix(lig, GRDE_LIBETAPE) = libetape
    
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Function build_cond(ByVal v_numchp As Long, _
                            ByVal v_oper As String, _
                            ByVal v_valeur As String) As String
    
    Dim scond As String, sql As String, stype As String, s As String
    
    scond = ""
    
    sql = "select FOREC_Label, FOREC_Type from FormEtapeChp" _
        & " where FOREC_Num=" & v_numchp
    If Odbc_RecupVal(sql, scond, stype) = P_ERREUR Then
        Exit Function
    End If

    If v_oper = "=" Then
        scond = scond + " = "
    Else
        scond = scond + " <> "
    End If
    
    If stype = "RADIO" Or stype = "SELECT" Or stype = "CHECK" Then
        If v_valeur = "" Then
            s = "<Non renseigné>"
        Else
            sql = "select VC_Lib from ValChp" _
                & " where VC_Num=" & v_valeur
            If Odbc_RecupVal(sql, s) = P_ERREUR Then
                Exit Function
            End If
        End If
        scond = scond + s
    Else
        scond = scond + v_valeur
    End If

    build_cond = scond
    
End Function

Private Sub initialiser()
    Dim ret As Boolean
    Dim numfiltre As Integer
    Dim i As Integer
    
    cmd(CMD_OK).Enabled = False
    
    g_mode_saisie = False
    
    grdForm.Visible = True
    grdForm.Cols = 3
    grdForm.ColWidth(0) = 30
    grdForm.ColWidth(1) = grdForm.Width / 2
    grdForm.ColWidth(2) = grdForm.Width / 2
    grdForm.Rows = 0
    
    If grdForm.Rows = 0 Then
        cmd(CMD_SUPPR_FORM).Visible = False
    Else
        grdForm.row = 0
        grdForm.RowSel = 0
        grdForm.col = grdForm.Cols - 1
        grdForm.ColSel = grdForm.col
    End If
    ' Initialise le tableau des champs et condition
    grdFeuille.Visible = True
    grdFeuille.Cols = 3
    grdFeuille.ColWidth(0) = 0
    grdFeuille.ColWidth(1) = grdFeuille.Width - 300
    grdFeuille.ColWidth(2) = 200
    grdFeuille.Rows = 0
    
    ret = InitTabChp()
    
    'Label(LAB_FEUILLE).Visible = True
    'Label(LAB_FORM).Visible = True
    ' Initialise le grid pour une fenetre
    If g_numfeuille > 0 Then
        ' afficher le grid de la feuille
        Call InitGrdCell(g_numfeuille)
    End If
    
    ' Initialise le grid des formulaires
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CONDF" Then
            numfiltre = tbl_fich(i).CmdForNum
            If ajouter_form_grd(numfiltre, True, False) = P_OUI Then
                ' charger le grid des regles pour ce formulaire
                Call ChargerGridRegle(numfiltre)
            End If
        End If
    Next i
    
    g_mode_saisie = True

End Sub

Private Function InitTabChp()
    Dim fd As Integer
    Dim ligne As String
    Dim i As Integer
    Dim s As String
    
    ' Ouvrir le fichier texte
    ' lire les lignes et charger le tableau
    
    NomFichierParam = Mid(g_CheminModele, 1, Len(g_CheminModele) - 4)
    ' ouvrir le fichier Excel
    VerifOuvrir (NomFichierParam & ".xls")
    'OuvrirModele (NomFichierParam & ".xls")
    RemplirTabFenetre
    ' s'il n'existe pas, on le crée
    If Not FICH_FichierExiste(NomFichierParam & ".txt") Then
        If FICH_OuvrirFichier(NomFichierParam & ".txt", FICH_ECRITURE, fd) = P_ERREUR Then
        End If
        Close #fd
    End If
    
    If FICH_OuvrirFichier(NomFichierParam & ".txt", FICH_LECTURE, fd) = P_ERREUR Then
        InitTabChp = P_ERREUR
        Exit Function
    End If
    ' #Type = Cond (Condition sur le Formulaire) ou Chp
    ' #Type      |  Fornum      |   ChpNum    |   Num Fenetre      |          Condition
    ' COND       |81            |             |                    |condition
    ' CHP        |81            |65           |                    |

    i = 0
    While Not EOF(fd)
        Line Input #fd, ligne
        If left(ligne, 1) <> "#" Then
            ReDim Preserve tbl_fich(i) As SFICH_PARAM
            s = Trim(STR_GetChamp(ligne, "|", 0))
            tbl_fich(i).CmdType = s
            s = Trim(STR_GetChamp(ligne, "|", 1))
            tbl_fich(i).CmdForNum = s
            s = Trim(STR_GetChamp(ligne, "|", 2))
            tbl_fich(i).CmdChpNum = s
            s = Trim(STR_GetChamp(ligne, "|", 3))
            tbl_fich(i).CmdFenNum = s
            s = Trim(STR_GetChamp(ligne, "|", 4))
            tbl_fich(i).CmdX = s
            s = Trim(STR_GetChamp(ligne, "|", 5))
            tbl_fich(i).CmdY = val(s)
            s = Trim(STR_GetChamp(ligne, "|", 6))
            tbl_fich(i).CmdMenFormeChp = s
            s = Trim(STR_GetChamp(ligne, "|", 7))
            tbl_fich(i).CmdCondition = s
            i = i + 1
        End If
    Wend
    ' on doit récupérer la condition du filtre
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CONDF" Then
            RecupereConditionF tbl_fich(i).CmdForNum
        End If
        If tbl_fich(i).CmdType = "CONDL" Then
            RecupereConditionL i
        End If
    Next i
    Close #fd

End Function

Private Function RecupereNomChamp(ByVal v_chpnum As Integer, v_trait As String)
    Dim sql As String, ChpNom As String
    Dim rs As rdoResultset
    Dim ChpType As String
    Dim ValChp As String
    
    If Odbc_RecupVal("select forec_nom,forec_type,forec_valeurs_possibles from formetapechp where forec_num=" & v_chpnum, _
                     ChpNom, ChpType, ValChp) = P_ERREUR Then
        ChpNom = "???"
        Exit Function
    End If
    If v_trait = "nom" Then
        RecupereNomChamp = ChpNom
    ElseIf v_trait = "valchp" Then
        RecupereNomChamp = ValChp
    End If
End Function

Private Function RecupereConditionF(ByVal v_ForNum As Integer)
    Dim sql As String, strcond As String, unecond As String
    Dim indTbCond As Integer
    Dim rs As rdoResultset
    Dim nbcond As Integer, i As Integer
    Dim i_form As Integer, i_chp As Integer
    Dim LaCondFrancais As String, LeNomChp As String
    Dim LeChp As String, LaCond As String, leOP As String, LaCondF As String
    Dim indRDO As Integer
    Dim RDOlaRequete As String, RDOleOP As String
    Dim RDOetat As String, RDOfornum As Integer, RDOnum As Integer
    Dim RDOstrcond As String
    Dim LaValChp As String
    
    '61 | 53.1562|(X like {%V171;%{)|=|171;º
    '    102 | 53.1498|(to_date(X, {dd/mm/YYYY{) < {2009-01-01{)|<|01/01/2009º53.1525
    '|(X like {%V146;%{ or X like {%V147;%{)|=|146;147;º53.1517|(upper(translate(X,{Ó
    'ÔÚÞÛ¯¶¨¹þ{,{AAEEEIOUUC{)) like {*E*{)|contient|eº
    
    If Odbc_RecupVal("select ff_cond,ff_etat,ff_fornum,ff_num from filtreform where ff_num=" & v_ForNum, _
                     RDOstrcond, RDOetat, RDOfornum, RDOnum) = P_ERREUR Then
        Exit Function
    End If
    On Error GoTo Err_Tab
    indTbCond = UBound(tbl_cond()) + 1
    GoTo Suite_Tab
Err_Tab:
    indTbCond = 0
    Resume Suite_Tab
Suite_Tab:
    ' tableau des recordset
    On Error GoTo Err_Tab_RDO
    indRDO = UBound(tbl_rdoF()) + 1
    GoTo Suite_Tab_RDO
Err_Tab_RDO:
    indRDO = 0
    Resume Suite_Tab_RDO
Suite_Tab_RDO:
    On Error GoTo 0
    ReDim Preserve tbl_rdoF(indRDO)
    tbl_rdoF(indRDO).RDOF_etat = RDOetat
    tbl_rdoF(indRDO).RDOF_fornum = RDOfornum
    tbl_rdoF(indRDO).RDOF_num = RDOnum
    
    nbcond = STR_GetNbchamp(RDOstrcond, "§")
    RDOlaRequete = ""
    RDOleOP = ""
    For i = 0 To nbcond - 1
        unecond = STR_GetChamp(RDOstrcond, "§", i)
        LeChp = STR_GetChamp(unecond, "|", 0)
        i_form = STR_GetChamp(LeChp, ".", 0)
        i_chp = STR_GetChamp(LeChp, ".", 1)
        LaCond = STR_GetChamp(unecond, "|", 1)
        leOP = STR_GetChamp(unecond, "|", 2)
        LaCondF = STR_GetChamp(unecond, "|", 3)
        LaCond = Replace(LaCond, "{", "'")
        LaCond = Replace(LaCond, "}", "'")
        LeNomChp = RecupereNomChamp(i_chp, "nom")
        LaValChp = RecupereNomChamp(i_chp, "valchp")
        If LaValChp <> "" Then
            LaCondF = Transforme(LaCondF, LaValChp)
        End If
        LaCondFrancais = LeNomChp & " " & leOP & " " & LaCondF
        
        ReDim Preserve tbl_cond(indTbCond) As SCOND_PARAM
        tbl_cond(indTbCond).CondNumFiltre = v_ForNum
        LaCond = Replace(LaCond, "X", LeNomChp)
        tbl_cond(indTbCond).CondOper = LaCond
        tbl_cond(indTbCond).CondFrancais = LaCondFrancais
        tbl_cond(indTbCond).CondType = "CONDF"  ' condition du filtre
        
        RDOlaRequete = RDOlaRequete & RDOleOP & LaCond
        RDOleOP = " And "
        
        indTbCond = indTbCond + 1
    Next i
    tbl_rdoF(indRDO).RDOF_sql = RDOlaRequete

End Function

Private Function RecupereConditionL(ByVal v_i As Integer)
    Dim sql As String, strcond As String, unecond As String
    Dim indTbCond As Integer
    Dim rs As rdoResultset
    Dim nbcond As Integer, i As Integer
    Dim i_form As Integer, i_chp As Integer
    Dim LaCondFrancais As String, LeNomChp As String
    Dim s As String
    Dim LeChp As String, LaCond As String, leOP As String, LaCondF As String
    Dim indRDO As Integer
    Dim RDOlaRequete As String, RDOleOP As String
    Dim RDOetat As String, RDOfornum As Integer, RDOnum As Integer
    Dim RDOstrcond As String
    Dim LaValChp As String, v_ForNum As Integer
    Dim déjà As Boolean
    Dim laS As String, gauche As String
    Dim pos As Integer, encore As Boolean
    Dim laSF As String
    Dim itbl_rdoL As Integer
    
    itbl_rdoL = -1
    ' tableau des recordset
    ' découper la condition en plusieurs
    laS = STR_GetChamp(tbl_fich(v_i).CmdCondition, "µ", 0)
    laSF = STR_GetChamp(tbl_fich(v_i).CmdCondition, "µ", 1)
    encore = True
    'laS = tbl_fich(v_i).CmdMenFormeChp
    While encore
        pos = InStr(laS, "OP:OU")
        If pos = 0 Then
            encore = False
            gauche = laS
        Else
            gauche = Mid(laS, 1, pos - 1)
            laS = Mid(laS, pos + 6)
        End If
        On Error GoTo Err_Tab_RDOL
        indRDO = UBound(tbl_rdoL()) + 1
        GoTo Suite_Tab_RDO
Err_Tab_RDOL:
        indRDO = 0
        Resume Suite_Tab_RDO
Suite_Tab_RDO:
        On Error GoTo 0
        ReDim Preserve tbl_rdoL(indRDO)
        tbl_rdoL(indRDO).RDOL_fornum = tbl_fich(v_i).CmdForNum
        If itbl_rdoL = -1 Then
            itbl_rdoL = indRDO
        End If
        tbl_rdoL(indRDO).RDOL_num = itbl_rdoL
        tbl_rdoL(indRDO).RDOL_sql = Replace(gauche, "\", "|")
        tbl_rdoL(indRDO).RDOL_sqlFrancais = laSF
    Wend
    On Error GoTo 0
    
End Function

Private Function Transforme(LaCond As String, LaValChp As String)
    Dim sql As String, rs As rdoResultset
    Dim nb As Integer
    Dim i As Integer
    Dim s As String
    Dim LaCondOut As String
    Dim leOP As String
    
    nb = STR_GetNbchamp(LaCond, ";")
    LaCondOut = LaCond
    leOP = ""
    For i = 0 To nb - 1
        s = Trim(STR_GetChamp(LaCond, ";", i))
        sql = "select * from valchp where vc_num=" & s
        If Odbc_SelectV(sql, rs) <> P_ERREUR Then
            LaCondOut = Replace(LaCondOut, s & ";", leOP & rs("vc_lib"))
            leOP = " ou "
        End If
    Next i
    Transforme = LaCondOut
End Function

Private Sub RemplirTabFenetre()
    Dim i As Integer
    Dim prem As Integer
    prem = 0
    For i = 1 To exc_obj.ActiveWorkbook.Sheets.Count
        ReDim Preserve tbl_fen(i) As SFEN_EXCEL
        If prem = 0 Then
            g_numfeuille = i
            prem = 1
        End If
        tbl_fen(i).FenNum = i
        tbl_fen(i).FenNom = exc_obj.ActiveWorkbook.Sheets(i).Name
        tbl_fen(i).FenLoad = False
        ' ajouter dans le grid
        ajouter_feuille_grd (i)
    Next i
End Sub

Private Sub OuvrirModele(ByVal v_chemin As String)
    Dim encore As Boolean
    Dim retour As Integer
    Dim FichierIn As String, cmd As String
    Dim v_chemin_For As String, v_chemin_Fil As String, v_chemin_Excel As String
    ' Ouvrir le modele
'      Chemin_Parametrage = v_chemin_Excel _
'                         & "For" & tbParam(Ind_Numfor) & "\" _
'                         & "Fil" & tbParam(Ind_Special_Pere) & "\" _
'                         & "Valeurs\Param.xls"
      If FICH_FichierExiste(v_chemin) Then
         If Excel_Init(exc_obj) = P_OK Then
            Excel_OuvrirDoc v_chemin, "", Exc_wrk, False
            exc_obj.Visible = True
         End If
      End If
   Exit Sub
End Sub

Private Sub InitGrdCell(ByVal v_idgrid As Integer)
    Dim ColWidth As Integer
    Dim RowHeight As Integer
    Dim i As Integer, j As Integer
    Dim ij As Integer
    Dim sql As String
    Dim rs As rdoResultset
        
    ColWidth = 1000
    RowHeight = 300
    
    grdFeuille.col = 0
    grdFeuille.ColSel = grdFeuille.Cols - 1
    grdFeuille.row = v_idgrid - 1
    
    Load grdCell(v_idgrid)
    bfaire_Click = False
    
    g_numfeuille = v_idgrid
    tbl_fen(v_idgrid).FenLoad = True
    grdFeuille.TextMatrix((v_idgrid - 1), 2) = "X"
    grdCell(v_idgrid).Cols = ColMax + 1
    grdCell(v_idgrid).Rows = RowMax + 1
    
    grdCell(v_idgrid).ColWidth(0) = ColWidth
    grdCell(v_idgrid).RowHeight(0) = RowHeight
    
    grdCell(v_idgrid).TextMatrix(0, 0) = tbl_fen(v_idgrid).FenNom
    
    bfaire_RowColChange = False
    For i = 1 To ColMax
        grdCell(v_idgrid).ColWidth(i) = ColWidth
        grdCell(v_idgrid).col = i
        grdCell(v_idgrid).CellAlignment = flexAlignCenterCenter
        grdCell(v_idgrid).TextMatrix(0, i) = Mid(Alpha, i, 1)
    Next i
    For i = 1 To RowMax
        grdCell(v_idgrid).TextMatrix(i, 0) = i
    Next i
    For i = 1 To ColMax
        grdCell(v_idgrid).ColWidth(i) = ColWidth
    Next i
    
    On Error GoTo Err_Tab
    ij = UBound(tbl_cell())
    GoTo Suite_Tab
Err_Tab:
    ij = 0
    Resume Suite_Tab
Suite_Tab:
    On Error GoTo 0
    For i = 1 To RowMax
        For j = 1 To ColMax
            ReDim Preserve tbl_cell(ij)
            tbl_cell(ij).CellFeuille = v_idgrid
            tbl_cell(ij).CellTag = ""
            tbl_cell(ij).CellX = j
            tbl_cell(ij).CellY = i
            ij = ij + 1
        Next j
    Next i
    
    Dim leX As Integer, leY As Integer
    Dim MenForme As String
    Dim sX As String
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = v_idgrid Then
                sX = tbl_fich(i).CmdX
                leX = InStr(Alpha, sX)
                leY = tbl_fich(i).CmdY
                MenForme = tbl_fich(i).CmdMenFormeChp
                ' Mettre le champ
                sql = "select * from formetapechp where forec_num = " & tbl_fich(i).CmdChpNum
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Exit Sub
                End If
                grdCell(v_idgrid).TextMatrix(leX, leY) = "    " & rs("forec_nom")
                grdCell(v_idgrid).row = leY
                grdCell(v_idgrid).col = leX
                Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_CHAMP).Picture
                ' Mise en forme
                Call MettreChamp("Mettre", leX, leY, MenForme, rs("forec_num"), v_idgrid)
            End If
        End If
    Next i
    
    ' ajouter le contenu du tableau
    AjouterContenuTableau v_idgrid
    
    grdCell(v_idgrid).Visible = True
    bfaire_Click = True
    bfaire_RowColChange = True
End Sub

Private Function MettreChamp(v_trait As String, v_leX As Integer, v_leY As Integer, v_MenForme As String, v_forec_num As Integer, v_idgrid As Integer)
    Dim sql As String, rs As rdoResultset
    Dim forec_type As String
    Dim liste_num As Integer
    Dim col As Integer, row As Integer
    Dim bool_liste As Boolean
    Dim Le_Lib As String, La_Val As Integer
    Dim NomCellDest As String
    Dim exc_sheet As Excel.Worksheet
    Dim ChpNom As String
    Dim ChpLabel As String
    Dim anc_bfaire_RowColChange As Boolean
    
    sql = "select * from formetapechp where forec_num = " & v_forec_num
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    
    anc_bfaire_RowColChange = bfaire_RowColChange
    bfaire_RowColChange = False
    
    forec_type = rs("forec_type")
    liste_num = val(rs("forec_valeurs_possibles"))
    ChpNom = rs("forec_nom")
    ChpLabel = rs("forec_label")
    
    bool_liste = False
    Select Case forec_type
    Case "RADIO"
        bool_liste = True
    Case "CHECK"
        bool_liste = True
    Case "SELECT"
        bool_liste = True
    Case "TEXT"
    End Select
    
    If bool_liste Then
        sql = "select * from valchp where vc_lvcnum=" & liste_num & " order by vc_ordre"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        Set exc_sheet = Exc_wrk.Sheets(tbl_fen(v_idgrid).FenNom)
        
        If v_trait = "Vider" Then
            If (v_MenForme = "Colonne_Lib" Or v_MenForme = "Colonne_Val" Or v_MenForme = "Colonne_Lib_Val") Then
                'en Colonne : IMG_FL_BAS
                For row = v_leY To v_leY + rs.RowCount
                    If (v_MenForme = "Colonne_Lib") Then  'Colonne_Lib
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = ""
                    ElseIf (v_MenForme = "Colonne_Val") Then  'Colonne_Val
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = ""
                    ElseIf (v_MenForme = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = ""
                        NomCellDest = Mid(Alpha, (v_leX + 1), 1) & row
                        exc_sheet.Range(NomCellDest).Value = ""
                    End If
                    rs.MoveNext
                    If rs.EOF Then
                        Exit For
                    End If
                Next row
            End If
            If (v_MenForme = "Ligne_Lib" Or v_MenForme = "Ligne_Val" Or v_MenForme = "Ligne_Lib_Val") Then  'en Ligne
                For col = v_leX To v_leX + rs.RowCount
                    If (v_MenForme = "Ligne_Lib") Then  'Ligne_Lib
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = ""
                    ElseIf (v_MenForme = "Ligne_Val") Then  'Ligne_Val
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = ""
                    ElseIf (v_MenForme = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = ""
                        NomCellDest = Mid(Alpha, col, 1) & (v_leY + 1)
                        exc_sheet.Range(NomCellDest).Value = ""
                    End If
                    rs.MoveNext
                    If rs.EOF Then
                        Exit For
                    End If
                Next col
            End If
        Else
            If (v_MenForme = "Colonne_Lib" Or v_MenForme = "Colonne_Val" Or v_MenForme = "Colonne_Lib_Val") Then
                'en Colonne : IMG_FL_BAS
                For row = v_leY To v_leY + rs.RowCount
                    Le_Lib = rs("vc_lib")
                    La_Val = row
                    grdCell(v_idgrid).col = v_leX
                    grdCell(v_idgrid).row = row
                    If row = v_leY Then
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULER).Picture
                    Else
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULE).Picture
                    End If
                    If (v_MenForme = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
                        grdCell(v_idgrid).col = v_leX + 1
                        grdCell(v_idgrid).row = row
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULE).Picture
                    End If
                    If (v_MenForme = "Colonne_Lib") Then  'Colonne_Lib
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = Le_Lib
                        grdCell(v_idgrid).TextMatrix(row, v_leX) = "  " & Le_Lib
                        MetTag v_idgrid, v_leX, row, "Lib", ChpNom, ChpLabel, Le_Lib
                    ElseIf (v_MenForme = "Colonne_Val") Then  'Colonne_Val
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = La_Val
                        grdCell(v_idgrid).TextMatrix(row, v_leX) = "  " & La_Val
                        MetTag v_idgrid, v_leX, row, "Val", ChpNom, ChpLabel, Le_Lib
                    ElseIf (v_MenForme = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
                        NomCellDest = Mid(Alpha, v_leX, 1) & row
                        exc_sheet.Range(NomCellDest).Value = Le_Lib
                        grdCell(v_idgrid).TextMatrix(row, v_leX) = "  " & Le_Lib
                        MetTag v_idgrid, v_leX, row, "Lib", ChpNom, ChpLabel, Le_Lib
                        NomCellDest = Mid(Alpha, (v_leX + 1), 1) & row
                        exc_sheet.Range(NomCellDest).Value = La_Val
                        grdCell(v_idgrid).TextMatrix(row, v_leX + 1) = "  " & La_Val
                        MetTag v_idgrid, v_leX + 1, row, "Val", ChpNom, ChpLabel, Le_Lib
                    End If
                    rs.MoveNext
                    If rs.EOF Then
                        Exit For
                    End If
                Next row
            End If
            If (v_MenForme = "Ligne_Lib" Or v_MenForme = "Ligne_Val" Or v_MenForme = "Ligne_Lib_Val") Then  'en Ligne
                For col = v_leX To v_leX + rs.RowCount
                    Le_Lib = rs("vc_lib")
                    La_Val = col
                    grdCell(v_idgrid).col = col
                    grdCell(v_idgrid).row = v_leY
                    If col = v_leX Then
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULER).Picture
                    Else
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULE).Picture
                    End If
                    If (v_MenForme = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
                        grdCell(v_idgrid).col = col
                        grdCell(v_idgrid).row = v_leY + 1
                        Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULE).Picture
                    End If
                    If (v_MenForme = "Ligne_Lib") Then  'Ligne_Lib
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = Le_Lib
                        grdCell(v_idgrid).TextMatrix(v_leY, col) = "  " & Le_Lib
                        MetTag v_idgrid, col, v_leY, "Lib", ChpNom, ChpLabel, Le_Lib
                    ElseIf (v_MenForme = "Ligne_Val") Then  'Ligne_Val
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = La_Val
                        grdCell(v_idgrid).TextMatrix(v_leY, col) = "  " & La_Val
                        MetTag v_idgrid, col, v_leY, "Val", ChpNom, ChpLabel, Le_Lib
                    ElseIf (v_MenForme = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
                        NomCellDest = Mid(Alpha, col, 1) & v_leY
                        exc_sheet.Range(NomCellDest).Value = Le_Lib
                        grdCell(v_idgrid).TextMatrix(v_leY, col) = "  " & Le_Lib
                        MetTag v_idgrid, col, v_leY, "Lib", ChpNom, ChpLabel, Le_Lib
                        NomCellDest = Mid(Alpha, col, 1) & (v_leY + 1)
                        exc_sheet.Range(NomCellDest).Value = La_Val
                        grdCell(v_idgrid).TextMatrix(v_leY + 1, col) = "  " & La_Val
                        MetTag v_idgrid, col, v_leY + 1, "Val", ChpNom, ChpLabel, Le_Lib
                    End If
                    rs.MoveNext
                    If rs.EOF Then
                        Exit For
                    End If
                Next col
            End If
        End If
    End If
    bfaire_RowColChange = anc_bfaire_RowColChange

End Function

Private Sub SimulMettreChamp(ByRef v_leX As Integer, ByRef v_leY As Integer, v_MenForme As String, v_libelle As String, v_valeur As String, v_idgrid As Integer, v_bool_liste As Boolean)
    Dim NomCellDest As String
    Dim exc_sheet As Excel.Worksheet
    Dim anc_bfaire_RowColChange As Boolean
    
    anc_bfaire_RowColChange = bfaire_RowColChange
    bfaire_RowColChange = False
    
    Set exc_sheet = Exc_wrk.Sheets(tbl_fen(v_idgrid).FenNom)
    If (v_MenForme = "Colonne_Lib" Or v_MenForme = "Colonne_Val" Or v_MenForme = "Colonne_Lib_Val") Then
        If (v_MenForme = "Colonne_Lib") Then  'Colonne_Lib
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_libelle
        ElseIf (v_MenForme = "Colonne_Val") Then  'Colonne_Val
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_valeur
        ElseIf (v_MenForme = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_libelle
            NomCellDest = Mid(Alpha, (v_leX + 1), 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_valeur
        End If
        v_leY = v_leY + 1
    End If
    If (v_MenForme = "Ligne_Lib" Or v_MenForme = "Ligne_Val" Or v_MenForme = "Ligne_Lib_Val") Then  'en Ligne
        If (v_MenForme = "Ligne_Lib") Then  'Ligne_Lib
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_libelle
        ElseIf (v_MenForme = "Ligne_Val") Then  'Ligne_Val
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_valeur
        ElseIf (v_MenForme = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
            NomCellDest = Mid(Alpha, v_leX, 1) & v_leY
            exc_sheet.Range(NomCellDest).Value = v_libelle
            NomCellDest = Mid(Alpha, v_leX, 1) & (v_leY + 1)
            exc_sheet.Range(NomCellDest).Value = v_valeur
        End If
        v_leX = v_leX + 1
    End If
    bfaire_RowColChange = anc_bfaire_RowColChange
End Sub

Private Sub MetTag(v_NumFeuille, v_X, v_Y, v_Lib, v_ChpNom, v_ChpLabel, v_Le_Lib)
    Dim ij As Integer
    
    For ij = 0 To UBound(tbl_cell())
        If tbl_cell(ij).CellFeuille = v_NumFeuille Then
            If tbl_cell(ij).CellX = v_X And tbl_cell(ij).CellY = v_Y Then
                If v_Lib = "Lib" Then
                    tbl_cell(ij).CellTag = "Champ " & v_ChpNom & " (" & v_ChpLabel & ")" & " item '" & v_Le_Lib & "'"
                    Exit For
                ElseIf v_Lib = "Val" Then
                    tbl_cell(ij).CellTag = "Valeur pour " & v_ChpNom & " (" & v_ChpLabel & ")" & " item '" & v_Le_Lib & "'"
                    Exit For
                End If
            End If
        End If
    Next ij
End Sub

Private Sub VerifOuvrir(v_CheminModele As String)
    Dim i As Integer
    
    ' vérifier si le fichier modèle est ouvert : si non l'ouvrir
    
    On Error GoTo Err_Excel
Test_Excel:
    If exc_obj.Windows.Count = 0 Then
    End If
    GoTo Suite_Excel
Err_Excel:
    'MsgBox Err & " " & Error$
    If Err = 462 Or Err = 91 Then
        If Excel_Init(exc_obj) = P_OK Then
        End If
    End If
    Resume Suite_Excel
Suite_Excel:
    On Error GoTo 0
    If exc_obj.Windows.Count = 0 Then
        'Il faut ré Ouvrir le fichier
        OuvrirModele (v_CheminModele)
    Else
        ' voir si c'est bien lui
        For i = 1 To exc_obj.Workbooks.Count
            If UCase(exc_obj.Workbooks(i).FullName) = UCase(v_CheminModele) Then
                exc_obj.Workbooks(i).Activate
                Exit Sub
            End If
        Next i
        OuvrirModele (v_CheminModele)
    End If

End Sub

Private Sub AjouterContenuTableau(v_idgrid As Integer)
    ' ajouter le contenu du tableau
    Dim Exc_wrk As Excel.Workbook
    Dim exc_sheet As Excel.Worksheet
    Dim i As Integer, j As Integer
    Dim NomCellDest As String
    Dim widthCell As Integer, coef As Integer
    
    PgBar.Visible = True
    PgBar.Max = ColMax * RowMax
    PgBar.Value = 0
    bfaire_RowColChange = False
    
    VerifOuvrir (g_CheminModele)
    
    'exc_obj.Visible = False
    Set Exc_wrk = exc_obj.ActiveWorkbook
    Set exc_sheet = Exc_wrk.Sheets(tbl_fen(v_idgrid).FenNom)
    exc_sheet.Activate
    exc_obj.Visible = False
    For i = 1 To RowMax
        For j = 1 To ColMax
            PgBar.Value = PgBar.Value + 1
            NomCellDest = Mid(Alpha, j, 1) & i
            grdCell(v_idgrid).row = i
            grdCell(v_idgrid).col = j
            If exc_sheet.Range(NomCellDest).HasFormula Then
                ' il y a une formule
                'Debug.Print exc_sheet.Range(NomCellDest).Formula
                Set grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_SOMME).Picture
                grdCell(v_idgrid).tag = exc_sheet.Range(NomCellDest).formula
            Else
                ' du texte
                If exc_sheet.Range(NomCellDest).Value <> "" Then
                    ' si champ de formulaire
                    If grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULE).Picture Or grdCell(v_idgrid).CellPicture = imglst.ListImages(IMG_BOULER).Picture Then
                        'MsgBox i
                        'grdCell(v_idgrid).TextMatrix(i, j) = exc_sheet.Range(NomCellDest).Value

                    Else
                        grdCell(v_idgrid).TextMatrix(i, j) = exc_sheet.Range(NomCellDest).Value
                    End If
                End If
                grdCell(v_idgrid).CellFontBold = exc_sheet.Range(NomCellDest).Font.Bold
                grdCell(v_idgrid).CellForeColor = exc_sheet.Range(NomCellDest).Font.Color
                grdCell(v_idgrid).CellFontName = exc_sheet.Range(NomCellDest).Font.Name
                grdCell(v_idgrid).CellFontItalic = exc_sheet.Range(NomCellDest).Font.Italic
                grdCell(v_idgrid).CellFontSize = exc_sheet.Range(NomCellDest).Font.Size
                grdCell(v_idgrid).CellFontName = exc_sheet.Range(NomCellDest).Font.Name
                grdCell(v_idgrid).CellBackColor = exc_sheet.Range(NomCellDest).Interior.Color
                If i = 1 Then
                    coef = 20
                    widthCell = exc_sheet.Range(NomCellDest).Width
                    grdCell(v_idgrid).ColWidth(j) = coef * widthCell
                End If
            End If
        Next j
    Next i
    exc_obj.Visible = True
    bfaire_RowColChange = True
    PgBar.Visible = False
End Sub
Private Function quitter() As Boolean

    Dim reponse As Integer
    
    reponse = MsgBox("Confirmez-vous l'abandon de la relecture ?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If reponse = vbNo Then
        quitter = False
        Exit Function
    End If
    On Error GoTo Err_Quit
    exc_obj.Application.Quit
        
    Set exc_obj = Nothing
    GoTo Suite_Quit
Err_Quit:
    Resume Suite_Quit:
Suite_Quit:
    Unload Me
    
    quitter = True
    
End Function

Private Sub supprimer_form()
    Dim i_grdregle As Integer
    
    i_grdregle = grdForm.TextMatrix(grdForm.row, 0)
    If grdForm.Rows = 1 Then
        grdForm.Rows = 0
        ' supprimer aussi le grid des regles
        Unload grdCond(i_grdregle)
        cmd(CMD_SUPPR_FORM).Visible = False
    Else
        grdForm.RemoveItem (grdForm.row)
        ' supprimer aussi le grid des regles
        Unload grdCond(i_grdregle)
        grdForm.row = 0
        Call grdForm_Click
    End If
    
    cmd(CMD_OK).Enabled = True
    
End Sub
Private Sub supprimer_cond()
    Dim lig As Integer
    Dim Index As Integer
    
    ' supprimer du grid et du tableau
    MsgBox "supprimer"
    Index = grdForm.TextMatrix(grdForm.row, 0)
    lig = grdCond(grdForm.TextMatrix(grdForm.row, 0)).row
    grdCond(Index).RemoveItem (lig)
    cmd(CMD_ENREGISTRER).Visible = True
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case CMD_PRM_FORM
        Call ajouter_form
    Case CMD_SUPPR_FORM
        Call supprimer_form
    Case CMD_PRM_COND
        Call ajouter_cond
    Case CMD_SUPPR_COND
        Call supprimer_cond
    Case CMD_OK
        'Call valider
    Case CMD_FERMER
        Call quitter
    Case CMD_RAFRAICHIR
        Call Rafraichir
    Case CMD_ENREGISTRER
        Call enregistrer
    Case CMD_EXCEL_CELL
        VerifOuvrir (g_CheminModele)
        exc_obj.Visible = True
    Case CMD_SIMULATION
        Call Simulation
    Case CMD_AJOUT_FENETRE
        Call Ajout_Fenetre
    End Select

End Sub

Private Sub Ajout_Fenetre()
    Dim Exc_wrk As Excel.Workbook
    Dim exc_sheet As Excel.Worksheet
    Dim exc_sheets As Sheets
    Dim NomFen As String, iNext As Integer
    Dim i As Integer
    
    MsgBox "ici"
    Set Exc_wrk = exc_obj.ActiveWorkbook
    iNext = Exc_wrk.Sheets.Count + 1
    NomFen = InputBox("Nom de la nouvelle feuille", "Ajouter une fenêtre", "Feuille" & iNext)
    If NomFen <> "" Then
        Exc_wrk.Sheets.Add after:=Exc_wrk.Sheets(Exc_wrk.Sheets.Count)
        MsgBox Exc_wrk.Sheets(Exc_wrk.Sheets.Count).Name
        Exc_wrk.Sheets(Exc_wrk.Sheets.Count).Name = NomFen
        MsgBox Exc_wrk.Sheets(Exc_wrk.Sheets.Count).Name
        
        i = iNext
        
        ReDim Preserve tbl_fen(i) As SFEN_EXCEL
        g_numfeuille = i
        tbl_fen(i).FenNum = i
        tbl_fen(i).FenNom = exc_obj.ActiveWorkbook.Sheets(i).Name
        tbl_fen(i).FenLoad = False
        grdFeuille.TextMatrix(i, 2) = ""
        ' ajouter dans le grid
        ajouter_feuille_grd (i)
    End If
End Sub

Private Sub Simulation()
    Dim i As Integer, j As Integer
    Dim numfiltre As Integer
    Dim bool_liste As Boolean
    Dim sql As String, ChpNom As String, ChpLabel As String
    Dim forec_num As Integer
    Dim leX As Integer, leY As Integer
    Dim strX As String
    Dim ff_num As Integer
    Dim MenForme As String
    Dim rs As rdoResultset
    Dim rsVal As rdoResultset
    Dim rsChp As rdoResultset
    Dim forec_type As String
    Dim liste_num As Integer
    Dim Cnd As String
    Dim sqlVal As String
    Dim rsResult As rdoResultset
    Dim f As Integer
    Dim exc_sheet As Excel.Worksheet
    Dim CndTot As String
    Dim CndOp As String
    Dim op As String
    Dim k As Integer
    Dim kk As Integer
    Dim maxK As Integer
    Dim leUbound As Integer
    
    ' Charger les requete SQL
    ' Ouvrir un nouveau fichier, copie du modèle
    ' g_CheminModele à copier dans le dossier p_CheminRapportType
    'fermer le fichier modele
    For i = 1 To exc_obj.Workbooks.Count
        'exc_obj.Workbooks(i).Activate
        If UCase(exc_obj.Workbooks(i).FullName) = UCase(NomFichierParam & ".xls") Then
            exc_obj.Workbooks(i).Close True
            Exit For
        End If
    Next i
    
    'fermer le fichier temp
    For i = 1 To exc_obj.Workbooks.Count
        'exc_obj.Workbooks(i).Activate
        If UCase(exc_obj.Workbooks(i).FullName) = UCase(p_CheminRapportType & "\Temp.xls") Then
            exc_obj.Workbooks(i).Close True
            Exit For
        End If
    Next i
    
    If FICH_FichierExiste(p_CheminRapportType & "\Temp.xls") Then
        FICH_EffacerFichier p_CheminRapportType & "\Temp.xls", False
    End If
    FICH_CopierFichier g_CheminModele, p_CheminRapportType & "\Temp.xls"
    
    OuvrirModele (p_CheminRapportType & "\Temp.xls")
        
    PgBarFeuille.Visible = True
    PgBarFeuille.Max = UBound(tbl_fen)
    LblSimulFeuille.Visible = True
    Set Exc_wrk = exc_obj.ActiveWorkbook
    For f = 1 To UBound(tbl_fen)
        ' se mettre sur la bonne feuille
        PgBarFeuille.Value = f
        LblSimulFeuille.Caption = tbl_fen(f).FenNom
        
        Set exc_sheet = Exc_wrk.Sheets(tbl_fen(f).FenNom)
        exc_sheet.Activate
    
        'Vider les emplacements
        PgBarChp.Value = 0
        PgBarFeuille.Value = 0
        PgBarChp.Visible = True
        PgBarChp.Max = ColMax * RowMax
        For i = 0 To UBound(tbl_fich())
            PgBarChp.Value = i
            If tbl_fich(i).CmdType = "CHP" Then
                If tbl_fich(i).CmdFenNum = f Then
                    ' emplacement du champ
                    strX = tbl_fich(i).CmdX
                    leX = InStr(Alpha, strX)
                    leY = tbl_fich(i).CmdY
                    ff_num = tbl_fich(i).CmdForNum
                    MenForme = tbl_fich(i).CmdMenFormeChp
                    ' Trouver le champ
                    sql = "select * from formetapechp where forec_num = " & tbl_fich(i).CmdChpNum
                    If Odbc_SelectV(sql, rsChp) = P_ERREUR Then
                        Exit Sub
                    End If
                    ' son type
                    forec_type = rsChp("forec_type")
                    liste_num = rsChp("forec_valeurs_possibles")
                    ChpNom = rsChp("forec_nom")
                    ChpLabel = rsChp("forec_label")
                    'forec_num = tbl_fich(i).CmdChpNum
                    rsChp.Close
                    bool_liste = False
                    If forec_type = "RADIO" Or forec_type = "CHECK" Or forec_type = "SELECT" Then
                        bool_liste = True
                    Else
                        bool_liste = False
                    End If
    
                    If bool_liste Then
                        For j = 0 To grdForm.Rows - 1
                            If grdForm.TextMatrix(j, 0) = ff_num Then
                                MettreChamp "Vider", leX, leY, MenForme, val(tbl_fich(i).CmdChpNum), f
                            End If
                        Next j
                    End If
                End If
            End If
        Next i
        
        'Effectuer le remplacement de chaque champ
        
        PgBarChp.Max = UBound(tbl_fich())
        For i = 0 To UBound(tbl_fich())
            PgBarChp.Value = i
            If tbl_fich(i).CmdType = "CHP" Then
                If tbl_fich(i).CmdFenNum = f Then
                    ' emplacement du champ
                    strX = tbl_fich(i).CmdX
                    leX = InStr(Alpha, strX)
                    leY = tbl_fich(i).CmdY
                    ff_num = tbl_fich(i).CmdForNum
                    MenForme = tbl_fich(i).CmdMenFormeChp
                    ' Trouver le champ
                    sql = "select * from formetapechp where forec_num = " & tbl_fich(i).CmdChpNum
                    If Odbc_SelectV(sql, rsChp) = P_ERREUR Then
                        Exit Sub
                    End If
                    ' son type
                    forec_type = rsChp("forec_type")
                    liste_num = rsChp("forec_valeurs_possibles")
                    ChpNom = rsChp("forec_nom")
                    ChpLabel = rsChp("forec_label")
                    'forec_num = tbl_fich(i).CmdChpNum
                    rsChp.Close
                    bool_liste = False
                    If forec_type = "RADIO" Or forec_type = "CHECK" Or forec_type = "SELECT" Then
                        bool_liste = True
                    Else
                        bool_liste = False
                    End If
    
                    If bool_liste Then
                        ' trouver la liste de valeurs
                        sql = "select * from valchp where vc_lvcnum=" & liste_num & " order by vc_ordre"
                        If Odbc_SelectV(sql, rsVal) = P_ERREUR Then
                            Exit Sub
                        End If
                        ' comptage SQL
                        For j = 0 To grdForm.Rows - 1
                            If grdForm.TextMatrix(j, 0) = ff_num Then
                                ' c'est le bon filtre pour ce champ
                                ' constituer la condition complète : tbl_rdoF(j).RDOF_sql and ( conditions locales)
                                On Error GoTo Err_Tab
                                leUbound = UBound(tbl_rdoL())
                                On Error GoTo 0
                                GoTo Suite_Tab
Err_Tab:
                                CndTot = tbl_rdoF(j).RDOF_sql
                                Resume Apres_Tab
Suite_Tab:
                                CndTot = tbl_rdoF(j).RDOF_sql
                                op = " And "
                                CndOp = ""
                                Cnd = ""
                                maxK = 0
                                For k = 0 To UBound(tbl_rdoL())
                                    If ff_num = tbl_rdoL(k).RDOL_fornum Then
                                        ' c'est le même filtre
                                        ' on cumule tous ceux du même num
                                        'Debug.Print tbl_rdoL(k).RDOL_num & " " & tbl_rdoL(k).RDOL_sql
                                        CndOp = ""
                                        Cnd = ""
                                        For kk = k To UBound(tbl_rdoL())
                                            If tbl_rdoL(kk).RDOL_num = tbl_rdoL(k).RDOL_num Then
                                                If tbl_rdoL(kk).RDOL_sql <> "" Then
                                                    Cnd = Cnd & CndOp & MenFormeCnd(tbl_rdoL(kk).RDOL_sql)
                                                    CndOp = " Or "
                                                End If
                                                maxK = kk
                                            End If
                                        Next kk
                                        k = kk + 1
                                        'Debug.Print Cnd
                                        'Debug.Print CndTot
                                        If Cnd <> "" Then
                                            CndTot = CndTot & op & "(" & Cnd & ")"
                                            op = " And "
                                        End If
                                        'Debug.Print CndTot
                                    End If
                                Next k
Apres_Tab:
                                sql = "select count(*) from donnees_" & tbl_rdoF(j).RDOF_fornum & " Where [CND] And (" & CndTot & ")"
                                ' une boucle pour chaque valeur
                                While Not rsVal.EOF
                                    Cnd = ChpNom & " like '%V" & rsVal("vc_num") & ";%'"
                                    ' lire la table donnees
                                    sqlVal = Replace(sql, "[CND]", Cnd)
                                    If Odbc_SelectV(sqlVal, rsResult) = P_ERREUR Then
                                        Exit Sub
                                    End If
                                    'MsgBox rsResult(0) & " " & sqlVal
                                    If rsResult(0) > 0 Then
                                        SimulMettreChamp leX, leY, MenForme, rsVal("vc_lib"), rsResult(0), f, bool_liste
                                    End If
                                    rsVal.MoveNext
                                Wend
                            End If
                        Next j
                        rsVal.Close
                    End If
                End If
            End If
        Next i
    Next f
    PgBarFeuille.Visible = False
    PgBarChp.Visible = False
    LblSimulFeuille.Visible = False

End Sub

Private Function MenFormeCnd(v_cnd As String)
    Dim nb As Integer
    Dim chp As String, oper As String, valeur As String
    Dim sret As String
    
    nb = STR_GetNbchamp(v_cnd, "|")
    chp = STR_GetChamp(v_cnd, "|", 0)
    oper = STR_GetChamp(v_cnd, "|", 1)
    valeur = STR_GetChamp(v_cnd, "|", 2)
    'Debug.Print chp & " "; oper & " " & valeur
    chp = STR_GetChamp(chp, ":", 1)
    oper = STR_GetChamp(oper, ":", 1)
    valeur = STR_GetChamp(valeur, ":", 1)
    'Debug.Print chp & " "; oper & " " & valeur
    sret = chp & " like '%V" & valeur & ";%'"
    'Debug.Print sret
    MenFormeCnd = sret
End Function

Private Sub enregistrer()
    Dim fp As Integer
    Dim i As Integer
    Dim NomFichierParam As String
    Dim ligne As String
    Dim rs As rdoResultset
    Dim strX As String, MenForme As String, sql As String
    Dim leX As Integer, leY As Integer
    Dim j As Integer
    Dim i_grdregle As Integer
    Dim numfor As Integer
    Dim laS As String, laSF As String
    Dim leOP As String, k As Integer
    Dim tbl() As String
    Dim ind_tbl As Integer
    Dim yestdéjà As Boolean
    Dim leUbound As Integer
    
    MsgBox "enregistrer"
    VerifOuvrir (g_CheminModele)
    'enlever les champs du modèle dans le tableau Excel (pour chaque feuille)
    For i = 1 To UBound(tbl_fen())
        Debug.Print tbl_fen(i).FenNom & " " & tbl_fen(i).FenNum
        For j = 0 To UBound(tbl_fich())
            If tbl_fich(j).CmdType = "CHP" Then
                If tbl_fich(j).CmdFenNum = tbl_fen(i).FenNum Then
                    strX = tbl_fich(j).CmdX
                    leX = InStr(Alpha, strX)
                    leY = tbl_fich(j).CmdY
                    MenForme = tbl_fich(j).CmdMenFormeChp
                    ' Mettre le champ
                    sql = "select * from formetapechp where forec_num = " & tbl_fich(j).CmdChpNum
                    If Odbc_SelectV(sql, rs) = P_ERREUR Then
                        Exit Sub
                    End If
                    ' Mise en forme
                    Call MettreChamp("Vider", leX, leY, MenForme, rs("forec_num"), 1)
                End If
            End If
        Next j
    Next i
    ' enregistrer le modèle
    ' enregistrer le fichier des paramètres (tbl_fich)
    NomFichierParam = Mid(g_CheminModele, 1, Len(g_CheminModele) - 4)
    FICH_EffacerFichier NomFichierParam & ".txt", False
    FICH_OuvrirFichier NomFichierParam & ".txt", FICH_ECRITURE, fp
    For i = 0 To UBound(tbl_fich())
        ligne = tbl_fich(i).CmdType & "|" & tbl_fich(i).CmdForNum & "|" & tbl_fich(i).CmdChpNum & "|"
        ligne = ligne & tbl_fich(i).CmdFenNum & "|" & tbl_fich(i).CmdX & "|" & tbl_fich(i).CmdY & "|" & tbl_fich(i).CmdMenFormeChp
                Debug.Print ligne
        If tbl_fich(i).CmdType <> "CONDL" Then
            Print #fp, ligne
        End If
    Next i
    
    ' enregistrer les conditions locales
MsgBox "locales"
    ind_tbl = 0
    For i = 0 To grdForm.Rows - 1
        numfor = grdForm.TextMatrix(i, 0)
        ligne = "CONDL|" & numfor & "||||||"
        For j = 0 To UBound(tbl_rdoL())
            If numfor = tbl_rdoL(j).RDOL_fornum Then
                ' c'est le même filtre
                ' on cumule tous ceux du même num
                ligne = "CONDL|" & numfor & "||||||"
                laS = ""
                laSF = ""
                leOP = ""
                For k = 0 To UBound(tbl_rdoL())
                    If tbl_rdoL(k).RDOL_num = tbl_rdoL(j).RDOL_num Then
                        laS = laS & leOP & tbl_rdoL(k).RDOL_sql
                        leOP = "OP:OU|"
                        laSF = tbl_rdoL(k).RDOL_sqlFrancais
                    End If
                Next k
            End If
            If ligne <> "" Then
                ligne = ligne & Replace(laS & "µ" & laSF, "|", "\")
                ' voir s'il n'y est pas deja
                yestdéjà = False
                On Error GoTo Err_Tab
                leUbound = UBound(tbl())
                GoTo Suite_Tab
Err_Tab:
                Resume Apres_Tab
Suite_Tab:
                On Error GoTo 0
                For k = 0 To UBound(tbl())
                    If ligne = tbl(k) Then
                        yestdéjà = True
                    End If
                Next k
Apres_Tab:
                If Not yestdéjà Then
                    ReDim Preserve tbl(ind_tbl)
                    tbl(ind_tbl) = ligne
                    ind_tbl = ind_tbl + 1
                    Debug.Print ligne
                    Print #fp, ligne
                End If
            End If
        Next j
    Next i
    
    Close #fp
    cmd(CMD_ENREGISTRER).Visible = False
End Sub

Private Sub Rafraichir()
    Dim i As Integer, j As Integer
    Dim anc_bfaire_RowColChange As Boolean
        
    'MsgBox "ouvrir"
    VerifOuvrir (g_CheminModele)
    PgBar.Value = 0
    PgBar.Visible = True
    PgBar.Max = ColMax * RowMax
    anc_bfaire_RowColChange = bfaire_RowColChange
    bfaire_RowColChange = False
    For i = 1 To RowMax
        For j = 1 To ColMax
            grdCell(g_numfeuille).TextMatrix(i, j) = ""
            grdCell(g_numfeuille).row = i
            grdCell(g_numfeuille).col = j
            Set grdCell(g_numfeuille).CellPicture = Nothing
            PgBar.Value = PgBar.Value + 1
        Next j
    Next i
    bfaire_RowColChange = anc_bfaire_RowColChange
    PgBar.Visible = False
    
    AjouterContenuTableau g_numfeuille

    Dim leX As Integer, leY As Integer
    Dim MenForme As String
    Dim sX As String, sql As String
    Dim rs As rdoResultset
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = g_numfeuille Then
                sX = tbl_fich(i).CmdX
                leX = InStr(Alpha, sX)
                leY = tbl_fich(i).CmdY
                MenForme = tbl_fich(i).CmdMenFormeChp
                ' Mettre le champ
                sql = "select * from formetapechp where forec_num = " & tbl_fich(i).CmdChpNum
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Exit Sub
                End If
                grdCell(g_numfeuille).TextMatrix(leX, leY) = "    " & rs("forec_nom")
                grdCell(g_numfeuille).row = leY
                grdCell(g_numfeuille).col = leX
                Set grdCell(g_numfeuille).CellPicture = imglst.ListImages(IMG_CHAMP).Picture
                ' Mise en forme
                Call MettreChamp("Mettre", leX, leY, MenForme, rs("forec_num"), g_numfeuille)
            End If
        End If
    Next i

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = CMD_FERMER Then
        g_mode_saisie = False
    End If
    
End Sub

Private Sub Form_Activate()

    If g_form_active Then
        Exit Sub
    End If
    
    g_form_active = True
    g_CheminModele = ChoixModele()
    Erase tbl_cond()
    If g_CheminModele = "" Then
    Else
        Call initialiser
    End If
End Sub

Private Function ChoixModele()
   
   'Dim v_drive As String, v_path As String
   Dim v_path As String
   Dim frm As Form
   Dim NomFich As String
   Dim sret As String, chemin As String
   Dim v_chemin_For As String, v_chemin_Fil As String
   ' Choisir un fichier résultat
   Set frm = Com_ChoixFichier
   chemin = p_CheminRapportType
   v_drive = "c:"
   v_path = p_CheminRapportType
Test_Path:
   If FICH_EstRepertoire(v_path, False) Then
      NomFich = Com_ChoixFichier.AppelFrm("Les fichiers de Résultat", v_drive, v_path, "*.xls", False)
      Set frm = Nothing
      ChoixModele = NomFich
   Else
      MkDir (v_path)
      GoTo Test_Path
   End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyO And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        'Call valider
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

Private Sub grdCell_Click(Index As Integer)
    Dim ret As Boolean
    Dim i As Integer
    Dim ij As Integer
    Dim LaCaption As String
    Dim stype As String
    Dim NomCellDest As String
    
    stype = ""

    VerifOuvrir (NomFichierParam & ".xls")
    ' voir si cet emplacement est déjà utilisé dans le tableau
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = Index Then
                'Debug.Print i & "=" & tbl_fich(i).CmdX & " " & Mid(Alpha, grdCell(index).ColSel, 1) & " : " & tbl_fich(i).CmdY & " " & grdCell(index).RowSel
                If tbl_fich(i).CmdX = Mid(Alpha, grdCell(Index).ColSel, 1) And tbl_fich(i).CmdY = grdCell(Index).RowSel Then
                    ' il y est déja
                    stype = "tableau"
                    Exit For
                End If
            End If
        End If
    Next i
    
    Me.LblHelp.Visible = False
    If grdCell(Index).CellPicture = imglst.ListImages(IMG_BOULE).Picture Or grdCell(Index).CellPicture = imglst.ListImages(IMG_BOULER).Picture Then
        For ij = 0 To UBound(tbl_cell())
            If tbl_cell(ij).CellFeuille = Index Then
                If tbl_cell(ij).CellX = grdCell(Index).ColSel And tbl_cell(ij).CellY = grdCell(Index).RowSel Then
                    LaCaption = "Paramétrage " & tbl_cell(ij).CellTag
                    Exit For
                End If
            End If
        Next ij
        Me.LblHelp.Visible = True
        Me.LblHelp.Caption = LaCaption
    Else
        NomCellDest = Mid(Alpha, grdCell(Index).ColSel, 1) & grdCell(Index).RowSel
        If exc_obj.Sheets(g_numfeuille).Range(NomCellDest).HasFormula Then
            stype = "formule"
        Else
            If exc_obj.Sheets(g_numfeuille).Range(NomCellDest).Value <> "" Then
                stype = "value"
            End If
        End If
    End If
    If stype = "formule" Then
        Me.LblHelp.Visible = True
        Me.LblHelp.Caption = "Excel Formule : " & exc_obj.Sheets(g_numfeuille).Range(NomCellDest).formula
    ElseIf stype = "value" Then
        Me.LblHelp.Visible = True
        Me.LblHelp.Caption = "Excel Texte : " & exc_obj.Sheets(g_numfeuille).Range(NomCellDest).Value
    End If
End Sub

Private Sub Priv_grdCellClick(v_index As Integer, v_col, v_row)
    MsgBox "ici"
    Dim ret As Boolean
    Dim i As Integer

    ' voir si cet emplacement est déjà utilisé dans le tableau
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = v_index Then
                If tbl_fich(i).CmdX = Mid(Alpha, grdCell(v_index).ColSel, 1) And tbl_fich(i).CmdY = grdCell(v_index).RowSel Then
                    i = i
                End If
            End If
        End If
    Next i
    If g_numfiltre_encours = 0 Then
        MsgBox "Vous devez choisir un formulaire"
        ret = ajouter_form()
        If ret Then
            Call ajouter_champ(v_index, g_numfiltre_encours, grdCell(v_index).RowSel, grdCell(v_index).ColSel)
        End If
    Else
        Call ajouter_champ(v_index, g_numfiltre_encours, grdCell(v_index).RowSel, grdCell(v_index).ColSel)
    End If
End Sub

Private Sub grdCell_DblClick(Index As Integer)
    Dim ret As Boolean
    Dim i As Integer
    Dim bdeja As Boolean
    Dim NomCellDest As String
    Dim formule As String, texte As String
    
    bdeja = False
    ' voir si cet emplacement est déjà utilisé dans le tableau
    For i = 0 To UBound(tbl_fich())
        If tbl_fich(i).CmdType = "CHP" Then
            If tbl_fich(i).CmdFenNum = Index Then
                'Debug.Print Mid(Alpha, grdCell(index).ColSel, 1)
                If tbl_fich(i).CmdX = Mid(Alpha, grdCell(Index).ColSel, 1) And tbl_fich(i).CmdY = grdCell(Index).RowSel Then
                    ' il y est déja
                    bdeja = True
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not bdeja Then
        If grdCell(Index).CellPicture = imglst.ListImages(IMG_BOULE).Picture Then
            MsgBox "cette valeur est liée à un champ de formulaire"
            Exit Sub
        End If
        ' si y a du texte ou une formule, on fait rien
        exc_obj.ActiveWorkbook.Sheets(Index).Activate
        NomCellDest = Mid(Alpha, grdCell(Index).col, 1) & grdCell(Index).row
        If exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).HasFormula Then
            formule = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).formula
            formule = InputBox("Votre formule", "Saisir une formule", formule)
            If formule = "" Then
                Exit Sub
            End If
            If Mid(formule, 1, 1) <> "=" Then
                texte = formule
                GoTo Met_Texte
            End If
Met_Formule:
            exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).formula = formule
            grdCell(Index).TextMatrix(grdCell(Index).row, grdCell(Index).col) = formule
            Exit Sub
        ElseIf exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value <> "" Then
            ' y a du texte
            texte = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value
            texte = InputBox("Votre texte", "Saisir un Texte", texte)
            If texte = "" Then
                Exit Sub
            End If
            If Mid(texte, 1, 1) = "=" Then
                formule = texte
                GoTo Met_Formule
            End If
Met_Texte:
            NomCellDest = Mid(Alpha, grdCell(Index).col, 1) & grdCell(Index).row
            exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value = texte
            grdCell(Index).TextMatrix(grdCell(Index).row, grdCell(Index).col) = texte
            Exit Sub
        End If
    End If
    
    If g_numfiltre_encours = 0 Then
        MsgBox "Vous devez choisir un formulaire"
        ret = ajouter_form()
        If ret Then
            Call ajouter_champ(Index, g_numfiltre_encours, grdCell(Index).RowSel, grdCell(Index).ColSel)
        End If
    Else
        Call ajouter_champ(Index, g_numfiltre_encours, grdCell(Index).RowSel, grdCell(Index).ColSel)
    End If

End Sub

Private Sub grdCell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim formula As String
    Dim NomCellDest As String, formule As String, texte As String
    Dim lettre As String
    
    If KeyCode = 16 Then ' shift
        Exit Sub
    End If
    If KeyCode = 187 Then ' =
        exc_obj.ActiveWorkbook.Sheets(Index).Activate
        NomCellDest = Mid(Alpha, grdCell(Index).col, 1) & grdCell(Index).row
        If exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).HasFormula Then
            formule = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).formula
        ElseIf exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value <> "" Then
            GoTo Met_Texte
        Else
            formule = "="
        End If
        formule = InputBox("Votre formule", "Saisir une formule", formule)
        If Mid(formule, 1, 1) <> "=" Then
            GoTo Met_Texte
        End If
        
        exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).formula = formule
        grdCell(Index).TextMatrix(grdCell(Index).row, grdCell(Index).col) = formule
    ElseIf KeyCode = 46 Then ' suppr
        exc_obj.ActiveWorkbook.Sheets(g_numfeuille).Activate
        NomCellDest = Mid(Alpha, grdCell(Index).col, 1) & grdCell(Index).row
        exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value = ""
        grdCell(Index).TextMatrix(grdCell(Index).row, grdCell(Index).col) = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value
    Else
        exc_obj.ActiveWorkbook.Sheets(Index).Activate
        NomCellDest = Mid(Alpha, grdCell(Index).col, 1) & grdCell(Index).row
        
        If exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value <> "" Then
            texte = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value
        End If
Met_Texte:
        If Shift = 0 Then
            lettre = LCase(Chr(KeyCode))
        Else
            lettre = UCase(Chr(KeyCode))
        End If
        exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value & lettre
        grdCell(Index).TextMatrix(grdCell(Index).row, grdCell(Index).col) = exc_obj.ActiveWorkbook.Sheets(Index).Range(NomCellDest).Value
    End If
End Sub

Private Sub grdCell_RowColChange(Index As Integer)
    If bfaire_RowColChange Then
        Call grdCell_Click(Index)
    End If
End Sub

Private Sub grdCond_Click(Index As Integer)
    If grdCond(Index).TextMatrix(grdCond(Index).row, 3) = "F" Then
        cmd(CMD_SUPPR_COND).Visible = False
    Else
        cmd(CMD_SUPPR_COND).Visible = True
    End If
End Sub

Private Sub grdFeuille_Click()
    Dim i As Integer
    
    g_numfeuille = grdFeuille.RowSel + 1
    ' les mettre tous en invisibles
    For i = 1 To UBound(tbl_fen())
        If tbl_fen(i).FenLoad Then
            grdCell(i).Visible = False
        End If
    Next i
    If Not tbl_fen(g_numfeuille).FenLoad Then
        ' on le charge
        Call InitGrdCell(g_numfeuille)
        grdCell(g_numfeuille).Visible = True
    End If
    grdCell(g_numfeuille).Visible = True
    VerifOuvrir (NomFichierParam & ".xls")
    exc_obj.ActiveWorkbook.Sheets(g_numfeuille).Activate
End Sub

Private Sub grdForm_Click()
    'Call charge_grdfiltre(grdForm.Row)
    Dim i As Integer
    Dim i_grdregle As Integer
    
    For i = 0 To grdForm.Rows - 1
        i_grdregle = grdForm.TextMatrix(i, 0)
        grdCond(i_grdregle).Visible = False
    Next i
    
    grdForm.ColSel = grdForm.col
    grdForm.RowSel = grdForm.row

    i_grdregle = grdForm.TextMatrix(grdForm.RowSel, 0)
    grdCond(i_grdregle).Visible = True
    g_numfiltre_encours = grdForm.TextMatrix(grdForm.RowSel, 0)
End Sub

Private Sub grdform_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call ajouter_form
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_form
    End If
    
End Sub

