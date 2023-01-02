VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Publier 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Génération et publication des résultats d'un rapport"
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
      Width           =   11745
      Begin VB.Frame FrmDocument 
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
         Height          =   7695
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   11295
         Begin VB.Frame FrmHTTPD 
            BackColor       =   &H00C0C0C0&
            Height          =   1815
            Left            =   1920
            TabIndex        =   25
            Top             =   4800
            Visible         =   0   'False
            Width           =   8175
            Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
               Height          =   255
               Left            =   2880
               TabIndex        =   26
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
               TabIndex        =   27
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
               TabIndex        =   30
               Top             =   240
               Width           =   7455
            End
            Begin VB.Label lblHTTPDTemps 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   360
               TabIndex        =   29
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label lblHTTPDTaille 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   360
               TabIndex        =   28
               Top             =   840
               Width           =   2295
            End
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Index           =   0
            Left            =   5400
            Picture         =   "Publier.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Accès à l'aide"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   435
         End
         Begin VB.Frame Frm1Doc 
            BackColor       =   &H00C0C0C0&
            Height          =   4815
            Left            =   480
            TabIndex        =   10
            Top             =   2880
            Visible         =   0   'False
            Width           =   10335
            Begin VB.CheckBox ChkHyperlien 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Inclure les hyperliens"
               Height          =   255
               Left            =   3600
               TabIndex        =   22
               Top             =   840
               Width           =   1935
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Voir dans KaliWeb"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   15
               Left            =   3840
               Picture         =   "Publier.frx":0359
               Style           =   1  'Graphical
               TabIndex        =   21
               TabStop         =   0   'False
               ToolTipText     =   "Voir les fichiers résultats"
               Top             =   2520
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Diffuser le rapport dans KaliDoc"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   12
               Left            =   6240
               Style           =   1  'Graphical
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "Diffuser ce rapport dans KaliDoc"
               Top             =   4320
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   3315
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Voir les fichiers résultats générés"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   9
               Left            =   240
               Picture         =   "Publier.frx":0A4C
               Style           =   1  'Graphical
               TabIndex        =   19
               TabStop         =   0   'False
               ToolTipText     =   "Voir les fichiers résultats"
               Top             =   3480
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   3315
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Générer tous les documents"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   8
               Left            =   240
               Picture         =   "Publier.frx":0EF1
               Style           =   1  'Graphical
               TabIndex        =   13
               TabStop         =   0   'False
               ToolTipText     =   "Lancer le calcul pour tous les  documents"
               Top             =   1680
               UseMaskColor    =   -1  'True
               Width           =   3315
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Générer le Fichier"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   5
               Left            =   240
               Picture         =   "Publier.frx":12BF
               Style           =   1  'Graphical
               TabIndex        =   12
               TabStop         =   0   'False
               ToolTipText     =   "Lancer le calcul pour ce document"
               Top             =   840
               UseMaskColor    =   -1  'True
               Width           =   3315
            End
            Begin VB.CommandButton cmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Publier ce rapport dans KaliDoc"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   11
               Left            =   240
               Picture         =   "Publier.frx":163E
               Style           =   1  'Graphical
               TabIndex        =   11
               TabStop         =   0   'False
               ToolTipText     =   "Permet d'envoyer ce rapport dans la GED KaliDoc"
               Top             =   2520
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   3315
            End
            Begin MSFlexGridLib.MSFlexGrid grdDest 
               Height          =   3405
               Index           =   0
               Left            =   5520
               TabIndex        =   14
               Top             =   840
               Visible         =   0   'False
               Width           =   4485
               _ExtentX        =   7911
               _ExtentY        =   6006
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
            Begin ComctlLib.ProgressBar PgBarChp 
               Height          =   255
               Left            =   1920
               TabIndex        =   15
               Top             =   4440
               Visible         =   0   'False
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin ComctlLib.ProgressBar PgBarGener 
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   4440
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin ComctlLib.ProgressBar PgBarFeuille 
               Height          =   255
               Left            =   6960
               TabIndex        =   17
               Top             =   4440
               Visible         =   0   'False
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin ComctlLib.ProgressBar PgBarDoc 
               Height          =   255
               Left            =   4800
               TabIndex        =   18
               Top             =   4440
               Visible         =   0   'False
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin VB.Label lblPublic 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tout Public"
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
               Left            =   3600
               TabIndex        =   24
               Top             =   1680
               Visible         =   0   'False
               Width           =   1815
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdFeuille 
            Height          =   1965
            Left            =   5880
            TabIndex        =   6
            Top             =   600
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   3466
            _Version        =   393216
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   16777215
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
            TabIndex        =   7
            Top             =   600
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   3466
            _Version        =   393216
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   16777215
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
         Begin ComctlLib.ImageList ImageListS 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   28
            ImageHeight     =   25
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   8
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":1BDB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":212D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":24F3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":2AAD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":3067
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":3621
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":3BEF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Publier.frx":41A9
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
            TabIndex        =   9
            Top             =   240
            Width           =   2055
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
            TabIndex        =   8
            Top             =   240
            Width           =   3255
         End
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
         TabIndex        =   4
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
         TabIndex        =   3
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
      Width           =   11745
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
         Picture         =   "Publier.frx":483B
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "Publier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index sur les objets cmd
Private Const CMD_AIDE = 0
Private Const CMD_FERMER = 1
Private Const CMD_ICONE_KALIDOC = 4
Private Const CMD_AJOUT_DEST = 7
Private Const CMD_SUPPR_DEST = 6
Private Const CMD_GENERER_UN = 5
Private Const CMD_GENERER_TOUS = 8
Private Const CMD_VOIR_RESULTATS = 9
Private Const CMD_CHOIX_DOSSIER = 10
Private Const CMD_CHOIX_NATURE = 13
Private Const CMD_CHOIX_MODELE = 14
Private Const CMD_VOIR_DANS_KALIWEB = 15

Private g_numfeuille As Integer
Private g_numModele As Long
Private g_numDocument As Integer
Private g_CheminModele As String
Private g_mode_saisie As Boolean
Private g_form_active As Boolean
Private g_DocParamDefaut As String

Private Const IMG_LOAD_HTML = 1
Private Const IMG_LOAD_EXCEL = 2
Private Const IMG_KALIDOC = 3
Private Const IMG_PAS_KALIDOC = 4
Private Const IMG_ETAT_0 = 5
Private Const IMG_ETAT_1 = 6
Private Const IMG_ETAT_2 = 7
Private Const IMG_PUBLI_FAITE = 8

Private Faire_Doc_Click As Boolean

' pour le grid documents
Private Const GRDDOC_ETAT = 0
Private Const GRDDOC_TITRE = 1
Private Const GRDDOC_AFEN = 2
Private Const GRDDOC_PUBLIER_KD = 3
Private Const GRDDOC_LSTFEN = 4
Private Const GRDDOC_LSTDEST = 5
Private Const GRDDOC_EXCEL = 6
Private Const GRDDOC_IMGKD = 7
Private Const GRDDOC_NUMNAT = 8
Private Const GRDDOC_HTML = 9
Private Const GRDDOC_ImgPublié_KaliDoc = 10
Private Const GRDDOC_PUBLIC = 11
Private Const GRDDOC_Créé = 12
Private Const GRDDOC_NUMDOC = 13
Private Const GRDDOC_NUMDOS = 14
Private Const GRDDOC_MODELE = 15
Private Const GRDDOC_Publié_KaliDoc = 16

' pour le grid destinataires
Private Const GRDDEST_NUM = 0
Private Const GRDDEST_LIB = 1

' pour le grid feuilles
Private Const GRDFEUIL_ETAT = 0
Private Const GRDFEUIL_NUM = 1
Private Const GRDFEUIL_TAG = 2
Private Const GRDFEUIL_ADOC = 3
Private Const GRDFEUIL_LIB = 4

Dim p_BoolFaireChkClick As Boolean
Dim p_BoolFaireDocumentClick As Boolean

' Tableau des fenetres
Private Type SFEN_EXCEL
    FenNum As Integer
    FenNom As String
    FenDest As String
End Type
Dim tbl_fen() As SFEN_EXCEL

Public Function AppelFrm(ByVal v_nummodele As Long) As String

    g_numModele = v_nummodele
        
    Show 1
    
End Function

Private Sub RemplirTabFenetre()
    Dim I As Integer
    Dim prem As Integer
    Dim lig As Integer
    
    If Exc_obj Is Nothing Then
        Exit Sub
    End If
    
    grdFeuille.Visible = True
    grdFeuille.Cols = 5
    grdFeuille.ColWidth(GRDFEUIL_ETAT) = 400
    grdFeuille.ColWidth(GRDFEUIL_NUM) = 0
    grdFeuille.ColWidth(GRDFEUIL_TAG) = 0
    grdFeuille.ColWidth(GRDFEUIL_ADOC) = 400
    grdFeuille.ColWidth(GRDFEUIL_LIB) = grdFeuille.Width - 1300
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
        grdFeuille.TextMatrix(lig, GRDFEUIL_ETAT) = ""
        grdFeuille.TextMatrix(lig, GRDFEUIL_NUM) = I
        grdFeuille.TextMatrix(lig, GRDFEUIL_ADOC) = ""
        grdFeuille.TextMatrix(lig, GRDFEUIL_LIB) = Exc_obj.ActiveWorkbook.Sheets(I).Name   'tbl_fen(v_i).FenNom
        grdFeuille.TextMatrix(lig, GRDFEUIL_TAG) = ""
    Next I
End Sub


Private Sub initialiser()

    Dim NomFichierParam As String, nommod_loc As String
    Dim ret As Integer
    
    ' Chargement du paramétrage des des documents
    Call ChargerParam
    
    ' Remplir le tableau des fenetres
    ' ouvrir le fichier Excel
    g_CheminModele = p_Chemin_Modeles_Serveur & "/RP_" & g_numModele & p_PointExtensionXls
    nommod_loc = p_chemin_appli & "\tmp\RP" & Format(Time, "hhmmss") & p_PointExtensionXls
    If KF_GetFichier(g_CheminModele, nommod_loc) = P_ERREUR Then
        Exit Sub
    End If
    Public_VerifOuvrir nommod_loc, False, False, p_tbl_FichExcelPublier
    
    RemplirTabFenetre
    
    g_numDocument = -1
    ' se mettre sur le premier document
    If grdDocument.Rows > 0 Then
        grdDocument.row = 0
        grdDocument.ColSel = GRDDOC_TITRE
        p_BoolFaireDocumentClick = False
        grddocument_click
    Else
        ' pas encore de document : proposer création
        ' Droit de publier ?
        If PiloteExcelBis.VoirSiDroit("PUBLIER", val(p_nummodele), val(p_NumUtil)) Then
            ret = MsgBox("Ce tableau de bord ne contient aucun document" & Chr(13) & Chr(10) & "Vous devez retourner dans le paramétrage du modèle", vbExclamation + vbOKOnly, "Nouveau document")
        Else
            ret = MsgBox("Ce tableau de bord ne contient aucun document" & Chr(13) & Chr(10) & "Contactez l'administrateur de ce Modèle", vbExclamation + vbOKOnly, "Nouveau document")
        End If
        p_boolRetournerAuParam = True
        Unload Me
        Exit Sub
    End If
    
    If grdDocument.Rows = 1 Then
        cmd(CMD_GENERER_TOUS).Visible = False
    End If
    
    Faire_Doc_Click = True
    
    p_BoolFaireChkClick = True
    
    g_mode_saisie = True

End Sub

Private Sub ChargerParam()
    
    Dim sql As String
    Dim lig As Integer
    Dim lnb As Long
    Dim rs As rdoResultset
    
    grdDocument.Rows = 0
    grdDocument.Cols = 17
    grdDocument.ColWidth(GRDDOC_ETAT) = 400
    grdDocument.ColWidth(GRDDOC_TITRE) = 2850
    grdDocument.ColWidth(GRDDOC_AFEN) = 300
    grdDocument.ColWidth(GRDDOC_PUBLIER_KD) = 0
    grdDocument.ColWidth(GRDDOC_IMGKD) = 300
    grdDocument.ColWidth(GRDDOC_LSTFEN) = 0
    grdDocument.ColWidth(GRDDOC_LSTDEST) = 0
    grdDocument.ColWidth(GRDDOC_EXCEL) = 300
    grdDocument.ColWidth(GRDDOC_HTML) = 300
    grdDocument.ColWidth(GRDDOC_NUMNAT) = 0
    grdDocument.ColWidth(GRDDOC_NUMDOS) = 0
    grdDocument.ColWidth(GRDDOC_PUBLIC) = 0
    grdDocument.ColWidth(GRDDOC_NUMDOC) = 0
    grdDocument.ColWidth(GRDDOC_Créé) = 0
    grdDocument.ColWidth(GRDDOC_MODELE) = 0
    grdDocument.ColWidth(GRDDOC_PUBLIER_KD) = 0
    grdDocument.ColWidth(GRDDOC_ImgPublié_KaliDoc) = 300
    
    sql = "select * from rp_document where rpd_rpnum=" & g_numModele
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
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
        grdDocument.TextMatrix(lig, GRDDOC_MODELE) = rs("rpd_modele").Value
        ' Y a t il des fichiers générés à publier
        sql = "select count(*) from rp_fichier where rpf_rpdnum=" & rs("rpd_num").Value
        Call Odbc_Count(sql, lnb)
        If lnb > 0 Then
            grdDocument.row = lig
            grdDocument.col = GRDDOC_EXCEL
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_EXCEL).Picture
            grdDocument.col = GRDDOC_HTML
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_HTML).Picture
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    cmd(CMD_GENERER_UN).Visible = False
    cmd(CMD_GENERER_TOUS).Visible = False
    Call evaluer_btn_voir_resultats("RPNUM", g_numModele)
    
End Sub


Private Function quitter() As Boolean

    Dim reponse As Integer
    Dim LaUbound As Integer
    Dim I As Integer
    Dim j As Integer
    
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
    
    If Not Exc_obj Is Nothing Then
        If Exc_obj.Workbooks.Count = 0 Then
            Exc_obj.Quit
            Set Exc_obj = Nothing
        End If
    End If
    Unload Me
    
    quitter = True
    
End Function


Private Sub cmd_Click(Index As Integer)
    
    Dim sql As String, rs As rdoResultset
    Dim strModele As String
    Dim iRow As Integer
    Dim retDos As Long
    Dim frm As Form
    Dim rp As String
    
    Select Case Index
    Case CMD_AIDE
        Call Appel_Aide
    Case CMD_VOIR_DANS_KALIWEB
        Call VoirDocument(grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC))
    Case CMD_VOIR_RESULTATS
        Set frm = VoirFichiers
        Call VoirFichiers.AppelFrm(g_numModele, "RES", rp)
        Set frm = Nothing
        Call evaluer_btn_voir_resultats("RPNUM", g_numModele)
    Case CMD_FERMER
        Call quitter
    Case CMD_GENERER_UN
        Call GenererResultat("Un")
    Case CMD_GENERER_TOUS
        Call GenererResultat("Tous")
    End Select
End Sub

Private Function VoirDocument(v_numdoc As String)

'ATTENTION : - il faut encoder l’url : remplacer ? par %F3 et & par %26
'            - il faut crypter le V_util pour éviter de demander à l’utilisateur de s’identifier.
    Dim url As String
    Dim V_url As String
    Dim util As String
    Dim cnd_sversconf As String
    
    V_url = "acces_doc.php%3FV_doc=" & v_numdoc & "-1"     '    & "%26V_numdon=0"
    
    util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
    
    
    If p_S_Vers_Conf <> "" Then
        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
    End If
    url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & V_url
    ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & url, vbMaximizedFocus

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
            cmd(CMD_CHOIX_NATURE).tag = v_numNature
            cmd(CMD_CHOIX_NATURE).Caption = "Nature"
        Else
            ChercherNomNature = ""
            cmd(CMD_CHOIX_NATURE).tag = ""
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

Private Sub evaluer_btn_voir_resultats(ByVal Trait As String, ByVal Num As Integer)

    Dim sql As String
    Dim lnb As Long
    Dim lnb1 As Long
    
    If Trait = "NUMDOC" Then
        sql = "select count(*) from rp_fichier where rpf_rpdnum=" & Num
    Else
        sql = "select count(*) from rp_fichier where rpf_rpnum=" & Num
    End If
    Call Odbc_Count(sql, lnb)
    If lnb > 0 Then
        cmd(CMD_VOIR_RESULTATS).Visible = True
        If Trait = "NUMDOC" Then
            sql = "select count(*) from rp_fichier where rpf_rpdnum=" & Num _
                & " and rpf_diff_faite='f'"
        Else
            sql = "select count(*) from rp_fichier where rpf_rpnum=" & Num _
                & " and rpf_diff_faite='f'"
        End If
        Call Odbc_Count(sql, lnb1)
        If lnb1 > 0 Then
            cmd(CMD_VOIR_RESULTATS).BackColor = vbRed
            cmd(CMD_VOIR_RESULTATS).Caption = lnb1 & " fichiers résultats sont à diffuser"
        Else
            cmd(CMD_VOIR_RESULTATS).BackColor = &HCCCCCC
            cmd(CMD_VOIR_RESULTATS).Caption = lnb & " fichiers résultats disponibles"
        End If
    Else
        cmd(CMD_VOIR_RESULTATS).Visible = False
    End If

End Sub

Private Sub GenererResultat(v_Trait As String)
    Dim I As Integer, NomFichier As String, lstFen As String
    Dim bok As Boolean
    Dim n As Integer, s As String
    Dim j As Integer, reponse As Integer
    Dim sql As String, rs As rdoResultset
    Dim z As Integer
    Dim iRow As Integer, iDoc As Integer
    Dim bFaire As Boolean
    Dim CheminTemp As String, liberr As String
    Dim cheminRep As String, FicTmp As String
    Dim FicIn As String, FicIn_simple As String, FicOut As String
    Dim exc_sheet As Excel.Worksheet
    Dim F As Integer, b_laisser As Boolean, nomBook As String, NomFen As String
    Dim ret As Integer
    Dim numTemp As String
    Dim FicOutHTML As String
    Dim lstdest As String
    Dim stitre As String
    Dim sTitreAjoutG As String, sTitreAjoutD As String
    Dim sTitreG As String, sTitreD As String, nomfich_tmp As String
    Dim x As Integer
    Dim iret As Integer
    Dim II As Integer
    Dim strDocChk As String
    Dim UboundDiff As Integer
    Dim fctnum As String, fctlibelle As String
    Dim srvnum As String, srvnom As String
    Dim lib As String
    Dim iCar2 As Integer
    Dim strFichG As String, strFichd As String
    Dim strModele As String
    Dim strCheminDuTmp As String
    Dim iSovGrdDocument As Integer, iSovGrdFeuille As Integer
    Dim sChemin As String, nomrep_serv As String
    Dim NumFichier As Long, lbid As Long, numdoc As Long
    Dim CheminLiens As String
    Dim titreDocument As String
    
    ' Charger les feuilles si pas fait
    Me.FrmHTTPD.Visible = True
    Me.lblMaj.Caption = "Traitement du document " & grdDocument.TextMatrix(iRow, 1)
    iSovGrdFeuille = grdFeuille.row
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.row = I
        grdFeuille.col = GRDDOC_ETAT
        Set grdFeuille.CellPicture = ImageListS.ListImages(IMG_ETAT_0).Picture
    Next I
    grdFeuille.row = iSovGrdFeuille
    
    sql = "select nextval('rpf_seq')"
    If Odbc_RecupVal(sql, p_numFichier_Liens) = P_ERREUR Then
        Exit Sub
    End If
    
    p_ModePublication = "Publier"
    p_dansExcel = True
    CheminTemp = p_Chemin_Modeles_Local & "\Temp_" & g_numModele & p_PointExtensionXls
    ' Vider Temp ?
    'If FICH_FichierExiste(CheminTemp) Then
        Call PiloteExcelBis.Vider_TEMP(True, True, True)
    'End If
    p_chemin_fichier_liens = p_Chemin_Modeles_Local & "\RP_Liens_" & p_numFichier_Liens & ".txt"

    ' nombre aleatoire mais unique
    p_numdoc_liens = Format(Date, "YYYYMMDD") & "_" & Format(Time(), "hhmmss")
    p_nomdocument_encours = cmd(CMD_GENERER_UN).Caption
    
    p_nummodele_encours = g_numModele
    p_numdoc_encours = cmd(CMD_GENERER_UN).tag
    
    p_demander_titre = True
    Call Public_VerifOuvrir(CheminTemp, False, False, p_tbl_FichExcelOuverts)
    p_dansGrid = False
    p_dansExcel = True
    p_bPlusDeQuestion = False
    Me.PgbarHTTPDTaille.max = UBound(tbl_fenExcel)
    Me.PgbarHTTPDTemps.max = 100
    For I = 1 To UBound(tbl_fenExcel)
        Me.PgbarHTTPDTaille = I
        p_Simul_IFen = I
        Me.lblHTTPDTaille.Caption = "Traitement Feuille " & I
        Me.lblHTTPDTemps.Caption = "Simulation Excel"
        Me.PgbarHTTPDTemps.Value = 1
        Call PiloteExcelBis.Excel_Simulation("Feuille", strCheminDuTmp, iret)
        If iret = P_OK Then
            Me.lblHTTPDTemps.Caption = "Liens pour Feuille " & I
            Call PiloteExcelBis.Remplir_XLS(p_Simul_IFen, 0)
            Exc_obj.Visible = False
            Me.PgbarHTTPDTemps.Value = 2
        End If
        'i = 9
    Next I
    Me.lblHTTPDTaille.Caption = I - 1 & " Feuilles traitées"
    Me.lblHTTPDTemps.Caption = ""
    If iret <> P_OK Then
        Exit Sub
    End If
    If strCheminDuTmp = "" Then
        Call MsgBox("La génération n'a pas pu être effectuée.", vbCritical + vbOKOnly, "")
        Exit Sub
    End If
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.row = I
        grdFeuille.col = GRDDOC_ETAT
        Set grdFeuille.CellPicture = Nothing
    Next I
    grdFeuille.row = iSovGrdFeuille

Test_Path:
    ' fermer le fichier temp
    For I = 1 To Exc_obj.Workbooks.Count
        'exc_obj.Workbooks(i).Activate
        strFichG = Replace(UCase(Exc_obj.Workbooks(I).FullName), "\", "$")
        strFichG = Replace(strFichG, "/", "$")
        strFichd = Replace(UCase(strCheminDuTmp), "\", "$")
        strFichd = Replace(strFichd, "/", "$")
        If strFichG = strFichd Then
        
        'If UCase(Exc_obj.Workbooks(i).FullName) = UCase(CheminTemp) Then
            Exc_obj.Workbooks(I).Close True
            Exit For
        End If
    Next I

    ' générer les fichiers (si v_trait = Tous) (sinon v_trait=Un, on ne fait que celui là)
    PgBarDoc.Visible = True
    PgBarDoc.Min = 0
    PgBarDoc.max = grdDocument.Rows
    PgBarFeuille.Min = 0
    PgBarFeuille.max = grdFeuille.Rows
    lblFeuille.Visible = True
    lblDoc.Visible = True
    bFaire = True
    Dim iCar As Integer
            
    If v_Trait = "Un" Then
        iRow = grdDocument.row
        I = iRow
        grdDocument.row = iRow
        grdDocument.col = GRDDOC_ETAT
        Set grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_0).Picture
        GoTo LabCeDocument
    Else
        iSovGrdDocument = grdDocument.row
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            grdDocument.col = GRDDOC_ETAT
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_0).Picture
        Next I
        grdDocument.row = iSovGrdDocument
    End If
    For I = 0 To grdDocument.Rows - 1
        grdDocument.row = I
        'Me.lblHTTPDTaille.Caption = "Traitement Document " & i + 1
LabCeDocument:
        grdDocument.col = GRDDOC_ETAT
        If grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_0).Picture Then
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_1).Picture
            grdDocument.col = GRDDOC_LSTFEN
        End If
        
lab_saisie_titre:
        Me.PgbarHTTPDTaille.max = grdDocument.Rows
        Me.PgbarHTTPDTaille.Value = I
        Me.PgbarHTTPDTemps.Value = 0
        titreDocument = grdDocument.TextMatrix(I, GRDDOC_TITRE)
        Me.lblMaj.Caption = "Finalisation Document " & titreDocument
        Me.lblHTTPDTemps.Caption = ""
        Me.PgbarHTTPDTaille.Value = 0
        Me.PgbarHTTPDTemps.Value = 0
        Me.lblHTTPDTaille.Caption = ""
        stitre = InputBox("Pour " & titreDocument & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Indiquez le titre du résultat", "Titre du résultat", titreDocument)
        If stitre = "" Then
            reponse = MsgBox("Si vous n'indiquez pas de titre à votre document, la génération sera annulée. Voulez-vous saisir un titre ?", vbYesNo)
            If reponse = vbNo Then
                Exit Sub
            End If
            GoTo lab_saisie_titre
        End If
        Me.lblMaj.Caption = "Finalisation Document " & stitre
        
        lblDoc.Caption = stitre
        PgBarDoc.Value = I + 1
    
        Me.lblHTTPDTaille.Caption = "Mise en forme"
        ' Fichier en local
        FicIn = p_Chemin_Modeles_Local & "\Temp" & Format(Time, "hhmmss") & p_PointExtensionXls
        If FICH_CopierFichier(strCheminDuTmp, FicIn) = P_ERREUR Then
            Exit Sub
        End If
        
        ' Voir si on doit faire un repertoire RP_
        nomrep_serv = p_Chemin_Résultats & "/RP_" & g_numModele
        If Not KF_EstRepertoire(nomrep_serv, False) Then
            If KF_CreerRepertoire(nomrep_serv) = P_ERREUR Then
                Exit Sub
            End If
        End If
        ' Voir si on doit faire un repertoire Doc_
        numdoc = grdDocument.TextMatrix(I, GRDDOC_NUMDOC)
        nomrep_serv = nomrep_serv & "/Doc_" & numdoc
        If Not KF_EstRepertoire(nomrep_serv, False) Then
            If KF_CreerRepertoire(nomrep_serv) = P_ERREUR Then
                Exit Sub
            End If
        End If
        
        FicOut = nomrep_serv & "/" & p_numFichier_Liens & p_PointExtensionXls
        ' voir s'il existe déjà (sauf si publié)
        iRow = grdDocument.row
         
        If KF_FichierExiste(FicOut) Then
            ret = MsgBox("Le fichier " & FicOut & Chr(13) & Chr(10) & "existe déjà ... Voulez vous le remplacer ?", vbQuestion + vbYesNo + vbDefaultButton1, "Nouveau résultat")
            If ret = vbNo Then
                GoTo Next_i
            End If
        End If
        
        grdDocument.RowHeight(I) = 400
        
        ' l'ouvrir
        Call Public_VerifOuvrir(FicIn, False, True, p_tbl_FichExcelPublier)
        '
        lstFen = grdDocument.TextMatrix(I, GRDDOC_LSTFEN)
        n = STR_GetNbchamp(lstFen, ";")
        PgBarFeuille.Visible = True
        For F = 0 To grdFeuille.Rows - 1
            PgBarFeuille.Value = F + 1
            lblFeuille.Caption = grdFeuille.TextMatrix(F, GRDFEUIL_LIB)
            b_laisser = False
            For j = 0 To n - 1
                s = STR_GetChamp(lstFen, ";", j)
                s = Replace(s, "F", "")
                If s = grdFeuille.TextMatrix(F, GRDFEUIL_NUM) Then
                    b_laisser = True
                    Exit For
                End If
            Next j
            If Not b_laisser Then
                ' la supprimer
                nomBook = Exc_obj.Workbooks(Mid$(FicIn, InStr(FicIn, "Temp"))).Name
                Set Exc_wrk = Exc_obj.Workbooks(nomBook)
                NomFen = grdFeuille.TextMatrix(F, GRDFEUIL_LIB)
                Exc_wrk.Sheets(NomFen).Select
                Exc_obj.DisplayAlerts = False
                Exc_wrk.Sheets(NomFen).Delete
            End If
        Next F
        ' est chargé
        grdDocument.col = GRDDOC_EXCEL
        Set grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_EXCEL).Picture
        
        Exc_wrk.Sheets(1).Activate
        On Error Resume Next
        Exc_obj.ActiveSheet.PageSetup.CenterHeader = stitre

        FicTmp = p_Chemin_Modeles_Local & "\Temp" & Format(Time, "hhmmss") & p_PointExtensionXls
        Exc_wrk.SaveAs FicTmp
        
        ' transformer en HTML
        Me.lblHTTPDTaille.Caption = "Conversion en HTML"
        sChemin = p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil
        If Not FICH_EstRepertoire(sChemin, False) Then
            Call FICH_CreerRepComp(sChemin, False, False)
        End If
        FicOutHTML = sChemin & "\" & p_numFichier_Liens & ".html"
        Exc_wrk.SaveAs FileName:=FicOutHTML, _
            FileFormat:=44, ReadOnlyRecommended:=False, CreateBackup:=False
        grdDocument.col = GRDDOC_HTML
        Set grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_HTML).Picture
        Call Exc_wrk.Close
        
        ' Transfère le .xls sur le serveur
        Me.lblHTTPDTaille.Caption = "Transfert vers le Serveur"
        If KF_PutFichier(FicOut, FicTmp) = P_ERREUR Then
            GoTo Next_i
        End If
        ' Transfère le fichier des liens sur le serveur
        If FICH_FichierExiste(p_chemin_fichier_liens) Then
            If KF_PutFichier(nomrep_serv & "/" & p_numFichier_Liens & ".txt", p_chemin_fichier_liens) = P_ERREUR Then
                GoTo Next_i
            End If
        End If
        ' Transfère le .html sur le serveur
        If HTTP_Appel_PutDos(p_chemin_appli & "\tmp\TransfertHTML" & p_NumUtil, _
                             nomrep_serv & "/", False, False, liberr) <> HTTP_OK Then
            MsgBox liberr
        End If
        ' Vider le dossier local de transfert
        Call FICH_EffacerRep(sChemin)
        Call FICH_EffacerFichier(FicTmp, False)
        
        ' Mise à jour de rp_fichier
        sql = "insert into rp_fichier (rpf_num, rpf_rpnum, rpf_rpdnum, rpf_titre, rpf_diff_faite)" _
            & " values (" & p_numFichier_Liens & "," & g_numModele & ", " & numdoc & ", " & Odbc_String(stitre) & ", 'f')"
        Call Odbc_Cnx.Execute(sql)
                        
        ' Document suivant
Next_i:
        Me.PgbarHTTPDTaille.Value = I + 1

        grdDocument.col = GRDDOC_ETAT
        If grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_1).Picture Then
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_ETAT_2).Picture
        End If
        
        If v_Trait = "Un" Then
            GoTo Lab_Apres_NextI
        End If
        
    Next I
Lab_Apres_NextI:
    ' seulement si à publier
'    grdDocument.col = ColGrdKaliDoc
'    If grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture Then
'        cmd(CMD_VOIR_DIFFUSION).Visible = True
'    Else
'        cmd(CMD_VOIR_DIFFUSION).Visible = False
'    End If
    lblFeuille.Visible = False
    lblDoc.Visible = False
    PgBarDoc.Visible = True
    PgBarFeuille.Visible = True
    If v_Trait = "Tous" Then
        ' rouvrir tout
        CheminTemp = p_Chemin_Modeles_Local & "\Temp" & numTemp & p_PointExtensionXls
    
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            NomFichier = grdDocument.TextMatrix(I, GRDDOC_TITRE)
            NomFichier = Replace(NomFichier, " ", "_")
            lblDoc.Caption = grdDocument.TextMatrix(I, GRDDOC_TITRE)
            PgBarDoc.Value = I + 1
        Next I
    Else
        cmd(CMD_GENERER_UN).Visible = False
    End If
    
    For I = 0 To grdDocument.Rows - 1
        grdDocument.row = I
        grdDocument.col = GRDDOC_ETAT
        Set grdDocument.CellPicture = Nothing
    Next I
    grdDocument.row = iSovGrdDocument
    
    Me.FrmHTTPD.Visible = False
    PgBarDoc.Visible = False
    PgBarFeuille.Visible = False
    
    If v_Trait = "Tous" Then
        MsgBox "La génération des fichiers est terminée", vbOKOnly + vbInformation
    Else
        MsgBox "La génération du fichier est terminée", vbOKOnly + vbInformation
    End If
    Me.ChkHyperlien.Visible = False
    
    Call evaluer_btn_voir_resultats("NUMDOC", grdDocument.TextMatrix(iRow, 1))
    
End Sub

Private Function RemplirDiffService(ByVal v_srv_num, ByRef r_UboundDiff As Integer, ByVal v_FicOut, ByVal v_nomfichier, ByVal v_numdoc)
    Dim sql As String, rs As rdoResultset
    Dim UboundDiff As Integer
    Dim srvnum As Integer, srvnom As String
    Dim lib As String
    
    If Odbc_RecupVal("select SRV_Num, SRV_Nom from service where SRV_Num=" & v_srv_num, _
                     srvnum, srvnom) <> P_ERREUR Then
                            
        lib = "Service : " & srvnom
        ' mettre les utilisateurs
        sql = "select * from utilisateur where U_spm like '%S" & v_srv_num & ";%'"
        'MsgBox sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Else
            While Not rs.EOF
                ReDim Preserve p_tbl_diff(r_UboundDiff)
                p_bool_tbl_diff = True
                p_tbl_diff(r_UboundDiff).CheminDoc = v_FicOut
                p_tbl_diff(r_UboundDiff).nomdoc = v_nomfichier
                p_tbl_diff(r_UboundDiff).NumDest = rs("U_num")
                p_tbl_diff(r_UboundDiff).Diffusé = False
                p_tbl_diff(r_UboundDiff).nomdest = rs("U_nom") & " " & rs("U_prenom") & " (" & lib & ")"
                p_tbl_diff(r_UboundDiff).numdoc = v_numdoc
                r_UboundDiff = r_UboundDiff + 1
                rs.MoveNext
            Wend
        End If
    End If
End Function

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

    If KeyCode = vbKeyEscape Then
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
    'Set p_HTTP_Form_Frame = PrmPublier
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If Not quitter() Then
            Cancel = True
        End If
    End If
    
End Sub


Private Sub grddocument_click()
    
    Dim nomfich_loc As String, nomfich_serv As String, s As String
    Dim lstFen As String, sql As String, lstdest As String, url As String
    Dim chemin_doc As String
    Dim n As Integer, I As Integer, j As Integer, iRow As Integer
    Dim laCol As Integer
    Dim lnb As Long, numfich As Long
    
    If Not g_form_active Then
        Exit Sub
    End If
    
    If p_BoolFaireDocumentClick Then
        Me.ChkHyperlien.Visible = True
    End If
    
    iRow = grdDocument.row
    laCol = grdDocument.ColSel
    If iRow < 0 Then
        ' pas encore de document : proposer création
        MsgBox "Vous devez d'abord créer un document"
        iRow = grdDocument.row
    End If
    
    If grdDocument.Rows > 0 And Faire_Doc_Click Then
        g_numDocument = iRow
        Frm1Doc.Visible = True
        
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            grdDocument.TextMatrix(I, GRDDOC_AFEN) = ""
            If I = iRow Then
                Call evaluer_btn_voir_resultats("NUMDOC", grdDocument.TextMatrix(I, GRDDOC_NUMDOC))
                For j = 0 To grdDocument.Cols - 1
                    grdDocument.col = j
                    If j = GRDDOC_AFEN Then grdDocument.TextMatrix(I, j) = ">>"
                    grdDocument.CellBackColor = grdDocument.BackColorFixed
                    grdDocument.CellFontBold = True
                Next j
            Else
                For j = 0 To grdDocument.Cols - 1
                    grdDocument.col = j
                    If j = GRDDOC_AFEN Then grdDocument.TextMatrix(I, j) = ""
                    grdDocument.CellBackColor = grdDocument.BackColorBkg
                    grdDocument.CellFontBold = False
                Next j
            End If
        Next I
        grdDocument.row = iRow
        
        cmd(CMD_GENERER_UN).Visible = True
        If grdDocument.Rows > 1 Then
            cmd(CMD_GENERER_TOUS).Visible = True
        Else
            cmd(CMD_GENERER_TOUS).Visible = False
        End If
        
        lstFen = grdDocument.TextMatrix(grdDocument.row, GRDDOC_LSTFEN)
        lstdest = grdDocument.TextMatrix(grdDocument.row, GRDDOC_LSTDEST)
        lblPublic.Visible = False
        ' mettre la ligne en surbrillance
        Faire_Doc_Click = False
        For I = 0 To grdDocument.Rows - 1
            grdDocument.row = I
            If I = iRow Then
                grdDocument.CellFontBold = True
            Else
                grdDocument.CellFontBold = False
            End If
        Next I
        ' mettre les fenetres en non surbrillance
        For I = 0 To grdFeuille.Rows - 1
            grdFeuille.row = I
            grdFeuille.col = GRDFEUIL_LIB
            grdFeuille.CellFontBold = False
            grdFeuille.TextMatrix(I, GRDFEUIL_ADOC) = ""
        Next I
        ' mettre les fenetres concernées en surbrillance
        If lstFen = "" Then
            For j = 0 To grdFeuille.Rows - 1
                grdFeuille.row = j
                grdFeuille.TextMatrix(j, GRDFEUIL_ADOC) = ""
            Next j
        Else
            n = STR_GetNbchamp(lstFen, ";")
            For I = 0 To n - 1
                s = STR_GetChamp(lstFen, ";", I)
                s = Replace(s, "F", "")
                For j = 0 To grdFeuille.Rows - 1
                    grdFeuille.row = j
                    grdFeuille.TextMatrix(j, GRDFEUIL_TAG) = " "
                    If grdFeuille.TextMatrix(j, GRDFEUIL_NUM) = s Then
                        grdFeuille.TextMatrix(j, GRDFEUIL_ADOC) = ">>"
                        grdFeuille.RowHeight(j) = 400
                        grdFeuille.col = GRDFEUIL_ADOC
                        grdFeuille.CellFontBold = True
                        grdFeuille.col = GRDFEUIL_LIB
                        grdFeuille.CellFontBold = True
                        Exit For
                    End If
                Next j
            Next I
        End If
        ' grdDest en invisible
        For I = 0 To grdDocument.Rows - 1
            On Error Resume Next
            grdDest(I).Visible = False
        Next I
        On Error GoTo 0
        ChkHyperlien.Visible = True
        ChkHyperlien.Value = 1
        cmd(CMD_GENERER_UN).Caption = "Générer '" & grdDocument.TextMatrix(iRow, GRDDOC_TITRE) & "'"
        cmd(CMD_GENERER_UN).tag = grdDocument.TextMatrix(iRow, GRDDOC_NUMDOC)
        grdDocument.row = iRow
        grdDocument.col = GRDDOC_IMGKD
        If CBool(grdDocument.TextMatrix(iRow, GRDDOC_PUBLIER_KD)) Then
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_KALIDOC).Picture
        Else
            Set grdDocument.CellPicture = ImageListS.ListImages(IMG_PAS_KALIDOC).Picture
        End If
            
            ' afficher les destinataires
            AfficheDest "D", iRow
            If p_bool_tbl_diff Then
                For I = 0 To UBound(p_tbl_diff())
                    If Replace(p_tbl_diff(I).nomdoc, "_", " ") = grdDocument.TextMatrix(iRow, GRDDOC_TITRE) Then
                        Exit For
                    End If
                Next I
            End If
        grdDocument.row = iRow
        
        Faire_Doc_Click = True
        
        grdDocument.col = laCol
        
        If grdDocument.ColSel = GRDDOC_EXCEL Then
            grdDocument.row = grdDocument.RowSel
            grdDocument.col = grdDocument.ColSel
            If grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_EXCEL).Picture Then
                ' Chargement du dernier fichier en local
                sql = "select max(rpf_num) from rp_fichier where rpf_rpdnum=" & grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC)
                Call Odbc_MinMax(sql, numfich)
                If numfich = 0 Then
                    Exit Sub
                End If
                nomfich_loc = p_chemin_appli & "\tmp\RP" & Format(Time, "hhmmss") & p_PointExtensionXls
                nomfich_serv = p_Chemin_Résultats & "/RP_" & g_numModele _
                            & "/Doc_" & grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC) _
                            & "/" & numfich & p_PointExtensionXls
                If KF_GetFichier(nomfich_serv, nomfich_loc) = P_ERREUR Then
                    Exit Sub
                End If
                Call Public_VerifOuvrir(nomfich_loc, True, True, p_tbl_FichExcelPublier)
                Exit Sub
            End If
        ElseIf grdDocument.ColSel = GRDDOC_HTML Then
            grdDocument.row = grdDocument.RowSel
            grdDocument.col = grdDocument.ColSel
            If grdDocument.CellPicture = ImageListS.ListImages(IMG_LOAD_HTML).Picture Then
                sql = "select max(rpf_num) from rp_fichier where rpf_rpdnum=" & grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC)
                Call Odbc_MinMax(sql, numfich)
                If numfich = 0 Then
                    Exit Sub
                End If
                nomfich_serv = p_Chemin_Résultats & "/RP_" & g_numModele _
                            & "/Doc_" & grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC) _
                            & "/" & numfich & ".html"
                If KF_FichierExiste(nomfich_serv) Then
' MODIF LN 06/04/14
                    url = p_HTTP_Résultats & "/RP_" & g_numModele _
                                & "/Doc_" & grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC) _
                                & "/" & numfich & ".html"
                    Call SYS_ExecShell("C:\Program Files\Internet Explorer\iexplore.exe " & url, True, True)
                    Exit Sub
                Else
                    MsgBox "Fichier " & nomfich_serv & " introuvable"
                End If
            End If
        End If
        grdDocument.row = g_numDocument
    End If
    
    If iRow >= 0 Then grdDocument.row = iRow
    
    ' voir si publié dans KaliDoc
    If grdDocument.TextMatrix(iRow, GRDDOC_Publié_KaliDoc) = "O" Then
        cmd(CMD_VOIR_DANS_KALIWEB).Visible = True
    Else
        cmd(CMD_VOIR_DANS_KALIWEB).Visible = False
    End If
    
    Call evaluer_btn_voir_resultats("NUMDOC", grdDocument.TextMatrix(grdDocument.row, GRDDOC_NUMDOC))
    
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
    
    g_numfeuille = grdFeuille.RowSel + 1
    For I = 0 To grdFeuille.Rows - 1
        grdFeuille.col = GRDFEUIL_LIB
        grdFeuille.row = I
        grdFeuille.CellFontBold = False
        grdFeuille.TextMatrix(I, GRDFEUIL_TAG) = " "
        If grdFeuille.row = g_numfeuille - 1 Then
            For j = 0 To grdFeuille.Cols
                grdFeuille.col = j
                If j = GRDFEUIL_ADOC Then grdFeuille.TextMatrix(I, j) = "<<"
                grdFeuille.CellBackColor = grdFeuille.BackColorFixed
                grdFeuille.CellFontBold = True
            Next j
        Else
            For j = 0 To grdFeuille.Cols
                grdFeuille.col = j
                If j = GRDFEUIL_ADOC Then grdFeuille.TextMatrix(I, j) = ""
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
            grdDocument.col = GRDDOC_TITRE
            grdDocument.TextMatrix(I, GRDDOC_AFEN) = "<<"
            grdDocument.CellFontBold = True
            grdDocument.col = GRDDOC_AFEN
            grdDocument.CellFontBold = True
        Else
            grdDocument.col = GRDDOC_TITRE
            grdDocument.TextMatrix(I, GRDDOC_AFEN) = ""
            grdDocument.CellFontBold = False
            grdDocument.col = GRDDOC_AFEN
            grdDocument.CellFontBold = False
        End If
    Next I
    grdDocument.row = AncDocRow
    Faire_Doc_Click = Anc_Faire_Doc_Click

    grdFeuille.row = g_numfeuille - 1
    grdFeuille.col = GRDFEUIL_NUM
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
    Dim spublic As String
    
    If v_Trait = "D" Then
        grdDocument.Visible = True
        ' les destinataires sont dans grdDocument de v_i
        spublic = STR_GetChamp(grdDocument.TextMatrix(v_i, GRDDOC_LSTDEST), "%", 1)
        If spublic = "" Then
            lblPublic.Visible = False
        Else
            lblPublic.Visible = STR_GetChamp(grdDocument.TextMatrix(v_i, GRDDOC_LSTDEST), "%", 1)
        End If
        lstdest = STR_GetChamp(grdDocument.TextMatrix(v_i, GRDDOC_LSTDEST), "%", 0)
        ' charger le grid
        On Error GoTo Err_Load
        Load grdDest(v_i)
        GoTo Suite_Load
Err_Load:
        Resume Suite_Load
Suite_Load:
        grdDest(v_i).Rows = 0
        grdDest(v_i).Cols = 2
        grdDest(v_i).ColWidth(GRDDEST_NUM) = 0    '1000
        grdDest(v_i).ColWidth(GRDDEST_LIB) = grdDest(v_i).Width - 100
        On Error GoTo 0
        grdDest(v_i).Visible = True
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
            grdDest(v_i).TextMatrix(lig, GRDDEST_NUM) = s
            grdDest(v_i).TextMatrix(lig, GRDDEST_LIB) = lib
        Next I
    End If
End Sub

