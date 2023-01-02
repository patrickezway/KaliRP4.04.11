VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ChoixDestinataire 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix des destinataires"
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
      Height          =   7485
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8265
      Begin VB.Frame frmfctspm 
         BackColor       =   &H00C0C0C0&
         Height          =   6525
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   8265
         Begin ComctlLib.TreeView tvSect 
            Height          =   1365
            Left            =   1440
            TabIndex        =   3
            Top             =   3360
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   2408
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
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   12
            Left            =   7620
            Picture         =   "ChoixDestinataire.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Liste des personnes associées à ce service"
            Top             =   3870
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   11
            Left            =   7650
            Picture         =   "ChoixDestinataire.frx":056D
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Liste des personnes associées à cette fonction"
            Top             =   2310
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   345
            Index           =   10
            Left            =   7680
            Picture         =   "ChoixDestinataire.frx":0ADA
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Liste des personnes associées à ce groupe"
            Top             =   690
            UseMaskColor    =   -1  'True
            Width           =   315
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
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":1047
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Accéder aux groupes"
            Top             =   210
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   6
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":149E
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le groupe"
            Top             =   1300
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
            Index           =   9
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":18E5
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter une personne"
            Top             =   5010
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   8
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer la personne"
            Top             =   6060
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   3
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":2183
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer la fonction"
            Top             =   2880
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
            Index           =   2
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":25CA
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Accéder aux fonctions"
            Top             =   1800
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
            Index           =   4
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":2A21
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Accéder aux services"
            Top             =   3350
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   5
            Left            =   7260
            Picture         =   "ChoixDestinataire.frx":2E78
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le service"
            Top             =   4440
            UseMaskColor    =   -1  'True
            Width           =   320
         End
         Begin MSFlexGridLib.MSFlexGrid grdFct 
            Height          =   1365
            Left            =   1440
            TabIndex        =   2
            Top             =   1800
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   2408
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
         Begin MSFlexGridLib.MSFlexGrid grdPers 
            Height          =   1365
            Left            =   1440
            TabIndex        =   4
            Top             =   5010
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   2408
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
         Begin MSFlexGridLib.MSFlexGrid grdGrp 
            Height          =   1365
            Left            =   1440
            TabIndex        =   1
            Top             =   210
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   2408
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
         Begin ComctlLib.ImageList ImageListS 
            Left            =   840
            Top             =   3600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   22
            ImageHeight     =   22
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   4
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "ChoixDestinataire.frx":32BF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "ChoixDestinataire.frx":3B91
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "ChoixDestinataire.frx":4463
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "ChoixDestinataire.frx":4D35
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Image img 
            Height          =   300
            Index           =   3
            Left            =   480
            Picture         =   "ChoixDestinataire.frx":5587
            Top             =   5280
            Width           =   300
         End
         Begin VB.Image img 
            Height          =   480
            Index           =   2
            Left            =   360
            Picture         =   "ChoixDestinataire.frx":59E6
            Top             =   3600
            Width           =   480
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   1
            Left            =   480
            Picture         =   "ChoixDestinataire.frx":5F6B
            Top             =   2160
            Width           =   300
         End
         Begin VB.Image img 
            Height          =   300
            Index           =   0
            Left            =   480
            Picture         =   "ChoixDestinataire.frx":63C5
            Top             =   600
            Width           =   300
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Groupes"
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
            TabIndex        =   21
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Personnes "
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
            Left            =   240
            TabIndex        =   18
            Top             =   5070
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Services"
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
            Left            =   240
            TabIndex        =   15
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fonctions"
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
            Left            =   240
            TabIndex        =   14
            Top             =   1830
            Width           =   1095
         End
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tout le personnel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   915
      Left            =   0
      TabIndex        =   5
      Top             =   7260
      Width           =   8265
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
         Left            =   7290
         Picture         =   "ChoixDestinataire.frx":6861
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
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
         Left            =   480
         Picture         =   "ChoixDestinataire.frx":6E1A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixDestinataire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_PRM_GRP = 7
Private Const CMD_SUPPR_GRP = 6
Private Const CMD_PRM_FCT = 2
Private Const CMD_SUPPR_FCT = 3
Private Const CMD_PRM_SRV = 4
Private Const CMD_SUPPR_SRV = 5
Private Const CMD_PRM_PERS = 9
Private Const CMD_SUPPR_PERS = 8
Private Const CMD_LOUPE_GRP = 10
Private Const CMD_LOUPE_FCT = 11
Private Const CMD_LOUPE_SRV = 12
Private Const CMD_FERMER = 1

Private Const CHK_TOUT = 0

Private Const IMGT_SERVICE_NOMI = 1
Private Const IMGT_POSTE_NOMI = 2
Private Const IMGT_SERVICE = 3
Private Const IMGT_POSTE = 4

Private Const GRDG_NUM = 0
Private Const GRDG_NOM = 1

Private Const GRDF_NUM = 0
Private Const GRDF_NOMI = 1
Private Const GRDF_NOM = 2

Private Const GRDP_NUM = 0
Private Const GRDP_NOMI = 1
Private Const GRDP_NOM = 2

Private g_bcr As Boolean
Private g_sfctspm As String
Private g_titre As String

'Private g_crfct_autor As Boolean

Private g_mode_saisie As Boolean
Private g_form_active As Boolean

Public Function AppelFrm(ByRef vr_sfctspm As String, _
                         ByVal v_titre As String) As Boolean

    g_sfctspm = vr_sfctspm
    g_titre = v_titre
    
    If g_titre = "" Then
    Else
        frm.Caption = g_titre
    End If
    
    ChoixDestinataire.Show 1
 
    AppelFrm = g_bcr
    If g_bcr Then
        If vr_sfctspm = g_sfctspm Then
            AppelFrm = False
        End If
    End If
    vr_sfctspm = g_sfctspm
    
End Function

Private Function afficher_GFSU() As Integer

    Dim s As String
    Dim n As Integer, i As Integer
    Dim numutil As Long, numgrp As Long, numfct As Long
    
    If g_sfctspm = "" Then
        afficher_GFSU = P_OK
        Exit Function
    End If
    If g_sfctspm = "0" Then
        chk(CHK_TOUT).Value = 1
        frmfctspm.Enabled = False
        afficher_GFSU = P_OK
        Exit Function
    End If
    
    n = STR_GetNbchamp(g_sfctspm, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(g_sfctspm, "|", i)
        Select Case left$(s, 1)
        Case "U"
            numutil = CLng(Mid$(s, 2))
            Call ajouter_pers_grd(numutil, True, False)
        Case "G"
            numgrp = CLng(Mid$(s, 2))
            Call ajouter_grp_grd(numgrp)
        Case "F"
            numfct = CLng(Mid$(s, 2))
            Call ajouter_fct_grd(numfct, True, False)
        Case "S"
            Call ajouter_spm_tv(s, True, False)
        End Select
    Next i
        
    afficher_GFSU = P_OK
    
End Function

Private Function ajouter_fct_grd(ByVal v_numfct As Long, _
                                  ByVal v_nomi As Boolean, _
                                  ByVal v_mess_y_est As Boolean) As Integer

    Dim libfct As String
    Dim lig As Integer, j As Integer
    
    If Odbc_RecupVal("select FT_Libelle from FctTrav where FT_Num=" & v_numfct, libfct) = P_ERREUR Then
        ajouter_fct_grd = P_ERREUR
        Exit Function
    End If
    lig = -1
    For j = 0 To grdFct.Rows - 1
        If grdFct.TextMatrix(j, GRDF_NUM) = v_numfct Then
            If v_mess_y_est Then
                Call MsgBox("'" & libfct & "' est déjà dans la liste.", vbInformation + vbOKOnly, "")
            End If
            ajouter_fct_grd = P_NON
            Exit Function
        End If
        If UCase(grdFct.TextMatrix(j, GRDF_NOM)) > UCase(libfct) Then
            lig = j
            Exit For
        End If
    Next j
    If lig >= 0 Then
        grdFct.AddItem "", lig
    Else
        grdFct.AddItem ""
        lig = grdFct.Rows - 1
    End If
    grdFct.TextMatrix(lig, GRDF_NUM) = v_numfct
    grdFct.TextMatrix(lig, GRDF_NOMI) = v_nomi
    grdFct.TextMatrix(lig, GRDF_NOM) = libfct
    If Not v_nomi Then
        grdFct.row = lig
        grdFct.col = GRDF_NOM
        grdFct.CellForeColor = P_GRIS_FONCE
    End If
    
    Call ajouter_pers_gfs("F" & v_numfct)
    
    ajouter_fct_grd = P_OUI
    
End Function

Private Sub ajouter_grp_grd(ByVal v_numgrp As Long)

    Dim nomgrp As String
    
    If Odbc_RecupVal("select GU_Nom from GroupeUtil where GU_Num=" & v_numgrp, _
                     nomgrp) = P_ERREUR Then
        Exit Sub
    End If
    grdGrp.AddItem v_numgrp & vbTab & nomgrp
   ' Call ajouter_fctspm_grp(v_numgrp, False)
    
End Sub


Private Sub ajouter_pers()

    Dim sret As String
    Dim bajout As Boolean
    Dim i As Integer, lig As Integer
    Dim numutil As Long
    Dim frm As Form
    
    p_siz_tblu = -1
    Set frm = ChoixUtilisateur
    sret = ChoixUtilisateur.AppelFrm("Liste des personnes", _
                                      "", _
                                     True, _
                                      False, _
                                     True, _
                                     True)
    Set frm = Nothing
    If sret = "" Then
        Exit Sub
    End If
    
    If p_siz_tblu_sel = -1 Then
        Exit Sub
    End If
    
    bajout = False
    For i = 0 To p_siz_tblu_sel
        numutil = p_tblu_sel(i)
        If ajouter_pers_grd(numutil, True, True) = P_OUI Then
            bajout = True
        End If
    Next i
    
    If bajout Then
        If grdPers.Rows > 0 Then
            cmd(CMD_SUPPR_PERS).Visible = True
        End If
        cmd(CMD_OK).Enabled = True
    End If
      
End Sub


Private Sub ajouter_pers_gfs(ByVal v_sgfs As String)

    Dim sql As String, spm As String, slst As String, s As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
Exit Sub

    Select Case left$(v_sgfs, 1)
    Case "G"
        If Odbc_Select("select GU_Lst from GroupeUtil" _
                            & " where GU_Num=" & Mid$(v_sgfs, 2), _
                          rs) = P_ERREUR Then
            Exit Sub
        End If
        slst = rs("GU_Lst").Value & ""
        rs.Close
        n = STR_GetNbchamp(slst, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(slst, "|", i)
            If left$(s, 1) = "U" Then
                Call ajouter_pers_grd(Mid$(s, 2), False, False)
            End If
        Next i
    Case "F"
        sql = "select U_Num from Utilisateur" _
            & " where U_FctTrav like '%" & v_sgfs & ";%'" _
            & " and U_Actif=true"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        While Not rs.EOF
            Call ajouter_pers_grd(rs("U_Num").Value, False, False)
            rs.MoveNext
        Wend
        rs.Close
    Case "S"
        spm = v_sgfs
        If InStr(spm, "S1;") = 0 Then
            If InStr(spm, "P1;") > 0 Then
                spm = STR_GetChamp(spm, ";", 0) & ";"
            ElseIf InStr(spm, "M1;") > 0 Then
                spm = STR_GetChamp(spm, ";", 0) & ";" & STR_GetChamp(spm, ";", 1) & ";"
            End If
        End If
        sql = "select U_Num from Utilisateur" _
            & " where U_SPM like '%" & spm & "%'" _
            & " and U_Actif=true"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        While Not rs.EOF
            Call ajouter_pers_grd(rs("U_Num").Value, False, False)
            rs.MoveNext
        Wend
        rs.Close
    End Select
    
End Sub

Private Function ajouter_pers_grd(ByVal v_numutil As Long, _
                                  ByVal v_nomi As Boolean, _
                                  ByVal v_mess_y_est As Boolean) As Integer

    Dim nomutil As String, prenom As String
    Dim actif As Boolean
    Dim lig As Integer, j As Integer
    
    If Odbc_RecupVal("select U_Prenom, U_Nom, U_Actif from Utilisateur where U_Num=" & v_numutil, _
                     prenom, nomutil, actif) = P_ERREUR Then
        ajouter_pers_grd = P_ERREUR
        Exit Function
    End If
    If Not actif Then
        ajouter_pers_grd = P_NON
        Exit Function
    End If
    If prenom <> "" Then
        nomutil = nomutil + " " + prenom
    End If
    lig = -1
    For j = 0 To grdPers.Rows - 1
        If grdPers.TextMatrix(j, GRDP_NUM) = v_numutil Then
            If v_mess_y_est Then
                Call MsgBox("'" & nomutil & "' est déjà dans la liste.", vbInformation + vbOKOnly, "")
            End If
            ajouter_pers_grd = P_NON
            Exit Function
        End If
        If UCase(grdPers.TextMatrix(j, GRDP_NOM)) > UCase(nomutil) Then
            lig = j
            Exit For
        End If
    Next j
    If lig >= 0 Then
        grdPers.AddItem "", lig
    Else
        grdPers.AddItem ""
        lig = grdPers.Rows - 1
    End If
    grdPers.TextMatrix(lig, GRDP_NUM) = v_numutil
    grdPers.TextMatrix(lig, GRDP_NOMI) = v_nomi
    grdPers.TextMatrix(lig, GRDP_NOM) = nomutil
    If Not v_nomi Then
        grdPers.row = lig
        grdPers.col = GRDP_NOM
        grdPers.CellForeColor = P_GRIS_FONCE
    End If
    
    ajouter_pers_grd = P_OUI
    
End Function

Private Function ajouter_spm_tv(ByVal v_spm As String, _
                                ByVal v_nomi As Boolean, _
                                ByVal v_mess_y_est As Boolean) As Integer

    Dim lib As String, s As String, sql As String
    Dim fajout As Boolean
    Dim img As Integer, nbch As Integer, i As Integer
    Dim num As Long
    Dim nd As Node
    
    fajout = False
    
    nbch = STR_GetNbchamp(v_spm, ";")
    For i = 1 To nbch
        s = STR_GetChamp(v_spm, ";", i - 1)
        If TV_NodeExiste(tvSect, s, nd) = P_NON Then
            num = Mid$(s, 2)
            If left$(s, 1) = "S" Then
                If P_RecupSrvNom(num, lib) = P_ERREUR Then
                    ajouter_spm_tv = P_ERREUR
                    Exit Function
                End If
                If v_nomi Then
                    img = IMGT_SERVICE_NOMI
                Else
                    img = IMGT_SERVICE
                End If
                If i = 1 Then
                    Set nd = tvSect.Nodes.Add(, tvwChild, s, lib, img, img)
                Else
                    Set nd = tvSect.Nodes.Add(nd, tvwChild, s, lib, img, img)
                End If
            Else
                If P_RecupPosteNomfct(num, lib) = P_ERREUR Then
                    ajouter_spm_tv = P_ERREUR
                    Exit Function
                End If
                If v_nomi Then
                    img = IMGT_POSTE_NOMI
                Else
                    img = IMGT_POSTE
                End If
                Set nd = tvSect.Nodes.Add(nd, tvwChild, s, lib, img, img)
            End If
            nd.Expanded = True
            nd.Sorted = True
            nd.tag = v_nomi
            fajout = True
        ElseIf Not v_nomi Then
            If left$(s, 1) = "S" Then
                img = IMGT_SERVICE
            Else
                img = IMGT_POSTE
            End If
            nd.image = img
            nd.SelectedImage = img
            nd.tag = False
        End If
    Next i
    
    If fajout Then
        Call ajouter_pers_gfs(v_spm)
        ajouter_spm_tv = P_OUI
    ElseIf v_mess_y_est Then
        s = STR_GetChamp(v_spm, ";", nbch - 1)
        num = Mid$(s, 2)
        If left$(s, 1) = "S" Then
            If P_RecupSrvNom(num, lib) = P_ERREUR Then
                ajouter_spm_tv = P_ERREUR
                Exit Function
            End If
            lib = "Service " & lib
        Else
            If P_RecupPosteNom(num, lib) = P_ERREUR Then
                ajouter_spm_tv = P_ERREUR
                Exit Function
            End If
            lib = "Poste " & lib
        End If
        Call MsgBox("'" & lib & "' est déjà dans la liste.", vbInformation + vbOKOnly, "")
        ajouter_spm_tv = P_NON
    End If
    
End Function

Private Sub ajouter_fctspm_grp(ByVal v_numgrp As Long, _
                               ByVal v_nomi As Boolean)

    Dim slst As String, s As String
    Dim n As Integer, i As Integer
    Dim numfct As Long, numutil As Long
    Dim rs As rdoResultset
    
    If Odbc_Select("select GU_Lst from GroupeUtil where GU_Num=" & v_numgrp, _
                     rs) = P_ERREUR Then
        Exit Sub
    End If
    slst = rs("GU_Lst").Value & ""
    rs.Close
    n = STR_GetNbchamp(slst, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(slst, "|", i)
        Select Case left$(s, 1)
        Case "F"
            numfct = CLng(Mid$(s, 2))
            Call ajouter_fct_grd(numfct, v_nomi, False)
        Case "S"
            Call ajouter_spm_tv(s, v_nomi, False)
        Case "U"
            numutil = CLng(Mid$(s, 2))
            Call ajouter_pers_grd(numutil, v_nomi, False)
        End Select
    Next i

End Sub

Private Function build_liste_pers(ByVal v_slst As String) As Integer

    Dim nomutil As String, sql As String, nom As String, prenom As String
    Dim s As String
    Dim n As Integer, i As Integer, nbitem As Integer
    Dim numutil As Long
    Dim rs As rdoResultset
    
    nbitem = 0
    
    n = STR_GetNbchamp(v_slst, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(v_slst, "|", i)
        Select Case left$(s, 1)
        Case "F"
            sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
                & " where U_FctTrav like '%" & s & ";%'" _
                & " and U_Actif=true"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                build_liste_pers = 0
                Exit Function
            End If
            While Not rs.EOF
                nomutil = rs("U_Nom").Value + " " + rs("U_Prenom").Value
                Call CL_AddLigne(nomutil, 0, "", False)
                nbitem = nbitem + 1
                rs.MoveNext
            Wend
            rs.Close
        Case "S", "P"
            sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
                & " where U_SPM like '%" & s & ";%'" _
                & " and U_Actif=true"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                build_liste_pers = 0
                Exit Function
            End If
            While Not rs.EOF
                nomutil = rs("U_Nom").Value + " " + rs("U_Prenom").Value
                Call CL_AddLigne(nomutil, 0, "", False)
                nbitem = nbitem + 1
                rs.MoveNext
            Wend
            rs.Close
        Case "U"
            numutil = Mid$(s, 2)
            sql = "select U_Nom, U_Prenom from Utilisateur" _
                & " where U_Num=" & numutil
            If Odbc_RecupVal(sql, nom, prenom) = P_ERREUR Then
                Exit Function
            End If
            nomutil = nom + " " + prenom
            Call CL_AddLigne(nomutil, 0, "", False)
            nbitem = nbitem + 1
        End Select
    Next i

    If nbitem > 0 Then
        Call CL_Tri(0)
    End If
    
    build_liste_pers = nbitem
    
End Function

Private Sub build_SP(ByRef r_spm As Variant)

    Dim s As String, sp As String
    Dim encore As Boolean
    Dim i As Integer, j As Integer, n As Integer
    Dim num As Long
    Dim nd As Node, ndp As Node
    
    r_spm = ""
    
    For i = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(i)
        If nd.Children > 0 Then
            GoTo lab_suivant
        End If
        If left$(nd.key, 1) = "S" Then
            r_spm = r_spm + nd.key + ";" + "|"
        Else
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
lab_suivant:
    Next i
    
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Sub build_SP2(ByRef r_spm As Variant)

    Dim s As String, sp As String
    Dim encore As Boolean
    Dim i As Integer, j As Integer, n As Integer
    Dim num As Long
    Dim nd As Node, ndp As Node
    
    r_spm = ""
    
    For i = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(i)
        If nd.Children > 0 Then
            GoTo lab_suivant
        End If
        If nd.image <> IMGT_SERVICE_NOMI And nd.image <> IMGT_POSTE_NOMI Then
            GoTo lab_suivant
        End If
        If left$(nd.key, 1) = "S" Then
            r_spm = r_spm + nd.key + ";" + "|"
        Else
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
lab_suivant:
    Next i
    
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Sub creer_fcttrav(ByRef r_Num As Long, _
                          ByRef r_lib As String)

    Dim sret As String
    Dim frm As Form
    
    'Set frm = KS_PrmFonction
    'sret = KS_PrmFonction.AppelFrm(True)
    Set frm = Nothing
    If sret = "" Then
        r_Num = 0
        Exit Sub
    End If
    r_Num = STR_GetChamp(sret, "|", 0)
    r_lib = STR_GetChamp(sret, "|", 1)
    
End Sub

Private Function creer_groupe() As Long

    Dim frm As Form
    
    'Set frm = PrmGroupePers
    'creer_groupe = PrmGroupePers.AppelFrm(0)
    Set frm = Nothing
    
End Function

Private Sub detail_fcttrav()

    Dim n As Integer
    
    Call CL_Init
    
    n = build_liste_pers("F" & grdFct.TextMatrix(grdFct.row, 0))
    
    If n = 0 Then
        Call MsgBox("Aucune personne ayant la fonction '" & grdFct.TextMatrix(grdFct.row, GRDF_NOM) & "' n'a été trouvée.", vbInformation + vbOKOnly, "")
        grdFct.SetFocus
        Exit Sub
    End If
    Call CL_InitTitreHelp("Listes des personnes ayant la fonction '" & grdFct.TextMatrix(grdFct.row, GRDF_NOM) & "'", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    Call CL_Tri(1)
    ChoixListe.Show 1
    
    grdFct.SetFocus
    
End Sub

Private Sub detail_groupe()

    Dim slst As String
    Dim n As Integer
    Dim rs As rdoResultset
    
    If Odbc_Select("select GU_Lst from GroupeUtil where GU_Num=" & grdGrp.TextMatrix(grdGrp.row, 0), _
                     rs) = P_ERREUR Then
        Exit Sub
    End If
    slst = rs("GU_Lst").Value & ""
    rs.Close
    Call CL_Init
    
    n = build_liste_pers(slst)
    
    If n = 0 Then
        Call MsgBox("Aucune personne n'a été trouvée dans le groupe '" & grdGrp.TextMatrix(grdGrp.row, GRDG_NOM) & "'", vbInformation + vbOKOnly, "")
        grdGrp.SetFocus
        Exit Sub
    End If
    Call CL_InitTitreHelp("Listes des personnes du groupe '" & grdGrp.TextMatrix(grdGrp.row, GRDG_NOM) & "'", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    Call CL_Tri(1)
    ChoixListe.Show 1
    
    grdGrp.SetFocus
    
End Sub

Private Sub detail_service()

    Dim sType As String
    Dim n As Integer
    Dim nd As Node
    
    Set nd = tvSect.SelectedItem
    On Error GoTo lab_no_sel
    If nd.Text = "" Then
    End If
    On Error GoTo 0
    While nd.Children > 0
        Set nd = nd.Child
    Wend
    Set tvSect.SelectedItem = nd
    
    Call CL_Init
    
    n = build_liste_pers(nd.key)
    
    If left$(nd.key, 1) = "S" Then
        sType = "service"
    Else
        sType = "poste"
    End If
    If n = 0 Then
        Call MsgBox("Aucune personne n'est rattachée au " & sType & " '" & nd.Text & ".", vbInformation + vbOKOnly, "")
        tvSect.SetFocus
        Exit Sub
    End If
    Call CL_InitTitreHelp("Listes des personnes rattachées au " & sType & " '" & nd.Text & "'", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    ChoixListe.Show 1
    tvSect.SetFocus
    
    Exit Sub
    
lab_no_sel:
    Call MsgBox("Veuillez sélectionner un poste ou un service.", vbInformation + vbOKOnly, "")
    
End Sub

Private Sub initialiser()

    'g_crfct_autor = P_UtilEstAutorFct("CR_FCTTRAV")
    
    cmd(CMD_OK).Enabled = False
    
    'If p_DestTousAutor Then
        chk(CHK_TOUT).Visible = True
    'Else
        'chk(CHK_TOUT).Visible = False
        frmfctspm.BorderStyle = 0
        frmfctspm.left = frmfctspm.left + 50
        frmfctspm.Width = frmfctspm.Width - 100
        frmfctspm.Top = frmfctspm.Top - 250
    'End If
    
    g_mode_saisie = False
    
    grdGrp.Cols = 2
    grdGrp.ColWidth(0) = 0
    grdGrp.ColWidth(1) = grdGrp.Width
    grdGrp.Rows = 0
    
    grdFct.Cols = 3
    grdFct.ColWidth(0) = 0
    grdFct.ColWidth(1) = 0
    grdFct.ColWidth(2) = grdFct.Width
    grdFct.Rows = 0
    
    grdPers.Cols = 3
    grdPers.ColWidth(0) = 0
    grdPers.ColWidth(1) = 0
    grdPers.ColWidth(2) = grdPers.Width
    grdPers.Rows = 0
    
    If afficher_GFSU() = P_ERREUR Then
        Call quitter
    End If
    
    If grdGrp.Rows = 0 Then
        cmd(CMD_SUPPR_GRP).Visible = False
        cmd(CMD_LOUPE_GRP).Visible = False
    Else
        grdGrp.row = 0
        grdGrp.RowSel = 0
        grdGrp.col = grdGrp.Cols - 1
        grdGrp.ColSel = grdGrp.col
    End If
    
    If grdFct.Rows = 0 Then
        cmd(CMD_SUPPR_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
    Else
        grdFct.row = 0
        grdFct.RowSel = 0
        grdFct.col = grdFct.Cols - 1
        grdFct.ColSel = grdFct.col
    End If
    
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_SUPPR_SRV).Visible = False
        cmd(CMD_LOUPE_SRV).Visible = False
    End If
    
    If grdPers.Rows = 0 Then
        cmd(CMD_SUPPR_PERS).Visible = False
    Else
        grdPers.row = 0
        grdPers.RowSel = 0
        grdPers.col = grdPers.Cols - 1
        grdPers.ColSel = grdPers.col
    End If
    
    g_mode_saisie = True

End Sub

Private Sub maj_fctspm()

    Dim s As String, spm As String
    Dim lig As Integer, i As Integer, n As Integer
    
    lig = 0
    While lig <= grdFct.Rows - 1
        If grdFct.TextMatrix(lig, GRDF_NOMI) = False Then
            If grdFct.Rows = 1 Then
                grdFct.Rows = 0
            Else
                grdFct.RemoveItem lig
            End If
        Else
            lig = lig + 1
        End If
    Wend
    
    i = 1
    While i <= tvSect.Nodes.Count
        If tvSect.Nodes(i).tag = False Then
            tvSect.Nodes.Remove (i)
        Else
            i = i + 1
        End If
    Wend
    
    For lig = 0 To grdGrp.Rows - 1
        Call ajouter_fctspm_grp(grdGrp.TextMatrix(lig, GRDG_NUM), False)
    Next lig
    
    Call maj_pers
    
End Sub

Private Sub maj_pers()

    Dim s As String, spm As String
    Dim lig As Integer, i As Integer, n As Integer
    
Exit Sub

    lig = 0
    While lig <= grdPers.Rows - 1
        If grdPers.TextMatrix(lig, GRDP_NOMI) = False Then
            If grdPers.Rows = 1 Then
                grdPers.Rows = 0
            Else
                grdPers.RemoveItem lig
            End If
        Else
            lig = lig + 1
        End If
    Wend
    
    For lig = 0 To grdGrp.Rows - 1
        Call ajouter_pers_gfs("G" & grdGrp.TextMatrix(lig, GRDG_NUM))
    Next lig
    For lig = 0 To grdFct.Rows - 1
        Call ajouter_pers_gfs("F" & grdFct.TextMatrix(lig, GRDF_NUM))
    Next lig
    Call build_SP(spm)
    n = STR_GetNbchamp(spm, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(spm, "|", i)
        Call ajouter_pers_gfs(s)
    Next i
    
End Sub

Private Function prm_fcttrav() As Integer

    Dim sql As String, lib As String
    Dim trouve As Boolean
    Dim n As Integer, i As Integer, lig As Integer, btn_sortie As Integer
    Dim num As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    n = 0
    For i = 0 To grdFct.Rows - 1
        Call CL_AddLigne(grdFct.TextMatrix(i, GRDF_NOM), grdFct.TextMatrix(i, GRDF_NUM), "", True, grdFct.TextMatrix(i, GRDF_NOMI))
        n = n + 1
    Next i
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        prm_fcttrav = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        trouve = False
        For i = 0 To grdFct.Rows - 1
            If grdFct.TextMatrix(i, GRDF_NUM) = rs("FT_Num").Value Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False, True)
            n = n + 1
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        prm_fcttrav = P_OK
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    'If g_crfct_autor Then
    '    Call CL_AddBouton("&Créer une fonction", "", 0, 0, 1800)
    '    btn_sortie = 2
    'Else
        btn_sortie = 1
    'End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitResteCachée(True)
lab_choix:
    ChoixListe.Show 1
    ' Sortie
    If CL_liste.retour = btn_sortie Then
        GoTo lab_fin
    End If
    
    ' Création
    If CL_liste.retour = 1 Then
        Call creer_fcttrav(num, lib)
        If num > 0 Then
            Call CL_AddLigne(lib, num, "", False)
            n = n + 1
        End If
        GoTo lab_choix
    End If
    
    lig = 0
    While lig <= grdFct.Rows - 1
        If grdFct.TextMatrix(lig, GRDF_NOMI) = True Then
            If grdFct.Rows = 1 Then
                grdFct.Rows = 0
            Else
                grdFct.RemoveItem lig
            End If
        Else
            lig = lig + 1
        End If
    Wend
        
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected And CL_liste.lignes(i).fmodif Then
            Call ajouter_fct_grd(CL_liste.lignes(i).num, True, True)
        End If
    Next i
    
    Call maj_pers
    
    If grdPers.Rows = 0 Then
        cmd(CMD_SUPPR_PERS).Visible = False
    End If
    
    If grdFct.Rows > 0 Then
        cmd(CMD_SUPPR_FCT).Visible = True
        cmd(CMD_LOUPE_FCT).Visible = True
    Else
        cmd(CMD_SUPPR_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
    End If
    cmd(CMD_OK).Enabled = True
    
lab_fin:
    Unload ChoixListe
    prm_fcttrav = P_OK

End Function

Private Function prm_groupe() As Integer

    Dim sql As String, lib As String
    Dim bajout As Boolean, trouve As Boolean, fdetail As Boolean
    Dim n As Integer, i As Integer, lig As Integer
    Dim num As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    n = 0
    For i = 0 To grdGrp.Rows - 1
        Call CL_AddLigne(grdGrp.TextMatrix(i, GRDG_NOM), grdGrp.TextMatrix(i, GRDG_NUM), "", True)
        n = n + 1
    Next i
    sql = "select GU_Num, GU_Nom from GroupeUtil" _
        & " order by GU_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        prm_groupe = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        trouve = False
        For i = 0 To grdGrp.Rows - 1
            If grdGrp.TextMatrix(i, GRDG_NUM) = rs("GU_Num").Value Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then
            Call CL_AddLigne(rs("GU_Nom").Value, rs("GU_Num").Value, "", False)
            n = n + 1
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        prm_groupe = P_OK
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Groupes de personnes", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    Call CL_InitResteCachée(True)
lab_choix:
    ChoixListe.Show 1
    ' Sortie
    If CL_liste.retour = 1 Then
        GoTo lab_fin
    End If
    
    grdGrp.Rows = 0
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
            sql = "select GU_Detailler from GroupeUtil where GU_Num=" & CL_liste.lignes(i).num
            If Odbc_RecupVal(sql, fdetail) = P_ERREUR Then
                prm_groupe = P_ERREUR
                Exit Function
            End If
            If fdetail Then
                Call ajouter_fctspm_grp(CL_liste.lignes(i).num, True)
            Else
                grdGrp.AddItem CL_liste.lignes(i).num & vbTab _
                            & CL_liste.lignes(i).texte
                grdGrp.row = grdGrp.Rows - 1
            End If
        End If
    Next i
    
 '   Call maj_fctspm
    
    If grdGrp.Rows > 0 Then
        cmd(CMD_SUPPR_GRP).Visible = True
        cmd(CMD_LOUPE_GRP).Visible = True
    Else
        cmd(CMD_SUPPR_GRP).Visible = False
        cmd(CMD_LOUPE_GRP).Visible = False
    End If
    cmd(CMD_OK).Enabled = True
    
lab_fin:
    Unload ChoixListe
    prm_groupe = P_OK

End Function

Private Sub prm_service()

    Dim s As String, ss As String, sret As String, sprm As String
    Dim fmodif As Boolean, encore As Boolean
    Dim i As Integer, n As Integer
    Dim numlabo As Long, numutil As Long
    Dim nd As Node
    Dim frm As Form
    
    encore = True
    Do
        n = 0
        i = 1
        Call CL_Init
        While i <= tvSect.Nodes.Count
            Set nd = tvSect.Nodes(i)
            If nd.Children = 0 Then
                s = nd.key & ";"
                fmodif = nd.tag
                While TV_NodeParent(nd)
                    s = nd.key & ";" & s
                Wend
                ReDim Preserve CL_liste.lignes(n)
                CL_liste.lignes(n).texte = s
                CL_liste.lignes(n).fmodif = fmodif
                n = n + 1
            End If
            i = i + 1
        Wend
        Set frm = KS_PrmService
        sret = KS_PrmService.AppelFrm("Choix des services / Postes", "S", True, "", "SP", True)
        Set frm = Nothing
        If sret = "" Then
            encore = False
        ElseIf left$(sret, 1) = "N" Then
            encore = False
        Else
'            Set frm = KS_PrmPersonne
            numutil = STR_GetChamp(sret, "|", 0)
            If numutil = 0 Then
                sprm = "POSTE=" & Mid$(STR_GetChamp(sret, "|", 1), 2)
            Else
                sprm = ""
            End If
'            Call KS_PrmPersonne.AppelFrm(numutil, sprm)
            Set frm = Nothing
        End If
    Loop Until encore = False
    
    If sret = "" Then
        Exit Sub
    End If
    
    cmd(CMD_OK).Enabled = True
    
    tvSect.Nodes.Clear
    n = CLng(Mid$(sret, 2))
    If n = 0 Then
        Exit Sub
    End If
    For i = 0 To n - 1
        s = CL_liste.lignes(i).texte
        Call ajouter_spm_tv(s, CL_liste.lignes(i).tag, True)
    Next i
    
    Call maj_pers
    
    tvSect.SetFocus
    If tvSect.Nodes.Count > 0 Then
        cmd(CMD_SUPPR_SRV).Visible = True
        cmd(CMD_LOUPE_SRV).Visible = True
        Set tvSect.SelectedItem = tvSect.Nodes(1)
        SendKeys "{PGDN}"
        SendKeys "{HOME}"
        DoEvents
    End If
    
End Sub

Private Function quitter() As Boolean

    Dim reponse As Integer
    
    If cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then
            quitter = False
            Exit Function
        End If
    End If
    
    g_bcr = False
    Unload Me
    
    quitter = True
    
End Function

Private Sub supprimer_fcttrav()

    If grdFct.TextMatrix(grdFct.row, GRDP_NOMI) = False Then
        Call MsgBox("Vous ne pouvez pas supprimer cette fonction car elle en fait partie dans les groupes indiqués.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    If grdFct.Rows = 1 Then
        grdFct.Rows = 0
        cmd(CMD_SUPPR_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
    Else
        grdFct.RemoveItem (grdFct.row)
        grdFct.row = 0
    End If
    Call maj_pers
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub supprimer_groupe()

    If grdGrp.Rows = 1 Then
        grdGrp.Rows = 0
        cmd(CMD_SUPPR_GRP).Visible = False
        cmd(CMD_LOUPE_GRP).Visible = False
    Else
        grdGrp.RemoveItem (grdGrp.row)
        grdGrp.row = 0
    End If
    cmd(CMD_OK).Enabled = True
    
    ' Call maj_fctspm
    
End Sub

Private Sub supprimer_service()

    Dim s As String
    
    If tvSect.Nodes.Count = 0 Then
        Exit Sub
    End If
    
    On Error GoTo err_tv
    s = tvSect.SelectedItem.key
    On Error GoTo 0
    
    If tvSect.SelectedItem.tag = False Then
        Call MsgBox("Vous ne pouvez pas supprimer ce service car il en fait partie dans les groupes indiqués.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    tvSect.Nodes.Remove (tvSect.SelectedItem.Index)
    tvSect.Refresh
    cmd(CMD_OK).Enabled = True
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_SUPPR_SRV).Visible = False
        cmd(CMD_LOUPE_SRV).Visible = False
    End If
    Call maj_pers
    Exit Sub
    
err_tv:
    MsgBox "Vous devez sélectionner l'élément à supprimer.", vbOKOnly, ""
    On Error GoTo 0
    
End Sub

Private Sub supprimer_pers()

    If grdPers.Rows = 1 Then
        grdPers.Rows = 0
        cmd(CMD_SUPPR_PERS).Visible = False
    Else
        grdPers.RemoveItem (grdPers.row)
        grdPers.row = 0
    End If
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub supprimer_str_SP(ByVal v_spm As String)

    Dim spm As String, s As String
    Dim n As Integer, i As Integer
    
    spm = ""
    n = STR_GetNbchamp(v_spm, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(v_spm, "|", i)
        If InStr(s, v_spm) = 0 Then
            If spm <> "" Then
                spm = spm + "|"
            End If
            spm = spm + s
        End If
    Next i
    g_sfctspm = spm
    
End Sub

Private Sub valider()

    Dim spm As String
    Dim i As Integer
    
    g_bcr = True
    If chk(CHK_TOUT).Value = 1 Then
        g_sfctspm = "0"
    Else
        g_sfctspm = ""
        For i = 0 To grdGrp.Rows - 1
            g_sfctspm = g_sfctspm & "G" & grdGrp.TextMatrix(i, GRDG_NUM) & "|"
        Next i
        For i = 0 To grdFct.Rows - 1
            If grdFct.TextMatrix(i, GRDF_NOMI) = True Then
                g_sfctspm = g_sfctspm & "F" & grdFct.TextMatrix(i, GRDF_NUM) & "|"
            End If
        Next i
        If tvSect.Nodes.Count > 0 Then
            Call build_SP2(spm)
            g_sfctspm = g_sfctspm & spm
        End If
        For i = 0 To grdPers.Rows - 1
            If grdPers.TextMatrix(i, GRDP_NOMI) = True Then
                g_sfctspm = g_sfctspm & "U" & grdPers.TextMatrix(i, GRDP_NUM) & "|"
            End If
        Next i
    End If
    
    Unload Me
    
End Sub

Private Sub chk_Click(Index As Integer)

    If Index = CHK_TOUT Then
        If chk(Index).Value = 0 Then
            frmfctspm.Visible = True
            frmfctspm.Enabled = True
        Else
            frmfctspm.Visible = False
        End If
    End If
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case CMD_PRM_GRP
        Call prm_groupe
    Case CMD_SUPPR_GRP
        Call supprimer_groupe
    Case CMD_LOUPE_GRP
        Call detail_groupe
    Case CMD_PRM_PERS
        Call ajouter_pers
    Case CMD_SUPPR_PERS
        Call supprimer_pers
    Case CMD_PRM_FCT
        Call prm_fcttrav
    Case CMD_LOUPE_FCT
        Call detail_fcttrav
    Case CMD_SUPPR_FCT
        Call supprimer_fcttrav
    Case CMD_PRM_SRV
        Call prm_service
    Case CMD_SUPPR_SRV
        Call supprimer_service
    Case CMD_LOUPE_SRV
        Call detail_service
    Case CMD_OK
        Call valider
    Case CMD_FERMER
        Call quitter
    End Select

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

Private Sub grdFct_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call prm_fcttrav
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_fcttrav
    End If
    
End Sub

Private Sub grdGrp_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call prm_groupe
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_groupe
    End If
    
End Sub

Private Sub grdPers_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call ajouter_pers
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_pers
    End If
    
End Sub

Private Sub tvSect_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call prm_service
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_service
    End If
    
End Sub

