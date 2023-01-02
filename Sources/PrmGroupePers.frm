VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PrmGroupePers 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8175
   ClientLeft      =   3795
   ClientTop       =   1695
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      Groupe de personnes"
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
      Height          =   7575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7935
      Begin ComctlLib.TreeView tvSect 
         Height          =   1725
         Left            =   1350
         TabIndex        =   4
         Top             =   3600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3043
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
         Index           =   10
         Left            =   7290
         Picture         =   "PrmGroupePers.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Liste des personnes associées à cette fonction"
         Top             =   2280
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   11
         Left            =   7290
         Picture         =   "PrmGroupePers.frx":056D
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Liste des personnes associées à ce service"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Détailler le groupe quand c'est un destinataire"
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
         Left            =   4740
         TabIndex        =   1
         Top             =   450
         Width           =   2985
      End
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
         Left            =   120
         TabIndex        =   22
         Top             =   7200
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
            TabIndex        =   23
            Top             =   660
            Width           =   7785
         End
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Top             =   540
         Width           =   2355
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   320
         Index           =   8
         Left            =   7140
         Picture         =   "PrmGroupePers.frx":0ADA
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer la personne"
         Top             =   6990
         UseMaskColor    =   -1  'True
         Width           =   300
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
         Height          =   315
         Index           =   7
         Left            =   7140
         Picture         =   "PrmGroupePers.frx":0F21
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ajouter une personne"
         Top             =   5580
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   2
         Top             =   1080
         Width           =   5805
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   320
         Index           =   6
         Left            =   7140
         Picture         =   "PrmGroupePers.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer le service"
         Top             =   5010
         UseMaskColor    =   -1  'True
         Width           =   300
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
         Height          =   320
         Index           =   5
         Left            =   7140
         Picture         =   "PrmGroupePers.frx":17BF
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Accéder aux services"
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   300
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
         Height          =   320
         Index           =   3
         Left            =   7170
         Picture         =   "PrmGroupePers.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Accéder aux fonctions"
         Top             =   1620
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   320
         Index           =   4
         Left            =   7170
         Picture         =   "PrmGroupePers.frx":206D
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer la fonction"
         Top             =   3030
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin MSFlexGridLib.MSFlexGrid grdFct 
         Height          =   1725
         Left            =   1350
         TabIndex        =   3
         Top             =   1620
         Width           =   5805
         _ExtentX        =   10239
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
      Begin MSFlexGridLib.MSFlexGrid grdPers 
         Height          =   1725
         Left            =   1320
         TabIndex        =   5
         Top             =   5610
         Width           =   5805
         _ExtentX        =   10239
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
      Begin VB.Image img 
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "PrmGroupePers.frx":24B4
         Top             =   0
         Width           =   300
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "PrmGroupePers.frx":2950
         Top             =   2040
         Width           =   300
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   2
         Left            =   360
         Picture         =   "PrmGroupePers.frx":2DAA
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   3
         Left            =   480
         Picture         =   "PrmGroupePers.frx":332F
         Top             =   6000
         Width           =   300
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
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
         Index           =   2
         Left            =   330
         TabIndex        =   21
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Personnes"
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
         TabIndex        =   20
         Top             =   5670
         Width           =   945
      End
      Begin ComctlLib.ImageList ImageListS 
         Left            =   1200
         Top             =   3480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmGroupePers.frx":378E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmGroupePers.frx":3FE0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
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
         Left            =   330
         TabIndex        =   17
         Top             =   1080
         Width           =   495
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
         Index           =   5
         Left            =   270
         TabIndex        =   16
         Top             =   1710
         Width           =   945
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
         Index           =   4
         Left            =   270
         TabIndex        =   15
         Top             =   3720
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   765
      Left            =   0
      TabIndex        =   6
      Top             =   7440
      Width           =   7935
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmGroupePers.frx":48B2
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
         Picture         =   "PrmGroupePers.frx":4E0E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   500
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
         Left            =   6930
         Picture         =   "PrmGroupePers.frx":5377
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   500
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmGroupePers.frx":5930
         Height          =   510
         Index           =   2
         Left            =   3600
         Picture         =   "PrmGroupePers.frx":5EBF
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer le groupe"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   500
      End
   End
End
Attribute VB_Name = "PrmGroupePers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TXT_CODE = 1
Private Const TXT_NOM = 0

Private Const CHK_DETAIL = 0

Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_SUPPRIMER = 2
Private Const CMD_PLUS_FCT = 3
Private Const CMD_MOINS_FCT = 4
Private Const CMD_PLUS_SERVICE = 5
Private Const CMD_MOINS_SERVICE = 6
Private Const CMD_PLUS_PERS = 7
Private Const CMD_MOINS_PERS = 8
Private Const CMD_LOUPE_FCT = 10
Private Const CMD_LOUPE_SRV = 11

' Images TreeView des services
Private Const IMGT_SERVICE = 1
Private Const IMGT_POSTE = 2

Private Const GRDP_NUMUTIL = 0
Private Const GRDP_NOMI = 1
Private Const GRDP_NOMUTIL = 2

Private g_numgrp As Long
Private g_slst As String

Private g_mode_prm As Boolean

Private g_crgrp_autor As Boolean
Private g_crfct_autor As Boolean

Private g_mode_saisie As Boolean
Private g_txt_avant As String
Private g_form_active As Boolean
Private g_form_width As Long, g_form_height As Long

Public Function AppelFrm(ByVal v_mode As Integer) As Long

    If v_mode = 0 Then
        g_mode_prm = False
    Else
        g_mode_prm = True
    End If
    
    PrmGroupePers.Show 1
    
    AppelFrm = g_numgrp
    
End Function

Private Sub afficher_frm_valid()

    frmValid.Visible = True
    frmValid.ZOrder 0
    Me.Height = frmValid.Height
    Me.Width = frmValid.Width
    frmValid.Top = 0
    frmValid.left = 0
    Call FRM_CentrerForm(Me)
    lblValid.Caption = "Mise à jour des changements dans les destinataires des documents."
    Me.Refresh
    DoEvents
    
End Sub

Private Function afficher_groupe() As Integer

    Dim sql As String, s As String
    Dim rs As rdoResultset
    
    Me.MousePointer = 11
    
    tvSect.Nodes.Clear
    grdFct.Rows = 0
    grdPers.Rows = 0
    
    If g_numgrp > 0 Then
        sql = "select * from GroupeUtil" _
            & " where GU_Num=" & g_numgrp
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        txt(TXT_CODE).Text = rs("GU_Code").Value
        txt(TXT_NOM).Text = rs("GU_Nom").Value
        chk(CHK_DETAIL).Value = IIf(rs("GU_Detailler").Value = True, 1, 0)
        g_slst = rs("GU_Lst").Value & ""
        If afficher_liste(g_slst) = P_ERREUR Then
            GoTo lab_err
        End If
        rs.Close
        If grdPers.Rows > 0 Then
            cmd(CMD_MOINS_PERS).Visible = True
        End If
        
        cmd(CMD_OK).Enabled = False
        cmd(CMD_SUPPRIMER).Enabled = True
    Else
        txt(TXT_CODE).Text = ""
        txt(TXT_NOM).Text = ""
        chk(CHK_DETAIL).Value = 0
        cmd(CMD_LOUPE_SRV).Visible = False
        cmd(CMD_MOINS_SERVICE).Visible = False
        cmd(CMD_MOINS_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
        cmd(CMD_OK).Enabled = True
        cmd(CMD_SUPPRIMER).Enabled = False
        cmd(CMD_MOINS_PERS).Visible = False
    End If
    
    txt(TXT_CODE).SetFocus
    g_mode_saisie = True
    
    Me.MousePointer = 0
    
    afficher_groupe = P_OK
    Exit Function
    
lab_err:
    Me.MousePointer = 0
    afficher_groupe = P_ERREUR

End Function

Private Function afficher_liste(ByVal v_lst As String) As Integer

    Dim s As String, lib As String, s1 As String, sql As String
    Dim n As Integer, i As Integer, lig As Integer, n2 As Integer, j As Integer
    Dim numfct As Long, numutil As Long, num As Long
    Dim nd As Node
    
    n = STR_GetNbchamp(v_lst, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(v_lst, "|", i)
        Select Case left$(s, 1)
        Case "F"
            numfct = CLng(Mid$(s, 2))
            If Odbc_RecupVal("select FT_Libelle from FctTrav where FT_Num=" & numfct, _
                             lib) = P_ERREUR Then
                afficher_liste = P_ERREUR
                Exit Function
            End If
            lig = -1
            For j = 0 To grdFct.Rows - 1
                If UCase(grdFct.TextMatrix(j, 1)) > UCase(lib) Then
                    lig = j
                    Exit For
                End If
            Next j
            If lig >= 0 Then
                grdFct.AddItem numfct & vbTab & lib, lig
            Else
                grdFct.AddItem numfct & vbTab & lib
            End If
'            Call ajouter_pers_fctspm(s)
        Case "S"
'            Call ajouter_pers_fctspm(s)
            n2 = STR_GetNbchamp(s, ";")
            For j = 1 To n2
                s1 = STR_GetChamp(s, ";", j - 1)
                If TV_NodeExiste(tvSect, s1, nd) = P_OUI Then
                    GoTo lab_sp_suiv
                End If
                num = CLng(Mid$(s1, 2))
                If left(s1, 1) = "S" Then
                    If P_RecupSrvNom(num, lib) = P_ERREUR Then
                        afficher_liste = P_ERREUR
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
                        afficher_liste = P_ERREUR
                        Exit Function
                    End If
                    Call tvSect.Nodes.Add(nd, tvwChild, "P" & num, lib, IMGT_POSTE, IMGT_POSTE)
                End If
lab_sp_suiv:
            Next j
        Case "U"
            numutil = Mid$(s, 2)
            Call ajouter_pers_grd(numutil, True, False)
        End Select
    Next i
        
    If grdFct.Rows = 0 Then
        cmd(CMD_MOINS_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
    Else
        cmd(CMD_MOINS_FCT).Visible = True
        cmd(CMD_LOUPE_FCT).Visible = True
        grdFct.row = 0
        grdFct.RowSel = 0
        grdFct.col = grdFct.Cols - 1
        grdFct.ColSel = grdFct.col
    End If
    
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_LOUPE_SRV).Visible = True
        cmd(CMD_MOINS_SERVICE).Visible = True
    Else
        cmd(CMD_LOUPE_SRV).Visible = False
        cmd(CMD_MOINS_SERVICE).Visible = False
    End If
    
    If grdPers.Rows = 0 Then
        cmd(CMD_MOINS_PERS).Visible = False
    Else
        cmd(CMD_MOINS_PERS).Visible = True
        grdPers.row = 0
        grdPers.RowSel = 0
        grdPers.col = grdPers.Cols - 1
        grdPers.ColSel = grdPers.col
    End If
    
    afficher_liste = P_OK
    
End Function

Private Sub ajouter_pers()

    Dim sret As String
    Dim bajout As Boolean
    Dim i As Integer, cr As Integer
    Dim numutil As Long
    Dim frm As Form
    
    p_siz_tblu = -1
    Set frm = ChoixUtilisateur
    sret = ChoixUtilisateur.AppelFrm("Choix des personnes", _
                                     "", _
                                     True, _
                                     False, _
                                     "", _
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
            cmd(CMD_MOINS_PERS).Visible = True
        End If
        cmd(CMD_OK).Enabled = True
    End If
      
End Sub

Private Sub ajouter_pers_fctspm(ByVal v_fctspm As String)

    Dim sql As String
    Dim rs As rdoResultset
    
    Select Case left$(v_fctspm, 1)
    Case "F"
        sql = "select U_Num from Utilisateur" _
            & " where U_FctTrav like '%" & v_fctspm & ";%'"
    Case "S"
        sql = "select U_Num from Utilisateur" _
            & " where U_SPM like '%" & v_fctspm & "%'"
    End Select
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call ajouter_pers_grd(rs("U_Num").Value, False, False)
        rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Function ajouter_pers_grd(ByVal v_numutil As Long, _
                                  ByVal v_nomi As Boolean, _
                                  ByVal v_mess_y_est As Boolean) As Integer

    Dim nomutil As String
    Dim lig As Integer, j As Integer
    
    If P_RecupUtilNomP(v_numutil, nomutil) = P_ERREUR Then
        ajouter_pers_grd = P_ERREUR
        Exit Function
    End If
    lig = -1
    For j = 0 To grdPers.Rows - 1
        If grdPers.TextMatrix(j, GRDP_NUMUTIL) = v_numutil Then
            If v_nomi = False Then
                grdPers.TextMatrix(j, GRDP_NOMI) = False
                grdPers.row = j
                grdPers.col = GRDP_NOMUTIL
                grdPers.CellForeColor = P_GRIS_FONCE
            ElseIf v_mess_y_est Then
                Call MsgBox("'" & nomutil & "' est déjà dans la liste.", vbInformation + vbOKOnly, "")
            End If
            ajouter_pers_grd = P_NON
            Exit Function
        End If
        If UCase(grdPers.TextMatrix(j, GRDP_NOMUTIL)) > UCase(nomutil) Then
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
    grdPers.TextMatrix(lig, GRDP_NUMUTIL) = v_numutil
    grdPers.TextMatrix(lig, GRDP_NOMI) = v_nomi
    grdPers.TextMatrix(lig, GRDP_NOMUTIL) = nomutil
    If Not v_nomi Then
        grdPers.row = lig
        grdPers.col = GRDP_NOMUTIL
        grdPers.CellForeColor = P_GRIS_FONCE
    End If
    
    ajouter_pers_grd = P_OUI
    
End Function

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
            sql = "select U_Num, U_Nom, U_Prenom, U_SPM from Utilisateur" _
                & " where U_FctTrav like '%" & s & ";%'" _
                & " and U_Actif=true" _
                & " order by U_Nom, U_Prenom"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                build_liste_pers = 0
                Exit Function
            End If
            While Not rs.EOF
                nomutil = rs("U_Nom").Value + " " + rs("U_Prenom").Value
'                If Odbc_RecupVal("select SRV_Nom from Service where SRV_Num=" & numsrv, lib) = P_ERREUR Then
'                    lib = ""
'                End If
                Call CL_AddLigne(nomutil, 0, "", False)
                nbitem = nbitem + 1
                rs.MoveNext
            Wend
            rs.Close
        Case "S", "P"
            sql = "select U_Num, U_Nom, U_Prenom, U_FctTrav from Utilisateur" _
                & " where U_SPM like '%" & s & "%'" _
                & " and U_Actif=true" _
                & " order by U_Nom, U_Prenom"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                build_liste_pers = 0
                Exit Function
            End If
            While Not rs.EOF
                nomutil = rs("U_Nom").Value + " " + rs("U_Prenom").Value
'                If Odbc_RecupVal("select FT_Libelle from FctTrav where FT_Num=" & numfct, lib) = P_ERREUR Then
'                    lib = ""
'                End If
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

    build_liste_pers = nbitem
    
End Function

Private Sub build_SP(ByRef r_srv As Variant)

    Dim s As String, sp As String, sql As String
    Dim encore As Boolean
    Dim i As Integer, j As Integer, n As Integer
    Dim numfct As Long, num As Long
    Dim nd As Node, ndp As Node
    
    r_srv = ""
    
    For i = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(i)
        If nd.Children = 0 Then
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
                r_srv = r_srv + STR_GetChamp(sp, ";", j) & ";"
            Next j
            r_srv = r_srv + "|"
        End If
    Next i
    
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Sub build_sql_dest(ByVal v_sdest As String, _
                           ByVal v_ssite As String, _
                           ByRef r_sql As String)

    Dim clause_labo As String, clause As String, s As String, sdest As String, slstgrp As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
    clause_labo = ""
    clause = " and ("
    n = STR_GetNbchamp(v_ssite, ";")
    For i = 0 To n - 1
        clause_labo = clause_labo & clause & "U_Labo like " & Odbc_String("*" + STR_GetChamp(v_ssite, ";", i) + ";*")
        clause = " or "
    Next i
    
    r_sql = "select U_Num, U_Nom, U_Prenom, U_AR" _
            & " from Utilisateur" _
            & " where U_Actif=True" _
            & clause_labo & ")"
    If v_sdest <> "" And v_sdest <> "0" Then
        n = STR_GetNbchamp(v_sdest, "|")
        sdest = ""
        For i = 0 To n - 1
            s = STR_GetChamp(v_sdest, "|", i)
            Select Case left$(s, 1)
            Case "G"
                If Odbc_Select("select GU_Lst from GroupeUtil where GU_Num=" & Mid$(s, 2), _
                                 rs) = P_ERREUR Then
                    Call quitter(True)
                    Exit Sub
                End If
                slstgrp = rs("GU_Lst").Value & ""
                rs.Close
                sdest = sdest + slstgrp
            Case "F", "S", "U"
                sdest = sdest & s & "|"
            End Select
        Next i
        n = STR_GetNbchamp(sdest, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(sdest, "|", i)
            If i = 0 Then
                r_sql = r_sql & " and ("
            Else
                r_sql = r_sql & " or"
            End If
            Select Case left$(s, 1)
            Case "F"
                r_sql = r_sql & " U_FctTrav like '%" & s & ";%'"
            Case "S"
                r_sql = r_sql & " U_SPM like '%" & s & "%'"
            Case "U"
                r_sql = r_sql & " U_Num=" & Mid$(s, 2)
            End Select
        Next i
        r_sql = r_sql & ")"
    End If
    
End Sub

Private Function choisir_groupe() As Integer

    Dim sql As String
    Dim n As Integer
    Dim rs As rdoResultset
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    g_mode_saisie = False
    
lab_affiche:
    Call CL_Init
    
    'Choix du groupe
    n = 0
    If g_crgrp_autor Then
        Call CL_AddLigne("<Nouveau>", 0, "", False)
        n = 1
    End If
    sql = "select * from GroupeUtil" _
        & " order by GU_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_groupe = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("GU_Nom").Value, rs("GU_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        Call MsgBox("Aucune groupe n'a été trouvé.", vbInformation + vbOKOnly, "")
        choisir_groupe = P_NON
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Liste des groupes", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_g_groupepers.htm")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnimprimer.gif", vbKeyI, vbKeyF3, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 2 Then
        choisir_groupe = P_NON
        Exit Function
    End If
    ' Imprimer
    If CL_liste.retour = 1 Then
        Call imprimer
        GoTo lab_affiche
    End If
    
    g_numgrp = CL_liste.lignes(CL_liste.pointeur).num
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    If afficher_groupe() = P_ERREUR Then
        choisir_groupe = P_ERREUR
        Exit Function
    End If
    
    choisir_groupe = P_OUI

End Function

Private Sub creer_fcttrav(ByRef r_num As Long, _
                          ByRef r_lib As String)

    Dim sret As String
    Dim frm As Form
    
    Set frm = KS_PrmFonction
    sret = KS_PrmFonction.AppelFrm(True)
    Set frm = Nothing
    If sret = "" Then
        r_num = 0
        Exit Sub
    End If
    r_num = STR_GetChamp(sret, "|", 0)
    r_lib = STR_GetChamp(sret, "|", 1)
    
End Sub

Private Sub detail_fcttrav()

    Dim n As Integer
       
    Call CL_Init
 
    On Error GoTo lab_no_sel
        n = build_liste_pers("F" & grdFct.TextMatrix(grdFct.row, 0))
    
    If n = 0 Then
        Call MsgBox("Aucune personne ayant la fonction '" & grdFct.TextMatrix(grdFct.row, 1) & "' n'a été trouvée.", vbInformation + vbOKOnly, "")
        grdFct.SetFocus
        Exit Sub
    End If
    Call CL_InitTitreHelp("Listes des personnes ayant la fonction '" & grdFct.TextMatrix(grdFct.row, 1) & "'", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    Call CL_Tri(1)
    ChoixListe.Show 1
    
    grdFct.SetFocus
  
    Exit Sub
  
lab_no_sel:
      Call MsgBox("Veuillez sélectionner une fonction.", vbInformation + vbOKOnly, "")

End Sub

Private Sub detail_service()

    Dim stype As String
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
        stype = "service"
    Else
        stype = "poste"
    End If
    If n = 0 Then
        Call MsgBox("Aucune personne n'est rattachée au " & stype & " '" & nd.Text & ".", vbInformation + vbOKOnly, "")
        tvSect.SetFocus
        Exit Sub
    End If
    Call CL_InitTitreHelp("Listes des personnes rattachées au " & stype & " '" & nd.Text & "'", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    ChoixListe.Show 1
    tvSect.SetFocus
    
    Exit Sub
    
lab_no_sel:
    Call MsgBox("Veuillez sélectionner un poste ou un service.", vbInformation + vbOKOnly, "")
    
End Sub

Private Function gerer_ajout_destgrp(ByVal v_numgrp As Long, _
                                     ByVal v_sfctspm As String) As Integer

    Dim sql As String, sfctspm As String, libvers As String
    Dim lnb As Long, numvers As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    
    sql = "select D_Num, D_Dest, D_Site from Document" _
            & " where D_Dest like '%G" & v_numgrp & "|%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        gerer_ajout_destgrp = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        ' Tous les utilisateurs des labos du dossier
        Call build_sql_dest(v_sfctspm, rs("D_Site").Value, sql)
        If Odbc_SelectV(sql, rs2) = P_ERREUR Then
            gerer_ajout_destgrp = P_ERREUR
            Exit Function
        End If
        While Not rs2.EOF
            ' Y est il déjà ?
            sql = "select count(*) from DocPrmDiffusion" _
                & " where DPD_DNum=" & rs("D_Num").Value _
                & " and DPD_UNum=" & rs2("U_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                gerer_ajout_destgrp = P_ERREUR
                Exit Function
            End If
            If lnb = 0 Then
                If P_AjouterDocUtil_Dest(rs("D_Num").Value, _
                                        rs2("U_Num").Value, _
                                        IIf(rs2("U_AR").Value, 2, 1)) = P_ERREUR Then
                    gerer_ajout_destgrp = P_ERREUR
                    Exit Function
                End If
                If rs2("U_AR").Value Then
                    If P_DiffuserLastVersion(rs("D_Num").Value, rs2("U_Num").Value) = P_ERREUR Then
                        gerer_ajout_destgrp = P_ERREUR
                        Exit Function
                    End If
                End If
            End If
            rs2.MoveNext
        Wend
        rs2.Close
        rs.MoveNext
    Wend
    rs.Close
    
    gerer_ajout_destgrp = P_OK
    
End Function

Private Function gerer_chgt_lst(ByVal v_numgrp As Long, _
                                ByVal v_new_slst As String) As Integer

    Dim s As String
    Dim f_aff_frmv As Boolean
    Dim i As Integer, n   As Integer
    
    f_aff_frmv = False
    
    n = STR_GetNbchamp(v_new_slst, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(v_new_slst, "|", i)
        If InStr(g_slst, s & "|") = 0 Then
            If Not f_aff_frmv Then
                f_aff_frmv = True
                Call afficher_frm_valid
            End If
            If gerer_ajout_destgrp(v_numgrp, s) = P_ERREUR Then
                gerer_chgt_lst = P_ERREUR
                Exit Function
            End If
        End If
    Next i
    
    n = STR_GetNbchamp(g_slst, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(g_slst, "|", i)
        If InStr(v_new_slst, s & "|") = 0 Then
            If Not f_aff_frmv Then
                f_aff_frmv = True
                Call afficher_frm_valid
            End If
            If gerer_suppr_destgrp(v_numgrp, s) = P_ERREUR Then
                gerer_chgt_lst = P_ERREUR
                Exit Function
            End If
        End If
    Next i
    
End Function

Private Function gerer_suppr_destgrp(ByVal v_numgrp As Long, _
                                     ByVal v_sfctspm As String) As Integer

    Dim sql As String, sfctspm As String
    Dim trouve As Boolean
    Dim n_util As Integer, i As Integer
    Dim tbl_util() As Long, lnb As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    
    ' Documents
    sql = "select D_Num, D_Dest, D_Site from Document" _
            & " where D_Dest like '%G" & v_numgrp & "|%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        gerer_suppr_destgrp = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        n_util = -1
        ' Tous les utilisateurs des labos du dossier
        Call build_sql_dest(sfctspm, rs("D_Site").Value, sql)
        If Odbc_SelectV(sql, rs2) = P_ERREUR Then
            gerer_suppr_destgrp = P_ERREUR
            Exit Function
        End If
        While Not rs2.EOF
            n_util = n_util + 1
            ReDim Preserve tbl_util(n_util) As Long
            tbl_util(n_util) = rs2("U_Num").Value
            rs2.MoveNext
        Wend
        rs2.Close
        ' Test pour chaque dest déjà présent
        sql = "select * from DocPrmDiffusion" _
            & " where DPD_DNum=" & rs("D_Num").Value
        On Error GoTo err_open_resultset
        Set rs2 = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo 0
        While Not rs2.EOF
            trouve = False
            For i = 0 To n_util
                If tbl_util(i) = rs2("DPD_UNum").Value Then
                    trouve = True
                    Exit For
                End If
            Next i
            ' Ce dest n'y est plus
            If Not trouve Then
                If P_SupprimerDiffusionLastVersion(rs2("DPD_UNum").Value, rs("D_Num").Value) = P_ERREUR Then
                    gerer_suppr_destgrp = P_ERREUR
                    Exit Function
                End If
                On Error GoTo err_edit
                rs2.Edit
                On Error GoTo err_delete
                rs2.Delete
                On Error GoTo 0
            End If
            rs2.MoveNext
        Wend
        rs2.Close
        rs.MoveNext
    Wend
    rs.Close
    
    gerer_suppr_destgrp = P_OK
    Exit Function

err_open_resultset:
    MsgBox "Erreur OpenResultset pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    rs.Close
    gerer_suppr_destgrp = P_ERREUR
    Exit Function
    
err_edit:
    MsgBox "Erreur Edit pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    rs2.Close
    rs.Close
    gerer_suppr_destgrp = P_ERREUR
    Exit Function
    
err_delete:
    MsgBox "Erreur Delete pour " + sql, vbOKOnly + vbCritical, ""
    rs2.Close
    rs.Close
    gerer_suppr_destgrp = P_ERREUR
    Exit Function

End Function

Private Function grp_dans_doc() As Integer

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like 'G" & g_numgrp & "|'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        grp_dans_doc = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        grp_dans_doc = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like 'G" & g_numgrp & "|'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        grp_dans_doc = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        grp_dans_doc = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Document" _
        & " where D_Dest like 'G" & g_numgrp & "|'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        grp_dans_doc = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        grp_dans_doc = P_OUI
        Exit Function
    End If
    
    grp_dans_doc = P_NON
    
End Function

Private Sub imprimer()

    Call MsgBox("A FAIRE")
    
End Sub

Private Sub inhiber_frm_valid()

    frmValid.Visible = False
    Me.Height = g_form_height
    Me.Width = g_form_width
    Me.Refresh
    DoEvents
    
End Sub

Private Sub initialiser()

    grdFct.Cols = 2
    grdFct.ColWidth(0) = 0
    grdFct.ColWidth(1) = grdFct.Width
    
    grdPers.Cols = 3
    grdPers.ColWidth(0) = 0
    grdPers.ColWidth(1) = 0
    grdPers.ColWidth(2) = grdPers.Width
    
    Call maj_droits
    
    If g_mode_prm Then
        If choisir_groupe() <> P_OUI Then
            Unload Me
            Exit Sub
        End If
    Else
        g_numgrp = 0
        If afficher_groupe() = P_ERREUR Then
            Unload Me
            Exit Sub
        End If
    End If
    
End Sub

Private Sub maj_droits()

    g_crgrp_autor = P_UtilEstAutorFct("CR_GRPUTIL")
    g_crfct_autor = P_UtilEstAutorFct("CR_FCTTRAV")
    cmd(CMD_OK).Visible = P_UtilEstAutorFct("MOD_GRPUTIL")
    cmd(CMD_SUPPRIMER).Visible = P_UtilEstAutorFct("SUPP_GRPUTIL")
    
End Sub

Private Sub maj_pers()

    Dim s As String, spm As String
    Dim lig As Integer, i As Integer, n As Integer
    
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
    
    For lig = 0 To grdFct.Rows - 1
        Call ajouter_pers_fctspm("F" & grdFct.TextMatrix(lig, 0))
    Next lig
    Call build_SP(spm)
    n = STR_GetNbchamp(spm, "|")
    For i = 0 To n - 1
        s = STR_GetChamp(spm, "|", i)
        Call ajouter_pers_fctspm(s)
    Next i
    
    
End Sub

Private Function prm_fcttrav() As Integer

    Dim sql As String, lib As String
    Dim bajout As Boolean, trouve As Boolean
    Dim n As Integer, i As Integer, btn_sortie As Integer
    Dim num As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    n = 0
    For i = 0 To grdFct.Rows - 1
        Call CL_AddLigne(grdFct.TextMatrix(i, 1), grdFct.TextMatrix(i, 0), "", True)
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
            If grdFct.TextMatrix(i, 0) = rs("FT_Num").Value Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
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
    If g_crfct_autor Then
        Call CL_AddBouton("&Créer une fonction", "", 0, 0, 1800)
        btn_sortie = 2
    Else
        btn_sortie = 1
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
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
    
    grdFct.Rows = 0
    bajout = False
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
'            Call ajouter_pers_fctspm("F" & CL_liste.lignes(i).num)
            grdFct.AddItem CL_liste.lignes(i).num & vbTab _
                        & CL_liste.lignes(i).texte
            grdFct.row = grdFct.Rows - 1
            bajout = True
        End If
    Next i
    
'    Call maj_pers
    If bajout Then
        cmd(CMD_MOINS_FCT).Visible = True
        cmd(CMD_LOUPE_FCT).Visible = True
        cmd(CMD_OK).Enabled = True
    End If
    
lab_fin:
    Unload ChoixListe
    prm_fcttrav = P_OK

End Function

Private Sub prm_service()

    Dim s As String, s1 As String, sql As String, lib As String
    Dim sret As String, ssite As String, s_srv As String, sprm As String
    Dim au_moins_un As Boolean, encore As Boolean
    Dim i As Integer, j As Integer, n As Integer, n2 As Integer, nbch As Integer
    Dim numlabo As Long, num As Long, numutil As Long
    Dim nd As Node
    Dim rs As rdoResultset
    Dim frm As Form
    
    Call CL_Init
    
    Call build_SP(s_srv)
    nbch = STR_GetNbchamp(s_srv, "|")
    n = 0
    For i = 1 To nbch
        s = STR_GetChamp(s_srv, "|", i - 1)
        ReDim Preserve CL_liste.lignes(n)
        CL_liste.lignes(n).texte = s
        CL_liste.lignes(n).fmodif = True
        n = n + 1
    Next i
    
    If Odbc_Select("select L_Num from Laboratoire", rs) = P_ERREUR Then
        Exit Sub
    End If
    ssite = ""
    While Not rs.EOF
        ssite = ssite & rs("L_Num").Value & ";"
        rs.MoveNext
    Wend
    rs.Close
        
    Set frm = KS_PrmService
    sret = KS_PrmService.AppelFrm("Choix des services / postes", "S", True, ssite, "SP", False)
    Set frm = Nothing
    p_NumLabo = numlabo
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
'        Call ajouter_pers_fctspm(s)
        n2 = STR_GetNbchamp(s, ";")
        For j = 1 To n2
            s1 = STR_GetChamp(s, ";", j - 1)
            If TV_NodeExiste(tvSect, s1, nd) = P_OUI Then
                GoTo lab_sp_suiv
            End If
            num = CLng(Mid$(s1, 2))
            If left(s1, 1) = "S" Then
                If P_RecupSrvNom(num, lib) = P_ERREUR Then
                    Exit Sub
                End If
                If j = 1 Then
                    Set nd = tvSect.Nodes.Add(, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                Else
                    Set nd = tvSect.Nodes.Add(nd, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                End If
                nd.Expanded = True
            Else
                If P_RecupPosteNom(num, lib) = P_ERREUR Then
                    Exit Sub
                End If
                Call tvSect.Nodes.Add(nd, tvwChild, "P" & num, lib, IMGT_POSTE, IMGT_POSTE)
            End If
lab_sp_suiv:
        Next j
    Next i
    
'    Call maj_pers
    
    tvSect.SetFocus
    If tvSect.Nodes.Count > 0 Then
        cmd(CMD_LOUPE_SRV).Visible = True
        cmd(CMD_MOINS_SERVICE).Visible = True
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
            If g_numgrp = 0 Then
                reponse = MsgBox("La création de ce groupe ne s'effectuera pas !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            Else
                reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            End If
            If reponse = vbNo Then
                quitter = False
                Exit Function
            End If
        End If
    End If
    
    If choisir_groupe() <> P_OUI Then
        Unload Me
        quitter = True
        Exit Function
    End If
    
    quitter = False
    
End Function

Private Function supprimer() As Integer

    Dim sql As String
    Dim reponse As Integer, cr As Integer
    Dim lnb As Long
    Dim rs As rdoResultset
    
    If p_appli_kalidoc > 0 Then
        cr = grp_dans_doc()
        If cr = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        If cr = P_OUI Then
            MsgBox "Ce groupe est associé à certains documents." & vbLf & vbCr & "Elle ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, ""
            supprimer = P_OK
            Exit Function
        End If
    End If
    
    reponse = MsgBox("Confirmez-vous la suppression de ce groupe ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        supprimer = P_OK
        Exit Function
    End If
    
    ' Effacement du groupe
    If Odbc_Delete("GroupeUtil", _
                   "GU_Num", _
                   "where GU_Num=" & g_numgrp, _
                   lnb) = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    supprimer = P_OK

End Function

Private Sub supprimer_fcttrav()

    If grdFct.Rows = 1 Then
        grdFct.Rows = 0
        cmd(CMD_MOINS_FCT).Visible = False
        cmd(CMD_LOUPE_FCT).Visible = False
    Else
        grdFct.RemoveItem (grdFct.row)
        grdFct.row = 0
    End If
'    Call maj_pers
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub supprimer_pers()

    If grdPers.TextMatrix(grdPers.row, GRDP_NOMI) = False Then
        Call MsgBox("Vous ne pouvez pas supprimer cette personne du groupe car elle est issue d'une des fonctions ou services que vous avez indiqués.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    If grdPers.Rows = 1 Then
        grdPers.Rows = 0
        cmd(CMD_MOINS_PERS).Visible = False
    Else
        grdPers.RemoveItem (grdPers.row)
        grdPers.row = 0
    End If
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub supprimer_service()

    If tvSect.Nodes.Count = 0 Then Exit Sub
    
    On Error GoTo err_tv
    tvSect.Nodes.Remove (tvSect.SelectedItem.Index)
    On Error GoTo 0
    tvSect.Refresh
    cmd(CMD_OK).Enabled = True
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_MOINS_SERVICE).Visible = False
        cmd(CMD_LOUPE_SRV).Visible = False
    End If
'    Call maj_pers
    Exit Sub
    
err_tv:
    MsgBox "Vous devez sélectionner l'élément à supprimer.", vbOKOnly, ""
    On Error GoTo 0
    
End Sub

Private Sub valider()

    Dim slst As String
    Dim ilig As Integer
    Dim numgrp As Long
    
    If verifier_tous_chp() = P_NON Then
        Exit Sub
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        Exit Sub
    End If
    
    ' Construction GU_Lst
    Call build_SP(slst)
    For ilig = 0 To grdFct.Rows - 1
        slst = slst & "F" & grdFct.TextMatrix(ilig, 0) & "|"
    Next ilig
    For ilig = 0 To grdPers.Rows - 1
        If grdPers.TextMatrix(ilig, GRDP_NOMI) = True Then
            slst = slst & "U" & grdPers.TextMatrix(ilig, GRDP_NUMUTIL) & "|"
        End If
    Next ilig
    
    If g_numgrp = 0 Then
        If Odbc_AddNew("GroupeUtil", _
                       "GU_Num", _
                       "gu_seq", _
                       True, _
                       numgrp, _
                        "GU_Code", txt(TXT_CODE).Text, _
                        "GU_Nom", txt(TXT_NOM).Text, _
                        "GU_Detailler", IIf(chk(CHK_DETAIL).Value = 1, True, False), _
                        "GU_Lst", slst) = P_ERREUR Then
            GoTo err_enreg
        End If
    Else
        numgrp = g_numgrp
        If Odbc_Update("GroupeUtil", _
                       "GU_Num", _
                       "where GU_Num=" & g_numgrp, _
                        "GU_Code", txt(TXT_CODE).Text, _
                        "GU_Nom", txt(TXT_NOM).Text, _
                        "GU_Detailler", IIf(chk(CHK_DETAIL).Value = 1, True, False), _
                        "GU_Lst", slst) = P_ERREUR Then
            GoTo err_enreg
        End If
        If gerer_chgt_lst(g_numgrp, slst) = P_ERREUR Then
            GoTo err_enreg
        End If
        Call inhiber_frm_valid
    End If

    If Odbc_CommitTrans() = P_ERREUR Then
        Exit Sub
    End If
    
    If g_mode_prm Then
        If choisir_groupe() <> P_OUI Then
            Unload Me
            Exit Sub
        End If
    Else
        g_numgrp = numgrp
        Unload Me
    End If
    
    Exit Sub
    
err_enreg:
    Call Odbc_RollbackTrans
    g_numgrp = 0
    Unload Me
        
End Sub

Private Function verifier_tous_chp() As Integer

    If txt(TXT_CODE).Text = "" Then
        Call MsgBox("Le code du groupe est une rubrique obligatoire.", vbExclamation + vbOKOnly, "")
        txt(TXT_CODE).SetFocus
        verifier_tous_chp = P_NON
        Exit Function
    End If
    If verifier_un_chp(TXT_CODE) = P_NON Then
        verifier_tous_chp = P_NON
        Exit Function
    End If
    
    If txt(TXT_NOM).Text = "" Then
        Call MsgBox("Le nom du groupe est une rubrique obligatoire.", vbExclamation + vbOKOnly, "")
        txt(TXT_NOM).SetFocus
        verifier_tous_chp = P_NON
        Exit Function
    End If
    
    If grdPers.Rows = 0 And grdFct.Rows = 0 And tvSect.Nodes.Count = 0 Then
        Call MsgBox("Ce groupe est vide." & vbCrLf & vbCrLf & "Veuillez indiquer des fonctions, services ou personnes.", vbExclamation + vbOKOnly, "")
        grdFct.SetFocus
        verifier_tous_chp = P_NON
        Exit Function
    End If
    
    verifier_tous_chp = P_OUI
    
End Function

Private Function verifier_un_chp(ByVal v_indtxt As Integer) As Integer

    Select Case v_indtxt
    Case TXT_CODE
        If Odbc_EstDoublon("GroupeUtil", _
                            "GU_Code", _
                            txt(TXT_CODE).Text, _
                            "GU_Num", _
                            g_numgrp) = P_OUI Then
            Call MsgBox("Le code '" & txt(TXT_CODE).Text & "' est déjà attribué à un autre groupe." & vbCrLf & "Veuillez choisir un autre code.", vbExclamation + vbOKOnly, "")
            txt(TXT_CODE).Text = ""
            verifier_un_chp = P_NON
        End If
        verifier_un_chp = P_OUI
        Exit Function
    Case Else
        verifier_un_chp = P_OUI
    End Select
    
End Function

Private Sub chk_Click(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_QUITTER
        Call quitter(False)
    Case CMD_SUPPRIMER
        Call supprimer
        Call quitter(True)
    Case CMD_PLUS_FCT
        Call prm_fcttrav
    Case CMD_MOINS_FCT
        Call supprimer_fcttrav
    Case CMD_LOUPE_FCT
        Call detail_fcttrav
    Case CMD_PLUS_SERVICE
        Call prm_service
    Case CMD_MOINS_SERVICE
        Call supprimer_service
    Case CMD_LOUPE_SRV
        Call detail_service
    Case CMD_PLUS_PERS
        Call ajouter_pers
    Case CMD_MOINS_PERS
        Call supprimer_pers
    End Select
    
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
            Call valider
        End If
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_SUPPRIMER).Enabled Then
            Call supprimer
            Call quitter(True)
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_g_groupepers.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
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

Private Sub grdFct_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        If prm_fcttrav() = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        If grdFct.Rows > 0 Then
            Call supprimer_fcttrav
        End If
    End If
    
End Sub

Private Sub grdPers_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call ajouter_pers
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        If grdPers.Rows > 0 Then
            Call supprimer_pers
        End If
    End If
    
End Sub

Private Sub tvSect_BeforeLabelEdit(Cancel As Integer)

    Cancel = True
    
End Sub

Private Sub tvSect_GotFocus()

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
        Call supprimer_service
    End If
    
End Sub

Private Sub tvSect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txt_Change(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_avant = txt(Index).Text
    
End Sub

Private Sub txt_LostFocus(Index As Integer)

    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            If verifier_un_chp(Index) = P_NON Then
                txt(Index).SetFocus
                Exit Sub
            End If
        End If
    End If
    
End Sub


