VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrmListeDetail 
   Caption         =   "Détail"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitter 
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
      Left            =   14640
      Picture         =   "PrmListeDetail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Quitter sans tenir compte des modifications"
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   550
   End
   Begin VB.CommandButton CmdListFormResp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Accès aux [NBRE] [NOM] sélectionné(e)s"
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
      Left            =   120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   14535
   End
   Begin MSFlexGridLib.MSFlexGrid grdCell 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12726
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   16777215
      BackColorBkg    =   16761024
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLegende 
      Alignment       =   2  'Center
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
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   15015
   End
End
Attribute VB_Name = "PrmListeDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public g_NumForm As String
Public g_SQL As String
Public g_NumFiltre As String
Public g_titre As String
Public iG As Integer, iD As Integer
Public numforG As Long, numforD As Long


Public Function AppelFrm(ByVal v_NumForm As String, ByVal v_NumFiltre As String, ByVal v_SQL As String, ByVal v_titre As String) As Boolean

    g_NumForm = v_NumForm
    g_NumFiltre = v_NumFiltre
    g_SQL = v_SQL
    g_titre = v_titre
Faire:
    
    Call initialiser
    
    Me.Show 1
End Function

Private Sub CmdListFormResp_Click()
    Dim url As String
    Dim util As String
        
    url = "filtres/liste_form_resp.php%3FV_numfiltre=" & g_NumFiltre & "%26V_numfor=" & g_NumForm & "%26V_etat=2%26V_etattermine=0" & "%26V_typaff=D%26V_quitter=1" & "%26V_RapportType=" & g_SQL
    
    ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
    util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
        
    If p_S_Vers_Conf <> "" Then
        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
    End If
    url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & url
    
    ' Chargement de la page
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & url, vbMaximizedFocus

End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub grdCell_Click()
    Dim url As String
    Dim util As String
        
    If Me.grdCell.col = iG Then
        If Not p_bool_tbl_detail Then Exit Sub
        url = "form_saisie.php%3FV_numfor=" & p_tbl_detail(Me.grdCell.row).fornumG & "%26V_numdon=" & p_tbl_detail(Me.grdCell.row).donnumG
    ElseIf Me.grdCell.col = iD Then
        If Not p_bool_tbl_detail Then Exit Sub
        url = "form_saisie.php%3FV_numfor=" & p_tbl_detail(Me.grdCell.row).fornumD & "%26V_numdon=" & p_tbl_detail(Me.grdCell.row).donnumD
    Else
        Me.lblLegende.Caption = grdCell.TextMatrix(0, grdCell.col) & " : " & grdCell.TextMatrix(grdCell.row, grdCell.col)
        Exit Sub
    End If
    ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
    util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
        
    If p_S_Vers_Conf <> "" Then
        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
    End If
    url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & url
    
    ' Chargement de la page
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & url, vbMaximizedFocus
    
End Sub

Private Sub initialiser()
    Dim I As Integer
    Dim iW As Integer
    Dim nbCol As Integer
    Dim LibEtapeChp As String, sql As String
    Dim rs As rdoResultset
    Dim coldeb As Integer, colfin As Integer
    
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    Me.grdCell.ScrollBars = flexScrollBarVertical
    
    If p_FaireHyperLienListeChamp Then
        Me.grdCell.ColWidth(0) = 1000
        If numforD > 0 Then
            Me.grdCell.ColWidth(Me.grdCell.Cols - 1) = 1000
            iW = Me.grdCell.Width - (2 * Me.grdCell.ColWidth(0)) - 100
        Else
            iW = Me.grdCell.Width - Me.grdCell.ColWidth(0) - 100
        End If
    Else
        iW = Me.grdCell.Width - 100
    End If
    If p_FaireHyperLienListeChamp Then
        If numforD > 0 Then
            nbCol = Me.grdCell.Cols - 2
            coldeb = 1
            colfin = Me.grdCell.Cols - 2
        Else
            nbCol = Me.grdCell.Cols - 1
            coldeb = 1
            colfin = Me.grdCell.Cols - 1
        End If
    Else
        nbCol = Me.grdCell.Cols
        coldeb = 0
        colfin = nbCol - 1
    End If
    If p_FaireHyperLienListeChamp Then
        For I = coldeb To colfin
            Me.grdCell.ColWidth(I) = (iW / nbCol) - 20
        Next I
    Else
        For I = coldeb To colfin
            Me.grdCell.ColWidth(I) = (iW / nbCol) - 20
        Next I
    End If
    If p_FaireHyperLienListeChamp Then
        For I = 1 To (Me.grdCell.Rows - 1)
            Me.grdCell.row = I
            Me.grdCell.col = 0
            Me.grdCell.CellForeColor = Me.CmdListFormResp.BackColor
            Me.grdCell.CellFontBold = True
            If numforD > 0 Then
                Me.grdCell.col = iD
                Me.grdCell.CellForeColor = Me.CmdListFormResp.BackColor
                Me.grdCell.CellFontBold = True
            End If
        Next I
    End If
    
    sql = "select For_Code from formulaire where For_Num = " & numforG
    Call Odbc_RecupVal(sql, libetape)
    Me.grdCell.TextMatrix(0, IIf(p_FaireHyperLienListeChamp, iG, 0)) = libetape
    
    If numforD > 0 Then
        sql = "select For_Code from formulaire where For_Num = " & numforD
        Call Odbc_RecupVal(sql, libetape)
        Me.grdCell.TextMatrix(0, IIf(p_FaireHyperLienListeChamp, iD, colfin)) = libetape
    End If
    Me.CmdListFormResp.Caption = Replace(Me.CmdListFormResp.Caption, "[NOM]", "[" & g_titre & "]")
    Me.CmdListFormResp.Caption = Replace(Me.CmdListFormResp.Caption, "[NBRE]", (Me.grdCell.Rows - 1))
End Sub

