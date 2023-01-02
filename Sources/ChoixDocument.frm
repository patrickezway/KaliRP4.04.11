VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ChoixDocument 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix d'un document"
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
      Height          =   7515
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11910
      Begin ComctlLib.TreeView tv 
         Height          =   6495
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   11456
         _Version        =   327682
         Indentation     =   0
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "img"
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
      Begin ComctlLib.ImageList img 
         Left            =   4050
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDocument.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDocument.frx":06C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDocument.frx":085C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDocument.frx":0E96
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   11910
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
         Index           =   1
         Left            =   900
         Picture         =   "ChoixDocument.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sélectionner"
         Top             =   230
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
         Index           =   0
         Left            =   10560
         Picture         =   "ChoixDocument.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Annuler l'opération"
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entrée : Toutes les documentations affichées (true/false) | sélection de dossier autorisée (true/false)
' De plus p_tabdoc_present contient le D_Num à ne pas afficher

Private Const CMD_OK = 1
Private Const CMD_QUITTER = 0

Private Const IMG_DOCS = 1
Private Const IMG_DOS = 2
Private Const IMG_DOC = 3
Private Const IMG_DOC_SEL = 4

' 0:la documentation en cours / 1: toutes les documentations que l'utilisateur peut voir / 2:toutes les documentations
Private g_toutes_les_docs As Integer
Private g_tous_les_doc As Boolean
Private g_seldos_autor As Boolean
Private g_plusieurs As Boolean
Private g_schx As String

Private g_mode_saisie As Boolean
Private g_form_active As Boolean

Public Function AppelFrm(ByVal v_toutes_les_docs As Integer, _
                         ByVal v_tous_les_doc As Boolean, _
                         ByVal v_seldos_autor As Boolean, _
                         ByVal v_plusieurs As Boolean) As String
                             
    g_toutes_les_docs = v_toutes_les_docs
    g_tous_les_doc = v_tous_les_doc
    g_seldos_autor = v_seldos_autor
    g_plusieurs = v_plusieurs
    
    ChoixDocument.Show 1
    
    AppelFrm = g_schx
    
End Function

Private Function afficher_documents() As Integer
    
    Dim sql As String, key As String, stext As String, sclause As String
    Dim a_afficher As Boolean
    Dim i As Integer, nbdoc As Integer
    Dim nd As Node
    Dim rs As rdoResultset
    
    tv.Nodes.Clear
    
    If g_toutes_les_docs > 0 Then
        sql = "select DO_Num, DO_Titre" _
            & " from Documentation" _
            & " order by DO_Titre"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_documents = P_ERREUR
            Exit Function
        End If
        While Not rs.EOF
'            If g_toutes_les_docs = 1 Then
'                a_afficher = P_UtilEstConcerneParDocs(p_NumUtil, rs("DO_Num").Value)
'                If Not a_afficher Then
'                    a_afficher = P_EstDocsPublic(rs("DO_Num").Value)
'                End If
'            Else
'                a_afficher = True
'            End If
'            If a_afficher Then
                Set nd = tv.Nodes.Add(, tvwChild, "O" & rs("DO_Num").Value, rs("DO_Titre").Value, IMG_DOCS, IMG_DOCS)
                nd.Expanded = True
'            End If
            rs.MoveNext
        Wend
        rs.Close
    Else
        sql = "select DO_Titre" _
            & " from Documentation" _
            & " where DO_Num=" & p_NumDocs
        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_documents = P_ERREUR
            Exit Function
        End If
        Set nd = tv.Nodes.Add(, tvwChild, "O" & p_NumDocs, rs("DO_Titre").Value, IMG_DOCS, IMG_DOCS)
        nd.Expanded = True
'        nd.Sorted = True
        rs.Close
    End If
    
    If g_toutes_les_docs = 2 Then
        sql = "select D_Num, D_Ordre, D_Ident, D_Titre, D_DSNum" _
            & ", DS_NumPere, DS_Ordre" _
            & " from Document, Dossier"
        sclause = " where "
    ElseIf Not g_tous_les_doc Then
        sql = "select D_Num, D_Ordre, D_Ident, D_Titre, D_DSNum" _
            & ", DS_NumPere, DS_Ordre" _
            & " from Document, Dossier" _
            & " where (D_UNumResp=" & p_NumUtil & " or" _
            & " D_LstResp like '%U" & p_NumUtil & ";1;%')"
        sclause = " and "
    Else
        sql = "select D_Num, D_Ordre, D_Ident, D_Titre, D_DSNum" _
            & ", DS_NumPere, DS_Ordre" _
            & " from Document, Dossier" _
            & " where (D_Public=true" _
            & " or D_Num in (select DU_DNum from DocUtil where DU_UNum=" & p_NumUtil & ")" _
            & " or D_Num in (select DPD_DNum from DocPrmDiffusion where DPD_UNum=" & p_NumUtil & ")" _
            & " or D_LstResp like '%U" & p_NumUtil & ";%')"
        sclause = " and "
    End If
    If g_toutes_les_docs = 0 Then
        sql = sql & " and D_DONum=" & p_NumDocs
        sclause = " and "
    End If
    sql = sql & sclause & "DS_Num=D_DSNum" _
        & " and D_CYordre<=" & p_cycle_consultable _
        & " order by DS_NumPere, DS_Ordre, D_Ordre"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_documents = P_ERREUR
        Exit Function
    End If
    nbdoc = 0
    While Not rs.EOF
        For i = 0 To CM_UboundL(p_tabdoc_present)
            If p_tabdoc_present(i) = rs("D_Num").Value Then
                GoTo lab_suivant
            End If
        Next i
        If P_AfficherArborescenceDoc(tv, _
                                     rs("D_DSNum").Value, _
                                     IMG_DOS, _
                                     IMG_DOS, _
                                     False) = P_ERREUR Then
            rs.Close
            afficher_documents = P_ERREUR
            Exit Function
        End If
        key = "D|" & rs("D_Num").Value & "|"
        stext = rs("D_Ident").Value & " / " & rs("D_Titre").Value
        Set nd = tv.Nodes("S" & rs("D_DSNum").Value)
'        nd.Sorted = True
        Call tv.Nodes.Add(nd, _
                           tvwChild, _
                           key, _
                           stext, _
                           IMG_DOC, _
                           IMG_DOC)
        nbdoc = nbdoc + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    
'    If nbdoc = 0 Then
'        MsgBox "Aucun document n'est disponible dans cette documentation.", vbInformation + vbOKOnly, ""
'        afficher_documents = P_NON
'        Exit Function
'    End If
    
    If g_toutes_les_docs > 0 And Not g_seldos_autor Then
        sql = "select DO_Num" _
            & " from Documentation"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_documents = P_ERREUR
            Exit Function
        End If
        While Not rs.EOF
            If TV_NodeExiste(tv, "O" & rs("DO_Num").Value, nd) = P_OUI Then
                If nd.Children = 0 Then
                    tv.Nodes.Remove nd.Index
                End If
            End If
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    If tv.Nodes.Count = 0 Then
        Call MsgBox("Aucun document à afficher.", vbExclamation + vbOKOnly, "")
        Call quitter
        afficher_documents = P_ERREUR
        Exit Function
    End If
    
    Set tv.SelectedItem = tv.Nodes(1).Root
    tv.SetFocus
    
    afficher_documents = P_OUI

End Function

Private Sub initialiser(ByVal v_param As String)

    If g_seldos_autor Then
        frm.Caption = "Choix d'un dossier/document"
    Else
        frm.Caption = "Choix d'un document"
    End If
    
    If afficher_documents() <> P_OUI Then
        Call quitter
        Exit Sub
    End If
    
    Call maj_boutons
    g_mode_saisie = True
    
End Sub

Private Sub maj_boutons()

    If Not g_seldos_autor Then
        If left$(tv.SelectedItem.key, 1) = "D" Then
            cmd(CMD_OK).Visible = True
        Else
            cmd(CMD_OK).Visible = False
        End If
    End If
    
End Sub

Private Sub quitter()

    g_schx = ""
    Unload Me
    
End Sub

Private Sub selectionner()

    Dim img As Long
    
    If g_plusieurs Then
        If tv.SelectedItem.image = IMG_DOC_SEL Then
            img = IMG_DOC
        Else
            img = IMG_DOC_SEL
        End If
        tv.SelectedItem.image = img
        tv.SelectedItem.SelectedImage = img
    Else
        Call valider
    End If
    
End Sub

Private Sub valider()

    Dim i As Integer, n As Integer
    Dim nd As Node
    
    If g_plusieurs Then
        Erase p_tabdoc_present()
        n = -1
        For i = 1 To tv.Nodes.Count
            Set nd = tv.Nodes(i)
            If left$(nd.key, 1) = "D" Then
                If nd.image = IMG_DOC_SEL Then
                    n = n + 1
                    ReDim Preserve p_tabdoc_present(n) As Long
                    p_tabdoc_present(n) = STR_GetChamp(nd.key, "|", 1)
                End If
            End If
        Next i
        If n = -1 Then
            ReDim Preserve p_tabdoc_present(0) As Long
            p_tabdoc_present(0) = STR_GetChamp(tv.SelectedItem.key, "|", 1)
        End If
        g_schx = "1"
    Else
        If left$(tv.SelectedItem.key, 1) = "D" Then
            g_schx = "D" & STR_GetChamp(tv.SelectedItem.key, "|", 1)
        Else
            g_schx = tv.SelectedItem.key
        End If
    End If
    Unload Me

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_QUITTER
        Call quitter
    End Select
    
End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub
    
    g_form_active = True
    Call initialiser(Me.tag)
    
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

Private Sub Form_Load()

    g_form_active = False
    g_mode_saisie = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub

Private Sub tv_BeforeLabelEdit(Cancel As Integer)

    Cancel = True
    
End Sub

Private Sub tv_Click()

    If left$(tv.SelectedItem.key, 1) = "D" Then
        If g_plusieurs Then Call selectionner
    End If
    
End Sub

Private Sub tv_DblClick()

    If left$(tv.SelectedItem.key, 1) = "D" Then
        Call selectionner
    End If
    
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If left$(tv.SelectedItem.key, 1) = "D" Then
            KeyCode = 0
            Call selectionner
        End If
    ElseIf KeyCode = vbKeySpace Then
        If g_plusieurs And left$(tv.SelectedItem.key, 1) = "D" Then
            KeyCode = 0
            Call selectionner
        End If
    End If
    
End Sub

Private Sub tv_NodeClick(ByVal Node As ComctlLib.Node)

    Call maj_boutons
    
End Sub


