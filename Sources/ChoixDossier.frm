VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ChoixDossier 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix d'un dossier"
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
      Height          =   7545
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11910
      Begin ComctlLib.TreeView tvs 
         Height          =   6495
         Left            =   120
         TabIndex        =   2
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
         Left            =   0
         Top             =   0
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
               Picture         =   "ChoixDossier.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDossier.frx":06C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDossier.frx":085C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixDossier.frx":0E96
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
      Top             =   7410
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
         Left            =   720
         Picture         =   "ChoixDossier.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sélectionner"
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
         Index           =   0
         Left            =   10380
         Picture         =   "ChoixDossier.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Annuler l'opération"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixDossier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OK = 1
Private Const CMD_QUITTER = 0

Private Const IMG_DOCS = 1
Private Const IMG_DOS = 2

Private g_afficher_docs_encours As Integer
Private g_numdos As Long
Private g_numdosinit As Long

Private g_mode_saisie As Boolean
Private g_form_active As Boolean

' v_afficher_docs_encours : "0" pour afficher uniquement les autres documentations
'                           "1" pour afficher toutes les documentations
'                           "2" pour la demande de creation d'un document : Affiche uniquement la documentation en cours

Public Function AppelFrm(ByVal v_afficher_docs_encours As Integer, _
                            Optional v_numdosinit As Long) As Long
                             
    g_afficher_docs_encours = v_afficher_docs_encours
    g_numdosinit = v_numdosinit
    Me.Show 1
    
    AppelFrm = g_numdos
    
End Function

Private Function afficher_dossiers() As Integer
    
    Dim sql As String, s As String, titre As String
    Dim afficher As Boolean
    Dim n As Integer, i As Integer
    Dim numdocs_i As Long
    Dim sresp As Variant
    Dim nds As Node
    Dim rs As rdoResultset
    
    tvs.Nodes.Clear
    
    ' Cas ou ce n'est pas une demande de creation de document
    If g_afficher_docs_encours < 2 Then
        sql = "select DO_Num from Documentation" _
            & " where DO_Intranet=true"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            afficher_dossiers = P_ERREUR
            Exit Function
        End If
        If rs.EOF Then
            numdocs_i = 0
        Else
            numdocs_i = rs("DO_Num").Value
        End If
        rs.Close
    End If

    sql = "select DS_Num, DS_Titre, DS_DONum, DS_LstResp" _
        & " from Dossier" _
        & " where DS_LstResp like '%U" & p_NumUtil & ";%'"
    If numdocs_i <> 0 Then
        sql = sql & " and DS_DONum<>" & numdocs_i
    End If
    If g_afficher_docs_encours = 0 Then
        sql = sql & " and DS_DONum<>" & p_NumDocs
    End If
        
    'Cas suite a une demande de creation de document
    If g_afficher_docs_encours = 2 Then
        sql = sql & " and DS_DONum =" & p_NumDocs
    End If
    
    sql = sql & " order by DS_DONum, DS_NumPere, DS_Ordre"
    
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_dossiers = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Il n'y a aucun dossier dont vous êtes responsable.", vbInformation + vbOKOnly, "")
        Call quitter
        afficher_dossiers = P_OK
        Exit Function
    End If
        
    While Not rs.EOF
        sresp = rs("DS_LstResp").Value
        n = STR_GetNbchamp(sresp, "|")
        afficher = False
        For i = 0 To n - 1
            s = STR_GetChamp(sresp, "|", i)
            If CLng(Mid$(STR_GetChamp(s, ";", P_DODS_RESP_NUMUTIL), 2)) = p_NumUtil Then
                If STR_GetChamp(s, ";", P_DODS_RESP_CRAUTOR) = 1 Then
                    afficher = True
                End If
                Exit For
            End If
        Next i
        If afficher Then
            If TV_NodeExiste(tvs, "O" & rs("DS_DONum").Value, nds) = P_NON Then
                If Odbc_RecupVal("select DO_Titre from Documentation where DO_Num=" & rs("DS_DONum").Value, _
                                 titre) = P_ERREUR Then
                    afficher_dossiers = P_ERREUR
                    Exit Function
                End If
                Set nds = tvs.Nodes.Add(, tvwChild, "O" & rs("DS_DONum").Value, titre, IMG_DOCS, IMG_DOCS)
                nds.Expanded = True
            End If
            Call P_AfficherArborescenceDoc(tvs, rs("DS_Num").Value, IMG_DOS, IMG_DOS, True)
            
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If tvs.Nodes.Count = 0 Then
        MsgBox "Aucun dossier dans lequel vous êtes autorisé à créer n'est disponible.", vbInformation + vbOKOnly, ""
        afficher_dossiers = P_NON
        Exit Function
    End If
    
    ' On selectionne le dossier choisi par le demandeur
    If g_numdosinit > 0 Then
        tvs.SelectedItem = tvs.Nodes("S" & g_numdosinit & "")
    
        SendKeys "{UP}"
        SendKeys "{DOWN}"
        DoEvents
    Else
        tvs.SelectedItem = tvs.Nodes(1)
    End If
    
    tvs.SetFocus
    
    afficher_dossiers = P_OUI

End Function

Private Sub initialiser()

    If afficher_dossiers() <> P_OUI Then
        Call quitter
        Exit Sub
    End If
    
    g_mode_saisie = True
    
End Sub

Private Sub quitter()

    g_numdos = 0
    
    Unload Me
    
End Sub

Private Sub valider()

    If left$(tvs.SelectedItem.key, 1) = "S" Then
        g_numdos = Mid$(tvs.SelectedItem.key, 2)
        Unload Me
    End If
    
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

Private Sub Form_Load()

    g_form_active = False
    g_mode_saisie = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub

Private Sub tvs_BeforeLabelEdit(Cancel As Integer)

    Cancel = True
    
End Sub

Private Sub tvs_DblClick()

    If left$(tvs.SelectedItem.key, 1) = "S" Then
        Call valider
    End If
    
End Sub

Private Sub tvs_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If left$(tvs.SelectedItem.key, 1) = "S" Then
            KeyCode = 0
            Call valider
        End If
    End If
    
End Sub




