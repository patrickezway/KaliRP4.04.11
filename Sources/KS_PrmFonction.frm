VERSION 5.00
Begin VB.Form KS_PrmFonction 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "     Fonction"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   390
         Index           =   4
         Left            =   7140
         Picture         =   "KS_PrmFonction.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Personnes ayant cette fonction"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Index           =   3
         Left            =   6525
         Picture         =   "KS_PrmFonction.frx":047F
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Services avec cette fonction"
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Un acteur avec cette fonction déclenchera l'inhibition du nom de l'acteur dans les documents"
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
         Left            =   390
         TabIndex        =   1
         Top             =   1590
         Width           =   6330
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1290
         MaxLength       =   80
         TabIndex        =   0
         Top             =   960
         Width           =   5415
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "KS_PrmFonction.frx":0A0E
         Top             =   10
         Width           =   300
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intitulé"
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
         Left            =   420
         TabIndex        =   7
         Top             =   990
         Width           =   615
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   800
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   7755
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KS_PrmFonction.frx":0E68
         Height          =   510
         Index           =   2
         Left            =   3270
         Picture         =   "KS_PrmFonction.frx":13F7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Supprimer la fonction"
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
         Left            =   5910
         Picture         =   "KS_PrmFonction.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KS_PrmFonction.frx":1F45
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
         Left            =   510
         Picture         =   "KS_PrmFonction.frx":24A1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Enregistrer les modifications"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "KS_PrmFonction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entree : Rien en param direct / 0 pour création d'ailleurs
' Sortie : ZNum|Nom si création d'ailleurs

' Index des objets cmd
Private Const CMD_OK = 0
Private Const CMD_DETRUIRE = 2
Private Const CMD_QUITTER = 1
Private Const CMD_LISTE_SRV = 3
Private Const CMD_LISTE_PERS = 4

' Index des objets txt
Private Const TXT_NOM = 0

Private Const CHK_GROUPE = 0

' No fction en saisie (0 si nouveau)
Private g_numfct As Long

Private g_mode_creation As Boolean
Private g_sret As String

Private g_crfct_autor As Boolean

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String

Public Function AppelFrm(ByVal v_numFct As String) As String

    g_sret = v_numFct
    
    Me.Show 1
    
    AppelFrm = g_sret
    
End Function

Private Function afficher_fct(ByVal v_numFct As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    
    If v_numFct > 0 Then
        sql = "select * from FctTrav" _
            & " where FT_Num=" & v_numFct
        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_fct = P_ERREUR
            Exit Function
        End If
        g_numfct = v_numFct
        txt(TXT_NOM).Text = rs("FT_Libelle").Value
        txt(TXT_NOM).tag = txt(TXT_NOM).Text
        chk(CHK_GROUPE).Value = IIf(rs("FT_EstGroupe").Value, 1, 0)
        rs.Close
        cmd(CMD_DETRUIRE).Enabled = True
        cmd(CMD_LISTE_SRV).Visible = True
        cmd(CMD_LISTE_PERS).Visible = True
    Else
        txt(TXT_NOM).Text = ""
        chk(CHK_GROUPE).Value = 0
        g_numfct = 0
        cmd(CMD_DETRUIRE).Enabled = False
        cmd(CMD_LISTE_SRV).Visible = False
        cmd(CMD_LISTE_PERS).Visible = False
    End If
    cmd(CMD_OK).Enabled = False
    
    txt(TXT_NOM).SetFocus
    Me.MousePointer = 0
    g_mode_saisie = True
    
    afficher_fct = P_OK
    
End Function

Private Function choisir_fct() As String

    Dim sret As String, sql As String
    Dim n As Integer
    Dim s As String
    Dim trouve As Boolean
    Dim i As Integer
    Dim nofct As Long
    Dim rs As rdoResultset
    
    Call FRM_ResizeForm(Me, 0, 0)

lab_affiche:
    Call CL_Init
    Call CL_InitMultiSelect(True, True)
    Call CL_AddLigne("=> Toutes les fonctions", 0, "", IIf(g_sret = "TOUTES", True, False))
    
    'Choix de la fonction
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_fct = P_ERREUR
        Exit Function
    End If
    n = 0
    If g_crfct_autor Then
        Call CL_AddLigne("<Nouvelle>", 0, "", False)
        n = 1
    End If
    While Not rs.EOF
        trouve = False
        For i = 0 To STR_GetNbchamp(g_sret, ";")
            s = STR_GetChamp(g_sret, ";", i)
            If s <> "" Then
                If s = rs("FT_Num").Value Then
                    trouve = True
                    Exit For
                End If
            End If
        Next i
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", trouve)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        Call MsgBox("Aucune fonction n'a été trouvée.", vbOKOnly + vbInformation, "")
        choisir_fct = P_NON
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Liste des fonctions", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_c_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    'Call CL_AddBouton("", p_chemin_appli + "\btnimprimer.gif", vbKeyI, vbKeyF3, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        choisir_fct = "-1"
        Exit Function
    End If
    
    '' Imprimer
    'If CL_liste.retour = 1 Then
    '    Call imprimer
    '    GoTo lab_affiche
    'End If
    
    sret = ""
    If CL_liste.lignes(0).selected Then
        sret = "0"
    Else
        For i = 1 To UBound(CL_liste.lignes())
            If CL_liste.lignes(i).selected Then
                If sret = "" Then
                    sret = CL_liste.lignes(i).num
                Else
                    sret = sret & ";" & CL_liste.lignes(i).num
                End If
            End If
        Next i
    End If
    choisir_fct = sret
'    choisir_fct = CL_liste.lignes(CL_liste.pointeur).num

End Function

Private Function enregistrer_fct() As Integer

    Dim lnb As Long
    
    If g_numfct = 0 Then
        If Odbc_AddNew("FctTrav", _
                        "FT_Num", _
                        "ft_seq", _
                        True, _
                        g_numfct, _
                        "FT_Libelle", txt(TXT_NOM).Text, _
                        "FT_EstGroupe", IIf(chk(CHK_GROUPE).Value = 1, True, False)) = P_ERREUR Then
            enregistrer_fct = P_ERREUR
            Exit Function
        End If
    Else
        If Odbc_Update("FctTrav", _
                        "FT_Num", _
                        "where FT_Num=" & g_numfct, _
                        "FT_Libelle", txt(TXT_NOM).Text, _
                        "FT_EstGroupe", IIf(chk(CHK_GROUPE).Value = 1, True, False)) = P_ERREUR Then
            enregistrer_fct = P_ERREUR
            Exit Function
        End If
        ' Chgt libellé des postes = libellé ancien fct
        If txt(TXT_NOM).Text <> txt(TXT_NOM).tag Then
            If Odbc_UpdateP("Poste", _
                            "PO_Num", _
                            "where PO_FTNum=" & g_numfct & " and PO_Libelle=" & Odbc_String(txt(TXT_NOM).tag), _
                            lnb, _
                            "PO_Libelle", txt(TXT_NOM).Text) = P_ERREUR Then
                enregistrer_fct = P_ERREUR
                Exit Function
            End If
        End If
    End If
    
    enregistrer_fct = P_OK
    
End Function

Private Function fct_dans_dest_do() As Integer

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like '%F" & g_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like '%F" & g_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = P_OUI
        Exit Function
    End If
    
    sql = "select count(*) from Document" _
        & " where D_Dest like '%F" & g_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = P_OUI
        Exit Function
    End If
    
    fct_dans_dest_do = P_NON
    
End Function

Private Function fct_dans_poste() As Integer

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Poste" _
        & " where PO_FTNum=" & g_numfct
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_poste = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_poste = P_OUI
        Exit Function
    End If
    
End Function

Private Function fct_dans_util() As Integer

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Utilisateur" _
        & " where U_FctTrav like '%F" & g_numfct & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_util = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_util = P_OUI
        Exit Function
    End If
    
    fct_dans_util = P_NON
    
End Function

Private Sub imprimer()

End Sub

Private Sub initialiser()

    g_crfct_autor = False
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    g_mode_saisie = False
    
    cmd(CMD_OK).Visible = False
    cmd(CMD_DETRUIRE).Visible = False
    
    'If g_mode_creation Then
    '    If afficher_fct(0) = P_ERREUR Then
    '        Call quitter(True)
    '        Exit Sub
    '    End If
    'Else
    g_sret = choisir_fct()
    Unload Me
    'End If
    
End Sub

Private Sub lister_personnes()

    Dim sql As String
    Dim n As Integer
    Dim rs As rdoResultset
    
    Call CL_Init
    
    sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
        & " where U_FctTrav like '%F" & g_numfct & ";%'" _
        & " order by U_Nom, U_Prenom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("U_Nom").Value & " " & rs("U_Prenom").Value, rs("U_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Call MsgBox("Aucune personne n'a cette fonction.", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    
    Call CL_InitTitreHelp("Liste des personnes ayant cette fonction", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_c_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1

End Sub

Private Sub lister_services()

    Dim sql As String
    Dim n As Integer
    Dim rs As rdoResultset
    
    Call CL_Init
    
    sql = "select SRV_Num, SRV_Nom from Poste, Service" _
        & " where PO_FTNum=" & g_numfct _
        & " and SRV_Num=PO_SRVNum" _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    n = 0
    While Not rs.EOF
        Call CL_AddLigne(rs("SRV_Nom").Value, rs("SRV_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Call MsgBox("Cette fonction n'est associée à aucun service.", vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    
    Call CL_InitTitreHelp("Liste des services avec cette fonction", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_c_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1

End Sub

Private Function quitter(ByVal v_bforce As Boolean) As Boolean

    Dim reponse As Integer
    
    If Not v_bforce Then
        If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
            If g_numfct = 0 Then
                reponse = MsgBox("La création de cette fonction ne s'effectuera pas !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            Else
                reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            End If
            If reponse = vbNo Then
                quitter = False
                Exit Function
            End If
        End If
    Else
        g_mode_creation = True
    End If
    
    If g_mode_creation Then
        g_sret = ""
        Unload Me
        quitter = True
        Exit Function
    End If
    
    g_sret = choisir_fct()
    Unload Me
    
End Function

Private Sub supprimer()

    Dim reponse As Integer, cr As Integer
    Dim lnb As Long
    
    ' Utilisateur associé à cette fonction ?
    cr = fct_dans_util()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_OUI Then
        Call MsgBox("Des personnes sont associées à cette fonction." & vbLf & vbCr & "Cette fonction ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, "")
        Exit Sub
    End If
    
    If p_appli_kalidoc > 0 Then
        cr = fct_dans_dest_do()
        If cr = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        If cr = P_OUI Then
            Call MsgBox("Des documentations/dossiers ou documents ont cette fonction comme destinataires." & vbLf & vbCr & "Cette fonction ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, "")
            Exit Sub
        End If
        cr = fct_dans_poste()
        If cr = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        If cr = P_OUI Then
            Call MsgBox("Des postes sont associés à cette fonction." & vbCrLf & "Cette fonction ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, "")
            Exit Sub
        End If
    End If
    
    reponse = MsgBox("Confirmez-vous la suppression de cette fonction ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        Exit Sub
    End If
    
    ' Maj table
    If Odbc_Delete("FctTrav", _
                   "FT_Num", _
                    "where FT_Num=" & g_numfct, _
                    lnb) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    
    If g_mode_creation Then
        g_sret = g_numfct & "|" & txt(TXT_NOM).Text
        Unload Me
        Exit Sub
    End If
    If choisir_fct() <> P_OUI Then Call quitter(True)
    
End Sub

Private Sub valider()

    Dim cr As Integer
    
    cr = verif_tous_chp()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_NON Then
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    cr = enregistrer_fct()
    Me.MousePointer = 0
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If g_mode_creation Then
        g_sret = g_numfct & "|" & txt(TXT_NOM).Text
        Unload Me
        Exit Sub
    End If
    
    If choisir_fct() <> P_OUI Then Call quitter(True)
    
End Sub

Private Function verif_code() As Integer

    Dim lib As String, sql As String
    Dim rs As rdoResultset
    
    lib = txt(TXT_NOM).Text
    If lib <> "" Then
        sql = "select FT_Num from FctTrav" _
            & " where FT_Libelle=" & Odbc_String(lib)
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            verif_code = P_ERREUR
            Exit Function
        End If
        If Not rs.EOF Then
            If rs("FT_Num").Value <> g_numfct Then
                rs.Close
                Call MsgBox("Fonction déjà existante.", vbOKOnly + vbExclamation, "")
                verif_code = P_NON
                Exit Function
            End If
        End If
        rs.Close
    End If
    
    verif_code = P_OUI

End Function

Private Function verif_tous_chp() As Integer

    If txt(TXT_NOM).Text = "" Then
        Call MsgBox("L' INTITULE de la fonction est une rubrique obligatoire.", vbOKOnly + vbExclamation, "")
        txt(TXT_NOM).SetFocus
        verif_tous_chp = P_NON
        Exit Function
    End If
    verif_tous_chp = P_OUI

End Function
    
Private Sub chk_Click(Index As Integer)

    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_DETRUIRE
        Call supprimer
    Case CMD_LISTE_SRV
        Call lister_services
    Case CMD_LISTE_PERS
        Call lister_personnes
    Case CMD_QUITTER
        Call quitter(False)
    End Select
    
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

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
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_DETRUIRE).Enabled Then
            Call supprimer
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_c_fonction.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
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

Private Sub txt_Change(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_avant = txt(Index).Text
    
End Sub

Private Sub txt_lostfocus(Index As Integer)

    Dim cr As Integer
    
    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            If Index = TXT_NOM Then
                cr = verif_code()
                If cr = P_ERREUR Then
                    Call quitter(True)
                    Exit Sub
                End If
                If cr = P_NON Then
                    txt(Index).Text = g_txt_avant
                    txt(Index).SetFocus
                    Exit Sub
                End If
            End If
            cmd(CMD_OK).Enabled = True
        End If
    End If
    
End Sub
