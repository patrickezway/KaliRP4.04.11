VERSION 5.00
Begin VB.Form ExportSpecial 
   BackColor       =   &H00808000&
   Caption         =   "KaliWeb"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Intégration de données vers Excel"
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
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton ComOuvrirModele 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ouvrir le Modèle Bis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   5715
      End
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   1920
         Picture         =   "ExportSpecial.frx":0000
         ScaleHeight     =   1395
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton ComResultats 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Accès aux fichiers résultats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   5715
      End
      Begin VB.CommandButton ComOuvrirModele 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ouvrir le Modèle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4560
         Visible         =   0   'False
         Width           =   5715
      End
   End
   Begin VB.Menu MnuQuitter 
      Caption         =   "Quitter"
   End
   Begin VB.Menu MnuAPropos 
      Caption         =   "?"
      Begin VB.Menu MnuVersion 
         Caption         =   "Version"
      End
      Begin VB.Menu MnuBaseKD 
         Caption         =   "Base KaliDoc = "
      End
      Begin VB.Menu MnuFichierIni 
         Caption         =   "Fichier Ini ="
      End
   End
End
Attribute VB_Name = "ExportSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private g_form_active As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Function PrepareDossierResultat(v_path)
   Dim Chemin_Resultats As String
Test_Path:
   If FICH_EstRepertoire(v_path, False) Then
      Chemin_Resultats = v_path
      If Not FICH_EstRepertoire(Chemin_Resultats, False) Then
         MkDir (Chemin_Resultats)
         GoTo Test_Path
      Else
         Chemin_Resultats = v_path & "\"
         If Not FICH_EstRepertoire(Chemin_Resultats, False) Then
            MkDir (Chemin_Resultats)
            GoTo Test_Path
         End If
      End If
   Else
      MkDir (v_path)
      GoTo Test_Path
   End If
   PrepareDossierResultat = True
End Function

Private Function ControleFichierExterne(TabFichier(), CheminFichier, StrTitre, StrFeuille)
   Dim i As Integer
   Dim LaDim As Integer
   Dim frm As Form
   Dim NomFich As String
   Dim NomXLS As String
   Dim ret As Integer
   
   LaDim = 0
   ControleFichierExterne = -1
   On Error GoTo Err_ControleFichierExterne
   For i = 1 To UBound(TabFichier(), 2)
      If StrTitre = TabFichier(1, i) Then
         ControleFichierExterne = i
         Exit For
      End If
   Next i
   If ControleFichierExterne = -1 Then
      LaDim = UBound(TabFichier(), 2)
      GoTo TestFichier
   End If
   GoTo Fin_Err_ControleFichierExterne:
Err_ControleFichierExterne:
   If Err = 9 Then
      'Resume Next
TestFichier:
      ret = MsgBox("Votre paramétrage fait référence à un fichier dénommé : " & StrTitre & Chr(13) & Chr(10) & "Voulez vous le choisir", vbQuestion + vbYesNo)
      If ret = vbYes Then
         Set frm = Com_ChoixFichier
         NomFich = Com_ChoixFichier.AppelFrm("Choix du Modèle Excel", "c:", CheminFichier, "*.xls", False)
         Set frm = Nothing
         If NomFich = "" Then
            ControleFichierExterne = -2
            Exit Function
         End If
      Else
         ControleFichierExterne = -2
         Exit Function
      End If
      CheminFichier = NomFich
      NomXLS = Excel_OuvrirDoc(CheminFichier, "", Exc_wrk, False)
      
      If FICH_FichierExiste(CheminFichier) Then
         LaDim = LaDim + 1
         ReDim Preserve TabFichier(4, LaDim)
         TabFichier(1, LaDim) = Trim(StrTitre)
         TabFichier(2, LaDim) = Trim(CheminFichier)
         TabFichier(3, LaDim) = ""
         TabFichier(4, LaDim) = NomXLS
         ControleFichierExterne = LaDim
      End If
   Else
      MsgBox "ControleFichierExterne : " & Err & " " & Error$
   End If
Fin_Err_ControleFichierExterne:
   On Error GoTo 0
End Function

Private Sub ComOuvrirModele_Click(Index As Integer)
    Dim frm As Form
    Dim bcr As Boolean
    
    Set frm = PiloteExcelBis
    bcr = PiloteExcelBis.AppelFrm(1)
    'Set frm = PiloteExcel
    'bcr = PiloteExcel.AppelFrm(1)
    Set frm = Nothing


End Sub

Private Function init_param_exe(ByVal v_scmd As String, _
                                ByRef r_numutil As Integer, _
                                ByRef r_direct As Boolean) As Integer
                                  
    Dim nom_bdd As String
    Dim numutil As String
    Dim sql As String, rs As rdoResultset
    
    'Dim saction As String, snumdos As String, tbldoscli() As String, snumcli As String
    'Dim snumdosp As String, titredos As String, lstresp As String
    'Dim etat As Boolean
    Dim nbprm As Integer, n As Integer, i As Integer
    'Dim frm As Form
    'Dim rs As rdoResultset
    
    ' 1 Chemin Export
    ' 2 Nom BDD
    ' 3 Chemin Application
    ' 4 Num Utilisateur
    nbprm = STR_GetNbchamp(v_scmd, ";")
    If nbprm < 3 Then
        Call MsgBox("Usage : Export <Chemin Export>;<Nom BDD>;<Chemin Appli>" & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    
    ' 1- Chemin export : p_CheminExportSpecial
    p_CheminExportSpecial = STR_GetChamp(v_scmd, ";", 0)
    If p_CheminExportSpecial = "" Then
        Call MsgBox("<Chemin Export> est vide." & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "Usage : KaliDoc <Chemin Export>;<Nom BDD>;<Chemin Appli>", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    Else
        Me.MnuFichierIni.Caption = "Chemin export = " & p_CheminExportSpecial
    End If
    ' 2- Nom de la base
    nom_bdd = STR_GetChamp(v_scmd, ";", 1)
    'nom_bdd = "kalidoc_DCSS_DCSS"
    If nom_bdd = "" Then
        Call MsgBox("Pas de nom de base.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    Else
        Me.MnuBaseKD.Caption = "Base KaliDoc = " & nom_bdd
    End If
    ' 3- Chemin application : p_chemin_appli
    p_chemin_appli = STR_GetChamp(v_scmd, ";", 2)
    If p_chemin_appli = "" Then
        Call MsgBox("<Chemin Appli> est vide." & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "Usage : KaliDoc <Chemin Export>;<Nom BDD>;<Chemin Appli>", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    
    ' Connexion à la base
    If Odbc_Init("PG", nom_bdd) = P_ERREUR Then
        init_param_exe = P_ERREUR
        Exit Function
    End If
    
    sql = "select APP_Num from Application" _
        & " where APP_Code='KALIDOC'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If Not rs.EOF Then
        p_appli_kalidoc = rs("APP_Num").Value
    Else
        MsgBox "Impossible de déterminer p_appli_kalidoc"
        Exit Function
    End If
    rs.Close
    
    ' 5- Numéro d'utilisateur
    numutil = STR_GetChamp(v_scmd, ";", 4)
    If numutil <> "" Then
        r_numutil = val(numutil)
    End If
    
lab_fin:
    init_param_exe = P_OK
    
End Function

Private Sub Form_Load()
    Dim scmd As String
    Dim direct As Boolean
    Dim stype_bdd As String, nom_bdd As String
    Dim P_Version As String
    Dim nbParam As Integer
    Dim ret As Integer
    Dim Param1 As String, Param2 As String
    Dim sql As String, rs As rdoResultset
    
    scmd = Command$
    'scmd = "C:\KaliDoc\publiweb\Excel\ExportSpecial;kalidoc;C:\KaliDoc\;31;73"
    
    P_Version = "1.02 du 30/04/2008"
    
    Me.MnuVersion.Caption = "Version " & P_Version & ""
    
Début:
    nbParam = STR_GetNbchamp(scmd, ";")
    If nbParam < 3 Then
        MsgBox ("Paramètres absents <Chemin Export> ; <Nom BDD> ; <Chemin Application> ; <Numéro Utilisateur>")
        End
    End If
    
    ' 1 Chemin Export
    ' 2 Nom BDD
    ' 3 Chemin Application
    ' 4 Num Utilisateur
    p_NumUtil = 0
    If init_param_exe(scmd, p_NumUtil, direct) = P_ERREUR Then
        scmd = InputBox("Initialisation", "KaliTech", scmd)
        If scmd = "" Then End
        GoTo Début
    Else
        If p_NumUtil = 0 Then
            ret = P_SaisirUtilIdent(10, 10, 10, 10)
            If ret <> P_OUI Then
                MsgBox "Utilisateur non autorisé"
                End
            End If
        Else
            p_CodeUtil = AppelDirect(p_NumUtil)
        End If
    End If
    
    If p_CodeUtil <> "ROOT" Then
        sql = "select * from utilisateur where u_num = " & p_NumUtil
        If Odbc_SelectV(sql, rs) <> P_ERREUR Then
            If Not rs.EOF Then
                Me.Caption = Me.Caption & "   (" & LCase(rs("u_prenom")) & " " & UCase(rs("u_nom")) & ")"
            Else
                MsgBox "Utilisateur non autorisé"
                End
            End If
        End If
    Else
        Me.Caption = Me.Caption & "   (SuperUtilisateur)"
    End If

    p_Bool_Modif_Resolution = False
    
    TesterResolutionEcran

    Call FRM_ResizeForm(Me, Me.Width, Me.Height)

End Sub
Private Sub TesterResolutionEcran()
    Dim Anc_ScreenResolution As String
    
    Anc_ScreenResolution = ScreenResolution(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    
    Me.Caption = Me.Caption & "  (" & Anc_ScreenResolution & ")"
    If pAnc_Largeur_Ecran <> pNew_Largeur_Ecran Or pAnc_Hauteur_Ecran <> pNew_Hauteur_Ecran Then
        MsgBox "La résolution de votre écran va être adaptée"
    
        Call ResolutionEcran(pNew_Largeur_Ecran, pNew_Hauteur_Ecran)
        
        p_Bool_Modif_Resolution = True
    End If
End Sub

Private Function AppelDirect(ByVal v_unum As Integer) As String
    Dim sql As String, rs As rdoResultset
    Dim bad_util As Boolean
    
    sql = "select UAPP_Code, UAPP_MotPasse from Utilisateur, UtilAppli" _
        & " where UAPP_APPNum=" & p_appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_Num=UAPP_UNum" _
        & " and U_Num=" & v_unum
    'MsgBox sql
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        AppelDirect = 0
        Exit Function
    End If
    AppelDirect = 0
    If rs.EOF Then
        bad_util = True
    Else
        If rs("UAPP_Code").Value <> "" Then
            AppelDirect = UCase(rs("UAPP_Code").Value)
        End If
    End If
    rs.Close
    
End Function
Private Function FctOuvrirFichier(ByVal v_CheminFichier As String)
    
    If FICH_FichierExiste(v_CheminFichier) Then
        StartProcess v_CheminFichier
    Else
        MsgBox "Fichier " & v_CheminFichier & " introuvable"
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Unload
    exc_obj.Quit
    Set exc_obj = Nothing
    
    If p_Bool_Modif_Resolution Then
        Call ResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    End If

Err_Unload:
   Resume Next
End Sub

Private Sub MnuQuitter_Click()
   On Error GoTo Err_Unload
   exc_obj.Quit
   Set exc_obj = Nothing
   End
Err_Unload:
   Resume Next
End Sub


