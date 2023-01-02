VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form LanceRapportType 
   BackColor       =   &H00808000&
   Caption         =   "Versionning de Rapport Type"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmHTTPD 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   7695
      Begin ComctlLib.ProgressBar PgbarHTTPDTemps 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar PgbarHTTPDTaille 
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblHTTPDTaille 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblHTTPDTemps 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblHTTPD 
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
         TabIndex        =   4
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   0
         Left            =   360
         Picture         =   "LanceRapportType.frx":0000
         ScaleHeight     =   975
         ScaleWidth      =   3135
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1485
         Index           =   1
         Left            =   3240
         Picture         =   "LanceRapportType.frx":A9DA
         ScaleHeight     =   1485
         ScaleWidth      =   1815
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Menu MnuQuitter 
      Caption         =   "Quitter"
   End
   Begin VB.Menu MnuAide 
      Caption         =   "Aide"
   End
   Begin VB.Menu MnuAPropos 
      Caption         =   "?"
      Begin VB.Menu MnuVersion 
         Caption         =   "Version"
      End
      Begin VB.Menu MnuFichierEvo 
         Caption         =   "Modifications de la version"
      End
      Begin VB.Menu MnuTrait 
         Caption         =   "---------------------------------------"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuBaseKD 
         Caption         =   "Base KaliDoc = "
      End
      Begin VB.Menu MnuCheminAppli 
         Caption         =   "Chemin application"
      End
      Begin VB.Menu MnuCheminMod_Serveur 
         Caption         =   "Chemin des Modèles (Serveur)"
      End
      Begin VB.Menu MnuCheminMod_Local 
         Caption         =   "Chemin des Modèles (Local)"
      End
      Begin VB.Menu MnuCheminRes 
         Caption         =   "Chemin des Résultats"
      End
      Begin VB.Menu MnuSVersConf 
         Caption         =   "S_Vers_Conf"
      End
      Begin VB.Menu MnuFichierIni 
         Caption         =   "Fichier Ini ="
      End
      Begin VB.Menu MnuScmd 
         Caption         =   "Commande"
      End
      Begin VB.Menu MnuHTTP 
         Caption         =   "HTTP"
         Begin VB.Menu mnuHTTPDConfig1 
            Caption         =   "1"
         End
         Begin VB.Menu mnuHTTPDConfig2 
            Caption         =   "2"
         End
      End
   End
End
Attribute VB_Name = "LanceRapportType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private g_form_active As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Function init_param_exe(ByVal v_scmd As String, _
                                ByRef r_numfor As Integer, _
                                ByRef r_numutil As Integer, _
                                ByRef r_nummodele As Integer, _
                                ByRef r_direct As Boolean) As Integer
    Dim numfor As String
    Dim nom_bdd As String
    Dim numutil As String
    Dim NumForm As String
    Dim NumModele As String
    Dim sql As String, rs As rdoResultset
    Dim nbprm As Integer, n As Integer, i As Integer
    
    'nbprm = STR_GetNbchamp(v_scmd, ";")
    'If nbprm < 3 Then
    '    init_param_exe = P_ERREUR
    '    GoTo ErrParametres
    'End If
    
    ' 1- Chemin export : p_CheminRapportType
    p_CheminRapportType = STR_GetChamp(v_scmd, ";", 0)
    If p_bool_ModeDebug Then MsgBox "p_CheminRapportType=" & p_CheminRapportType
    If p_CheminRapportType = "" Then
        init_param_exe = P_ERREUR
        GoTo ErrParametres
    Else
        Me.MnuCheminAppli.Caption = "Chemin Application = " & p_CheminRapportType
    End If
    
    ' 2- Nom Fichier Ini : p_CheminRapportType_Ini
    p_CheminRapportType_Ini = STR_GetChamp(v_scmd, ";", 1)
    If p_bool_ModeDebug Then MsgBox "p_CheminRapportType_Ini=" & p_CheminRapportType_Ini
    If p_CheminRapportType_Ini = "" Then
        init_param_exe = P_ERREUR
        GoTo ErrParametres
    Else
        p_CheminRapportType_Ini = p_CheminRapportType & "\" & p_CheminRapportType_Ini
    End If
    
    ' 3- Nom de la base
    nom_bdd = STR_GetChamp(v_scmd, ";", 2)
    If p_bool_ModeDebug Then MsgBox "nom_bdd=" & nom_bdd
    If nom_bdd = "VIDE" Or nom_bdd = "" Then
    Else
        Me.MnuBaseKD.Caption = "Base KaliDoc = " & nom_bdd
        p_nomBDD_ODBC = nom_bdd
    End If
    
    ' 4- Chemin application : p_chemin_appli
    p_chemin_appli = STR_GetChamp(v_scmd, ";", 3)
    If p_bool_ModeDebug Then MsgBox "p_chemin_appli=" & p_chemin_appli
    If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
    End If
    
    ' Connexion à la base
    If nom_bdd <> "" And nom_bdd <> "VIDE" Then
        If Odbc_Init("PG", nom_bdd) = P_ERREUR Then
            init_param_exe = P_ERREUR
            Exit Function
        End If
    End If
    
    ' 5- Numéro de Formulaire
    numfor = STR_GetChamp(v_scmd, ";", 4)
    If numfor <> "" Then
        r_numfor = val(numfor)
    End If
    
    ' 6- Numéro d'utilisateur
    numutil = STR_GetChamp(v_scmd, ";", 5)
    If numutil <> "" Then
        r_numutil = val(numutil)
    End If

    ' 7- Numéro de modèle
    NumModele = STR_GetChamp(v_scmd, ";", 6)
    If NumModele <> "" Then
        r_nummodele = val(NumModele)
    End If

Lab_Fin:
    init_param_exe = P_OK
    Exit Function
            
ErrParametres:
    Call MsgBox("Usage : RapportType <Chemin Export>;<Nom fichier Ini>;<Nom BDD>;<Chemin Appli KaliDoc>;<NumFormulaire>;<NumUtil>;<NumModèle>" & vbCr & vbLf _
            & "cmd:" & v_scmd & vbCr & vbLf _
            & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
    init_param_exe = P_ERREUR
End Function

Private Function VerifVersion()
    Dim strOut As String
    Dim sql As String, rs As rdoResultset
    Dim CheminServeur As String, CheminLocal As String, Session As String
    Dim strHTTP As String, nomIn_Chemin As String, nomIn_Fichier As String, nomIn_Extension As String
    Dim nomInCpyExe As String, nomInCpyPg As String
    Dim nomInCpy As String
    Dim iRet As Integer
    Dim fp As Integer
    Dim url As String
    Dim util As String
    Dim cnd_sversconf As String
    Dim Cmd1 As String
    Dim Cmd2 As String
    Dim NewVersion As String
    Dim NomFichierUser As String
    Dim FichServeur As String
    Dim FichLocal As String
    Dim VersionActuelle As String
    Dim PremiereInstal As Boolean
    Dim i As Integer, scmd As String
    
    Me.Visible = True
    ' version actuelle du poste ?
    FichLocal = p_CheminRapportType & "/Version.txt"
    If Not FICH_FichierExiste(FichLocal) Then
        VersionActuelle = ""
        PremiereInstal = True
    Else
        PremiereInstal = False
        If FICH_OuvrirFichier(FichLocal, FICH_LECTURE, fp) = P_ERREUR Then
            MsgBox "Impossible d'ouvrir " & FichLocal
        Else
            While Not EOF(fp)
                Line Input #fp, VersionActuelle
            Wend
            Close #fp
        End If
    End If
    
    ' Voir sur le serveur si un fichier "NewVersion.txt" existe
    Set p_HTTP_Form_Frame = LanceRapportType
    
    FichServeur = p_RapportTypeExe & "/NewVersion.txt"
    FichLocal = p_CheminRapportType & "/NewVersion.txt"
    If FICH_FichierExiste(FichLocal) Then
        Kill (FichLocal)
    End If
        
    ' chargement par HTTPD
    iRet = HTTP_Appel_getfile(HTTP_GET_LIB, FichServeur, FichLocal, False, False)
            
    p_HTTP_Form_Frame.FrmHTTPD.Visible = False
    
    If iRet = HTTP_GET_OK Or iRet = HTTP_GET_DEJA_EN_LOCAL Then
        If FICH_FichierExiste(FichLocal) Then
            ' ouvrir et comparer le numéro de version
            FichLocal = Replace(FichLocal, "_Session_" & Session, "")
            If FICH_OuvrirFichier(FichLocal, FICH_LECTURE, fp) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & FichLocal
                Exit Function
            End If
            While Not EOF(fp)
                Line Input #fp, NewVersion
            Wend
            Close #fp
        End If
    End If
    ' est ce la même version que l'exe en cours ?
    If NewVersion = VersionActuelle Then
        ' version déjà chargée sur ce poste
        GoTo Appel
    End If
        
    ' changement de version
    If Not Odbc_TableExiste("rapport_type_version") Then
        GoTo FaireQuandMême
    Else
        sql = "select * from rapport_type_version"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Version Serveur de RapportType introuvable"
            End
            Exit Function
        End If
    End If
    
    If rs("rpv_version") = NewVersion Then
        ' Nouvelle version déjà chargée par un poste
        ' on ne fait pas le pg
    Else
FaireQuandMême:
        
        MsgBox "Une mise à jour de la base RapportType va être faite sur le Serveur"

        ' chargement par HTTPD (voir si le php existe)
        FichServeur = p_RapportTypeExe & "/RapportType_MAJ.php"
        FichLocal = p_CheminRapportType & "/RapportType_MAJ.php"
        If FICH_FichierExiste(FichLocal) Then
            Kill (FichLocal)
        End If
        iRet = HTTP_Appel_getfile(HTTP_GET_LIB, FichServeur, FichLocal, False, False)
        If FICH_FichierExiste(FichLocal) Then
            Kill (FichLocal)
        End If
            
        p_HTTP_Form_Frame.FrmHTTPD.Visible = False
    
        If iRet = HTTP_GET_OK Or iRet = HTTP_GET_DEJA_EN_LOCAL Then
        Else
            MsgBox "Le fichier " & FichServeur & " est introuvable"
            Exit Function
        End If

        ' Charger le pg
        FichServeur = p_RapportTypeExe & "/RapportType_" & NewVersion & ".pg"
        FichLocal = p_CheminRapportType & "/RapportType_" & NewVersion & ".pg"
        If FICH_FichierExiste(FichLocal) Then
            Kill (FichLocal)
        End If
        
        ' chargement par HTTPD en lockant ou pas
        iRet = HTTP_Appel_getfile(HTTP_GET_LIB, FichServeur, FichLocal, False, False)
        
        p_HTTP_Form_Frame.FrmHTTPD.Visible = False
        
        If iRet = HTTP_GET_OK Or iRet = HTTP_GET_DEJA_EN_LOCAL Then
            FichLocal = Replace(FichLocal, "_Session_" & Session, "")

            If FICH_FichierExiste(FichLocal) Then
                Kill (FichLocal)
                ' la commande de mise à jour de la base
                url = "/RapportType/SERVEUR/RapportType_MAJ.php%3FV_Chemin=" & FichServeur
                url = url & "%26V_NewVersion=" & NewVersion
                url = url & "%26V_nomBDD=" & p_nomBDD_SERVEUR
                util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
                If p_S_Vers_Conf <> "" Then
                    cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
                End If
                url = "http:" & p_HTTP_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & url
                ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
                SYS_ExecShell "C:\Program Files\Internet Explorer\iexplore.exe " & url, True, True
            End If
        Else
            ' mettre à jour la base
            sql = "update rapport_type_version set rpv_version = '" & NewVersion & "'"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                MsgBox "Erreur lors de la mise à jour de la version"
            End If
        End If
    End If
    
    ' Charger la nouvelle version
    FichServeur = p_RapportTypeExe & "/RapportType.exe"
    FichLocal = p_CheminRapportType & "/RapportType_New.exe"
    If FICH_FichierExiste(FichLocal) Then
        Kill FichLocal
    End If
    
    ' chargement par HTTPD en lockant ou pas
    
    MsgBox "Votre version de RapportType doit être mise à jour"
    
    iRet = HTTP_Appel_getfile(HTTP_GET_LIB, FichServeur, FichLocal, False, False)
    
    p_HTTP_Form_Frame.FrmHTTPD.Visible = False
    
    If iRet = HTTP_GET_OK Or iRet = HTTP_GET_DEJA_EN_LOCAL Then
        nomInCpyExe = FichLocal
        ' copier l'actuel en .old
        If FICH_FichierExiste(Replace(nomInCpyExe, "_New", "")) Then
            FICH_CopierFichier Replace(nomInCpyExe, "_New", ""), Replace(nomInCpyExe, "_New", "_Old")
        End If
        ' effacer le fichier utilisateur (signe de nouvelle version pour ce poste)
        NomFichierUser = p_Drive_Modeles_Local & p_Chemin_Modeles_Local & "/User_" & p_NumUtil & ".User"
        If FICH_FichierExiste(NomFichierUser) Then
            Kill NomFichierUser
        End If

        If FICH_FichierExiste(nomInCpyExe) Then
            ' renommer le .exe
            FICH_CopierFichier nomInCpyExe, Replace(nomInCpyExe, "_New", "")
            ' Modifier la version
            FichLocal = p_CheminRapportType & "/Version.txt"
            If FICH_FichierExiste(FichLocal) Then
                Kill FichLocal
            End If
            fp = FreeFile
            If Not FICH_OuvrirFichier(FichLocal, FICH_ECRITURE, fp) = P_ERREUR Then
                Print #fp, NewVersion
                Close #fp
            End If
Appel:
            ' recomposer le Command$
            ' 1 - Chemin export       : p_CheminRapportType
            ' 2 - Nom du Fichier Ini  : p_CheminRapportType_Ini
            ' 3 - nom de la base ODBC : p_nomBDD_ODBC
            ' 4- Chemin application   : p_chemin_appli
            ' 5- Numéro de Formulaire : p_numform
            ' 6- Numéro d'utilisateur : p_numutil
            ' 7- Numéro de modèle     : p_NumModele
            'scmd = \\192.168.101.20\kalidoc\Sources_VB\RapportType;kalidoc;c:\kalidoc;102;73

            scmd = p_CheminRapportType & ";" & Replace(p_CheminRapportType_Ini, p_CheminRapportType & "\", "") & ";" & p_nomBDD_ODBC & ";" & p_chemin_appli & ";" & p_NumForm & ";" & p_NumUtil & ";" & p_NumModele
            If p_bool_ModeDebug Then InputBox "scmd=", "", scmd
            ' faire le ménage
            FichLocal = p_CheminRapportType & "/NewVersion.txt"
            If FICH_FichierExiste(FichLocal) Then
                Kill (FichLocal)
            End If
            FichLocal = p_CheminRapportType & "/RapportType_New.exe"
            If FICH_FichierExiste(FichLocal) Then
                Kill FichLocal
            End If
            
            SYS_ExecShell p_CheminRapportType & "\RapportType.exe " & scmd, False, True
            End
        End If
    Else
        MsgBox "Impossible de charger le fichier " & FichServeur & " à partir de " & p_HTTP_AdrServeur
    End If
Fin:
End Function

Private Sub Form_Load()
    Dim scmd As String
    Dim sret As String
    Dim direct As Boolean
    Dim stype_bdd As String, nom_bdd As String, nom_bddS As String
    Dim nbParam As Integer
    Dim ret As Integer
    Dim fd As Integer
    Dim param1 As String, param2 As String
    Dim sql As String, rs As rdoResultset
    Dim frm As Form, bcr As Boolean
    Dim Chemin_Résultats As String
    Dim NomFichierUser As String
    Dim fso As FileSystemObject
    Dim Dossier As Variant
    Dim fileItem As Variant
    Dim nbDossiers As Integer
    
    scmd = Command$
'    scmd = "DEBUG;C:\KaliDoc\RapportType;RapportTypeMULH_110.ini;kalidoc_mulhouse_ML110;C:\KaliDoc\;31;3;10"
    'scmd = "DEBUG;C:\KaliDoc\RapportType;RapportType_DCSS.ini;kalidoc_DCSS_DCSS;C:\KaliDoc\"
    'scmd = "C:\KaliDoc\RapportType;kalidoc_DCSS_DCSS;C:\KaliDoc\"
    'scmd = "C:\KaliDoc\RapportType;kalidoc;C:\KaliDoc\;31;73"
    'scmd = "C:\KaliDoc\RapportType;kalidoc;C:\KaliDoc\;31;73"
    'scmd = "C:\KaliDoc\RapportType;kalidoc;C:\KaliDoc\;;73"
    'scmd = "\\192.168.101.20\kalidoc\Sources_VB\RapportType;kalidoc;c:\kalidoc;102;73"
    'scmd = "C:\KaliDoc\Recupdu20\Dernier\RapportType;kalidoc;C:\KaliDoc\;31;73"
    'scmd = "C:\KaliDoc\RapportType;kalidoc_dcss_ML110;C:\KaliDoc\;;2"
    'scmd = "C:\KaliDoc\RapportType;kalidoc_dcss_vm110;C:\KaliDoc\;;2"
    'scmd = "C:\KaliDoc\RapportType;kalidoc_HMZ;C:\KaliDoc\;;54"
    'scmd = "C:\KaliDoc\RapportType;kalidemo_20;C:\KaliDoc\;;73;6"
    'scmd = "C:\KaliDoc\RapportType;VIDE;VIDE;15;73"
    'scmd = "C:\KaliDoc\RapportType;kalidemo_20;C:\KaliDoc\;50;73"
    'scmd = "C:/KaliDoc/RapportType;VIDE;VIDE;15;73;5"
    'scmd = "C:/KaliDoc/RapportType;VIDE;VIDE;;6;"
' scmd = "DEBUG;C:\KaliDoc\RapportType;RapportType.ini;kalidoc;C:\KaliDoc\"
    If scmd = "" Then
        scmd = "C:\KaliDoc\RapportType;RapportType.ini"
        GoTo Faire
    End If
    If InStr(scmd, "DEBUG;") > 0 Then
        scmd = Replace(scmd, "DEBUG;", "")
        p_bool_ModeDebug = True
    End If
    p_Command = scmd
    
    If p_bool_ModeDebug Then MsgBox "scmd en entrée=" & scmd
    
    Set p_HTTP_Form_Menu = LanceRapportType
    
Init:
    Me.MnuScmd.Caption = "Commande=" & scmd
    p_Version = "4_11_02C_00"
    
    Me.MnuVersion.Caption = "Version " & p_Version
    

Début:
    nbParam = STR_GetNbchamp(scmd, ";")
    If nbParam < 3 Then
        If Not FICH_EstRepertoire(scmd, False) Then
            p_CheminRapportType = scmd
            MsgBox ("Rapports Types Paramètres absents <Chemin Export> ; <Nom BDD> ; <Chemin Application> ; <Numéro Utilisateur>  " & Chr(13) & Chr(10) & "Commande = " & scmd)
            GoTo FaireEnAuto
        Else
            p_CheminRapportType = scmd
            If Not FICH_FichierExiste(p_CheminRapportType & "/RapportType.ini") Then
                MsgBox ("Rapports Types Paramètres absents <Chemin Export> ; <Nom BDD> ; <Chemin Application> ; <Numéro Utilisateur>  " & Chr(13) & Chr(10) & "Commande = " & scmd)
                GoTo Init_Ini
            Else
                GoTo IniExiste
            End If
        End If
    End If
    
    ' 1 Chemin Export
    ' 2 Nom BDD
    ' 3 Chemin Application
    ' 4 num formulaire
    ' 5 Num Utilisateur
Faire:
    p_NumUtil = 0
    If init_param_exe(scmd, p_NumForm, p_NumUtil, p_NumModele, direct) = P_ERREUR Then
        scmd = InputBox("Initialisation", "KaliTech", scmd)
        If scmd = "" Then End
        GoTo Début
    Else
        ' Voir si le dossier existe bien
FaireEnAuto:
        If Not FICH_EstRepertoire(p_CheminRapportType, False) Then
            ret = MsgBox("Le dossier " & p_CheminRapportType & " n'existe pas" & Chr(13) & Chr(10) & "Voulez vous le créer ?", vbDefaultButton1 + vbYesNo + vbQuestion)
            If ret = vbYes Then
                ret = FICH_CreerRepComp(p_CheminRapportType, False, False)
                If ret = 1 Then
                    ret = FICH_CreerRepComp(p_CheminRapportType & "\Fichiers", False, False)
                    ret = FICH_CreerRepComp(p_CheminRapportType & "\Modèles", False, False)
Init_Ini:
                    ' initialiser le .ini
                    If Not FICH_OuvrirFichier(p_CheminRapportType & "\RapportType.ini", FICH_ECRITURE, fd) = P_ERREUR Then
                        Print #fd, "[CHEMINS]"
                        Print #fd, "DRIVE_RESULTATS=C:"
                        Print #fd, "PATH_RESULTATS=\kalidoc\RapportType\Fichiers"
                        Print #fd, "DRIVE_MODELES=C:"
                        Print #fd, "PATH_MODELES=\kalidoc\RapportType\Modèles"
                        
                        Print #fd, "HTTP_SERVEUR=\\localhost"
            
                        Print #fd, "DRIVE_KALIDOC="
                        Print #fd, "PATH_KALIDOC="

Saisie_NomBDD:
                        sret = InputBox("nom de la base de données", "Initialisation", "kalidoc")
                        If sret = "" Then End
                        If Odbc_Init("PG", sret) = P_ERREUR Then
                            MsgBox "Base " & sret & " introuvable"
                            GoTo Saisie_NomBDD
                        End If
                        Print #fd, "Nom_BDD_ODBC=" & sret
                        
                        sql = "select APP_Num from Application" _
                            & " where APP_Code='KALIDOC'"
                        If Odbc_SelectV(sql, rs) = P_ERREUR Then
                            MsgBox "Chemin de l'application introuvable"
                            GoTo Fermer
                        End If
                        If Not rs.EOF Then
                            p_appli_kalidoc = rs("APP_Num").Value
                        Else
                            MsgBox "Impossible de déterminer p_appli_kalidoc"
                        End If
                        
                        Print #fd, "Chemin_Appli=C:\kalidoc"
                        
                        rs.Close
Fermer:
                        Close #fd
                    End If
                End If
            Else
                End
            End If
        End If
IniExiste:
        ' Voir si le fichier .ini existe
        If Not FICH_FichierExiste(p_CheminRapportType_Ini) Then
            
            param1 = MsgBox("Fichier ini : " & p_CheminRapportType_Ini & " est introuvable")
            Call MsgBox("Usage : RapportType <Chemin Export>;<Nom fichier Ini>;<Nom BDD>;<Chemin Appli KaliDoc>;<NumFormulaire>;<NumUtil>;<NumModèle>" & vbCr & vbLf _
                    & "cmd:" & p_Command & vbCr & vbLf _
                    & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
            End
            GoTo Début
        Else
            '
            'p_CheminRapportType_Ini = p_CheminRapportType & "/RapportType.ini"
            Me.MnuFichierIni.Caption = "Fichier Ini = " & p_CheminRapportType_Ini
            
            ' Récupérer l'adresse HTTP du serveur
            p_HTTP_AdrServeur = SYS_GetIni("Chemins", "HTTP_SERVEUR", p_CheminRapportType_Ini)
            If p_HTTP_AdrServeur = "" Then
                MsgBox "Adresse HTTP du serveur non renseignée"
            End If
            
            ' Récupérer l'adresse s_Vers_Conf
            p_S_Vers_Conf = SYS_GetIni("Chemins", "S_VERS_CONF", p_CheminRapportType_Ini)
            If p_S_Vers_Conf <> "" Then
                Me.MnuSVersConf.Caption = "s_vers_conf = " & p_S_Vers_Conf
            End If
            
            ' Récupérer le Chemin de Dépot HTTP du serveur
            p_HTTP_CheminDépot = SYS_GetIni("Chemins", "CHEMIN_DEPOT_HTTP", p_CheminRapportType_Ini)
            If p_HTTP_CheminDépot = "" Then
                MsgBox "Chemin de Dépot HTTP du serveur non renseignée"
            End If
            
            ' Récupérer le Chemin de destination des tableaux publiés dans KaliDoc
            p_Drive_KaliDoc = SYS_GetIni("Chemins", "DRIVE_KALIDOC", p_CheminRapportType_Ini)
            p_Path_KaliDoc = SYS_GetIni("Chemins", "PATH_KALIDOC", p_CheminRapportType_Ini)
            If p_Drive_KaliDoc & p_Path_KaliDoc = "" Then
                MsgBox "Chemin de Publication dans KaliDoc non renseignée"
            End If
                        
            If p_nomBDD_ODBC = "" Or p_nomBDD_ODBC = "VIDE" Then
                ' 2- Nom de la base
                nom_bdd = SYS_GetIni("Chemins", "Nom_BDD_ODBC", p_CheminRapportType_Ini)
                If nom_bdd = "" Then
                    Call MsgBox("Pas de nom de base (ODBC).", vbInformation + vbOKOnly)
                    End
                Else
                    Me.MnuBaseKD.Caption = "Base KaliDoc = " & nom_bdd
                    p_nomBDD_ODBC = nom_bdd
                End If
    
                ' Connexion à la base
                If Odbc_Init("PG", nom_bdd) = P_ERREUR Then
                    MsgBox "Connexion à la base " & nom_bdd & " impossible"
                    End
                End If
            End If
                
            If p_nomBDD_SERVEUR = "" Then
                ' 2- Nom de la base sur le serveur
                nom_bddS = SYS_GetIni("Chemins", "Nom_BDD_SERVEUR", p_CheminRapportType_Ini)
                If nom_bddS = "" Then
                    Call MsgBox("Pas de nom de base (Serveur).", vbInformation + vbOKOnly)
                    End
                Else
                    Me.MnuBaseKD.Caption = "Base KaliDoc ODBC = " & nom_bdd & " => Serveur=" & nom_bddS
                    p_nomBDD_SERVEUR = nom_bddS
                End If
            End If
                
            ' 3- Chemin application : p_chemin_appli
            If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
                p_chemin_appli = SYS_GetIni("Chemins", "Chemin_Appli", p_CheminRapportType_Ini)
                If p_chemin_appli = "" Or p_chemin_appli = "VIDE" Then
                    Call MsgBox("<Chemin Appli> est vide." & vbCr & vbLf _
                                & "Usage : KaliDoc <Chemin Export>;<Nom BDD>;<Chemin Appli>", vbInformation + vbOKOnly)
                    End
                End If
            End If
            
            sql = "select APP_Num from Application" _
                & " where APP_Code='KALIDOC'"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                MsgBox "Impossible de charger p_appli_kalidoc à partir de la base"
                End
            End If
            If Not rs.EOF Then
                p_appli_kalidoc = rs("APP_Num").Value
            Else
                MsgBox "Impossible de déterminer p_appli_kalidoc"
                End
            End If
            rs.Close
            
            ' Récupérer le chemin des Modèles sur le serveur
            p_Drive_Modeles_Serveur = SYS_GetIni("Chemins", "Drive_Modeles_Serveur", p_CheminRapportType_Ini)
            p_Chemin_Modeles_Serveur = SYS_GetIni("Chemins", "Path_Modeles_Serveur", p_CheminRapportType_Ini)
            Me.MnuCheminMod_Serveur.Caption = "Modèles => " & p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur
            If p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur = "" Then
                MsgBox "Chemin des Modèles non renseigné"
                End
            ' ElseIf Not FICH_EstRepertoire(p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur, False) Then
            '     MsgBox "Chemin des Modèles " & p_Drive_Modeles_Serveur & p_Chemin_Modeles_Serveur & " introuvable"
            End If
            
            ' Récupérer le chemin des Modèles en local
            p_Drive_Modeles_Local = SYS_GetIni("Chemins", "Drive_Modeles_Local", p_CheminRapportType_Ini)
            p_Chemin_Modeles_Local = SYS_GetIni("Chemins", "Path_Modeles_Local", p_CheminRapportType_Ini)
            Me.MnuCheminMod_Local.Caption = "Modèles => " & p_Drive_Modeles_Local & p_Chemin_Modeles_Local
            If p_Drive_Modeles_Local & p_Chemin_Modeles_Local = "" Then
                MsgBox "Chemin des Modèles non renseigné"
                End
            ElseIf Not FICH_EstRepertoire(p_Drive_Modeles_Local & p_Chemin_Modeles_Local, False) Then
                MsgBox "Chemin des Modèles " & p_Drive_Modeles_Local & p_Chemin_Modeles_Local & " introuvable"
            End If
            
            ' Récupérer le chemin de RapportType.exe sur le serveur
            p_RapportTypeExe = SYS_GetIni("Chemins", "RapportTypeExe", p_CheminRapportType_Ini)
            If p_RapportTypeExe = "" Then
                MsgBox "Chemin de RapportTypeExe non renseigné"
                End
            End If
            
            ' Récupérer le chemin des Résultats (en local)
            p_Drive_Résultats = SYS_GetIni("Chemins", "Drive_Resultats", p_CheminRapportType_Ini)
            p_Chemin_Résultats = SYS_GetIni("Chemins", "Path_Resultats", p_CheminRapportType_Ini)
            Me.MnuCheminRes.Caption = "Résultats => " & p_Drive_Résultats & p_Chemin_Résultats
            If p_Drive_Résultats & p_Chemin_Résultats = "" Then
                MsgBox "Chemin des Résultats non renseigné"
                End
            ElseIf Not FICH_EstRepertoire(p_Drive_Résultats & p_Chemin_Résultats, False) Then
                MsgBox "Chemin des Résultats " & p_Drive_Résultats & p_Chemin_Résultats & " introuvable"
            End If
        End If
        
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
        Me.MnuFichierIni.Caption = "Ouvrir le .Ini"
    End If

    p_Bool_Modif_Resolution = False
    
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)

    Set p_HTTP_Form_Menu = LanceRapportType

    Set p_HTTP_Form_Frame = LanceRapportType
    
    Call VerifVersion
    
End Sub

Private Sub TesterResolutionEcran()
    Dim Anc_ScreenResolution As String
    
    Anc_ScreenResolution = ScreenResolution(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    
    Me.Caption = Me.Caption & "  (" & Anc_ScreenResolution & ")"
    If pAnc_Largeur_Ecran < pNew_Largeur_Ecran Or pAnc_Hauteur_Ecran < pNew_Hauteur_Ecran Then
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
    Exc_obj.Quit
    Set Exc_obj = Nothing
    
    If p_Bool_Modif_Resolution Then
        Call ResolutionEcran(pAnc_Largeur_Ecran, pAnc_Hauteur_Ecran)
    End If

Err_Unload:
   Resume Next
End Sub

Private Sub MnuFichierIni_Click()
    FctOuvrirFichier (p_CheminRapportType_Ini)
End Sub

Private Sub MnuQuitter_Click()
   On Error GoTo Err_Unload
   Exc_obj.Quit
   Set Exc_obj = Nothing
   End
Err_Unload:
   Resume Next
End Sub


