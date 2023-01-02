Attribute VB_Name = "MCommon"
Option Explicit

Public Const P_NOIR = &H0&
Public Const P_GRIS = &HC0C0C0
Public Const P_GRIS_CLAIR = &HE0E0E0
Public Const P_GRIS_FONCE = &H808080
Public Const P_JAUNE = &H80FFFF
Public Const P_JAUNE_PASTEL = &HC0FFFF
Public Const P_BLANC = &HFFFFFF
Public Const P_ORANGE = &H80C0FF
Public Const P_ORANGE_FONCE = &H80FF&
Public Const P_CYAN = &HFF0000
Public Const P_BLEU = &H800000
Public Const P_ROSE = &H8080FF
Public Const P_ROUGE = &H80&
Public Const P_ROUGE_VIF = &HFF&
Public Const P_VERT_CLAIR = &HC0FFC0
Public Const P_VERT = &HFF00&
Public Const P_VERT_FONCE = &H8000&
Public Const P_MAGENTA = &H800080

Public Const P_ERREUR = -1
Public Const P_OK = 1
Public Const P_NON = 0
Public Const P_OUI = 1

Public Const FICH_FICHIER = 1
Public Const FICH_REP = 2
Public Const FICH_LECTURE = 1
Public Const FICH_ECRITURE = 2

'******************** HTML HELP **********************************************
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_TP_HELP_CONTEXTMENU = &H10
Public Const HH_TP_HELP_WM_HELP = &H11


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     ByVal dwData As Any) As Long
     
'********************* Fonctions système *****************************************
Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSiez As Long) As Long
Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSiez As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nsize As Long, _
    ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, _
    ByVal nsize As Long) As Long
Declare Function OpenIcon Lib "user32" _
    (ByVal hwnd As Long) As Long
Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAcess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" _
    (ByVal hwnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)
Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

'GetWindow constants
Public Const GW_HWNDPREV = 3
Public Const GW_CHILD = 5
Public Const GWL_STYLE = (-16)
Public Const WS_VSCROLL = &H200000

Const HWND_BROADCAST = &HFFFF
Const WM_WININICHANGE = &H1A

' Fenetre au 1er PLAN
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
' Premier niveau tout le temps
Private Const HWND_TOPMOST = -1
' Premier niveau mais derrière qd perd le focus
Private Const HWND_TOP = 0
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
                                                    ByVal cy As Long, ByVal wFlags As Long) As Long

' ***************** Pour gérer le dossier temporaire ********************************
' *********************************
' Pour modifier la résolution écran
' *********************************
Public p_Bool_Modif_Resolution As Boolean
Public Const pNew_Largeur_Ecran = 1024
Public Const pNew_Hauteur_Ecran = 768

Public pAnc_Largeur_Ecran As Single
Public pAnc_Hauteur_Ecran As Single
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DM_WIDTH = &H80000
Private Const DM_HEIGHT = &H100000
Private Const WM_DEVMODECHANGE = &H1B
Private Const HWND_DESKTOP = 0

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Dim dmEcran As DEVMODE
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Const CSIDL_ALTSTARTUP = &H1D ' * CSIDL_ALTSTARTUP - File system directory that corresponds to the user's nonlocalized Startup program group. (All Users\Startup?)
Private Const CSIDL_APPDATA = &H1A ' * CSIDL_APPDATA - File system directory that serves as a common repository for application-specific data. A common path is C:\WINNT\Profiles\username\Application Data.
Private Const CSIDL_BITBUCKET = &HA ' * CSIDL_BITBUCKET - Virtual folder containing the objects in the user's Recycle Bin.
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E ' * CSIDL_COMMON_ALTSTARTUP - File system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_APPDATA = &H23 ' * CSIDL_COMMON_APPDATA - Version 5.0. Application data for all users. A common path is C:\WINNT\Profiles\All Users\Application Data.
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 ' * CSIDL_DESKTOPDIRECTORY - File system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A common path is C:\WINNT\Profiles\username\Desktop
Private Const CSIDL_COMMON_DOCUMENTS = &H2E ' * CSIDL_COMMON_DOCUMENTS - File system directory that contains documents that are common to all users. A common path is C:\WINNT\Profiles\All Users\Documents. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_FAVORITES = &H1F ' * CSIDL_COMMON_FAVORITES - File system directory that serves as a common repository for all users' favorite items. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_PROGRAMS = &H17 ' * CSIDL_COMMON_PROGRAMS - File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A common path is c:\WINNT\Profiles\All Users\Start Menu\Programs. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_STARTMENU = &H16 ' * CSIDL_COMMON_STARTMENU - File system directory that contains the programs and folders that appear on the Start menu for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_STARTUP = &H18 ' * CSIDL_COMMON_STARTUP - File system directory that contains the programs that appear in the Startup folder for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu\Programs\Startup. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_TEMPLATES = &H2D ' * CSIDL_COMMON_TEMPLATES - File system directory that contains the templates that are available to all users. A common path is C:\WINNT\Profiles\All Users\Templates. Valid only for Windows NT systems.
Private Const CSIDL_COOKIES = &H21 ' * CSIDL_COOKIES - File system directory that serves as a common repository for Internet cookies. A common path is C:\WINNT\Profiles\username\Cookies.
Private Const CSIDL_DESKTOPDIRECTORY = &H10 ' * CSIDL_COMMON_DESKTOPDIRECTORY - File system directory that contains files and folders that appear on the desktop for all users. A common path is C:\WINNT\Profiles\All Users\Desktop. Valid only for Windows NT systems.
Private Const CSIDL_FAVORITES = &H6 ' * CSIDL_FAVORITES - File system directory that serves as a common repository for the user's favorite items. A common path is C:\WINNT\Profiles\username\Favorites.
Private Const CSIDL_FONTS = &H14 ' * CSIDL_FONTS - Virtual folder containing fonts. A common path is C:\WINNT\Fonts.
Private Const CSIDL_HISTORY = &H22 ' * CSIDL_HISTORY - File system directory that serves as a common repository for Internet history items.
Private Const CSIDL_INTERNET_CACHE = &H20 ' * CSIDL_INTERNET_CACHE - File system directory that serves as a common repository for temporary Internet files. A common path is C:\WINNT\Profiles\username\Temporary Internet Files.
Private Const CSIDL_LOCAL_APPDATA = &H1C ' * CSIDL_LOCAL_APPDATA - Version 5.0. File system directory that serves as a data repository for local (non-roaming) applications. A common path is C:\WINNT\Profiles\username\Local Settings\Application Data.
Private Const CSIDL_PROGRAMS = &H2 ' * CSIDL_PROGRAMS - File system directory that contains the user's program groups (which are also file system directories). A common path is C:\WINNT\Profiles\username\Start Menu\Programs.
Private Const CSIDL_PROGRAM_FILES = &H26 ' * CSIDL_PROGRAM_FILES - Version 5.0. Program Files folder. A common path is C:\Program Files.
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B ' * CSIDL_PROGRAM_FILES_COMMON - Version 5.0. A folder for components that are shared across applications. A common path is C:\Program Files\Common. Valid only for Windows NT and Windows® 2000 systems.
Private Const CSIDL_PERSONAL = &H5 ' * CSIDL_PERSONAL - File system directory that serves as a common repository for documents. A common path is C:\WINNT\Profiles\username\My Documents.
Private Const CSIDL_RECENT = &H8 ' * CSIDL_RECENT - File system directory that contains the user's most recently used documents. A common path is C:\WINNT\Profiles\username\Recent. To create a shortcut in this folder, use SHAddToRecentDocs. In addition to creating the shortcut, this function updates the shell's list of recent documents and adds the shortcut to the Documents submenu of the Start menu.
Private Const CSIDL_SENDTO = &H9 ' * CSIDL_SENDTO - File system directory that contains Send To menu items. A common path is c:\WINNT\Profiles\username\SendTo.
Private Const CSIDL_STARTUP = &H7 ' * CSIDL_STARTUP - File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows® 95. A common path is C:\WINNT\Profiles\username\Start Menu\Programs\Startup.
Private Const CSIDL_STARTMENU = &HB ' * CSIDL_STARTMENU - File system directory containing Start menu items. A common path is c:\WINNT\Profiles\username\Start Menu.
Private Const CSIDL_SYSTEM = &H25 ' * CSIDL_SYSTEM - Version 5.0. System folder. A common path is C:\WINNT\SYSTEM32.
Private Const CSIDL_TEMPLATES = &H15 ' * CSIDL_TEMPLATES - File system directory that serves as a common repository for document templates.
Private Const CSIDL_WINDOWS = &H24 ' * CSIDL_WINDOWS - Version 5.0. Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables. A common path is C:\WINNT.

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                        (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                         pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                        (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Function Rep_Documents(ByVal sCle As String, ByVal AncCle As String, ByRef NewCle As String) As String
    Dim lret As Long, IDL As ITEMIDLIST, sPath As String
    Dim Msg As String, s As String, sc As String

Debut:
    If AncCle = "TEST" Then
        GoTo Lab_Test
    ElseIf AncCle = "CSIDL_PERSONAL" Then
        lret = SHGetSpecialFolderLocation(100&, CSIDL_PERSONAL, IDL)
    ElseIf AncCle = "CSIDL_LOCAL_APPDATA" Then
        lret = SHGetSpecialFolderLocation(100&, CSIDL_LOCAL_APPDATA, IDL)
    ElseIf AncCle = "CSIDL_COMMON_DOCUMENTS" Then
        lret = SHGetSpecialFolderLocation(100&, CSIDL_COMMON_DOCUMENTS, IDL)
    ElseIf AncCle = "USERPROFILE" Then
        Rep_Documents = Environ$("USERPROFILE")
        Exit Function
    ElseIf AncCle = "TEMP" Then
        Rep_Documents = Environ$("TEMP")
        Exit Function
    ElseIf AncCle = "APPDATA" Then
        Rep_Documents = Environ$("APPDATA")
        Exit Function
    ElseIf AncCle = "KaliDoc" Or AncCle = "" Then
        Rep_Documents = p_chemin_appli
        Exit Function
    Else
        If AncCle <> "" Then MsgBox "Mauvaise syntaxe pour " & sCle
        GoTo Lab_Test
    End If
    If lret = 0 Then
        sPath = String$(512, Chr$(0))
        lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        Rep_Documents = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Else
        Rep_Documents = vbNullString
    End If
    Exit Function
Lab_Test:
    
    Call CL_Init
    
    sc = "CSIDL_PERSONAL"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_PERSONAL, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "CSIDL_LOCAL_APPDATA"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_LOCAL_APPDATA, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "CSIDL_COMMON_DOCUMENTS"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_COMMON_DOCUMENTS, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "USERPROFILE"
    s = Environ$(sc)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "TEMP"
    s = Environ$(sc)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "APPDATA"
    s = Environ$(sc)
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    sc = "KaliDoc"
    s = p_chemin_appli
    Call CL_AddLigne(sc & vbTab & s, 0, sc, True)
    
    Call CL_InitTitreHelp("Chemin possibles pour " & sCle, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    
    If CL_liste.retour = 1 Then
        Rep_Documents = ""
        NewCle = ""
        End
    End If
        
    AncCle = CL_liste.lignes(CL_liste.pointeur).tag
    NewCle = AncCle
    GoTo Debut
End Function


'******* FONCTIONS FICHIER **********************************

' Copie le fichier v_nomfich_src dans v_nomfich_dest
Public Function FICH_CopierFichier(ByVal v_nomfich_src As String, _
                                   ByVal v_nomfich_dest As String) As Integer

    If Not FICH_FichierExiste(v_nomfich_src) Then
        MsgBox "Le fichier " & v_nomfich_src & " n'existe pas.", vbCritical + vbOKOnly, "Fichier (FICH_CopierFichier)"
        FICH_CopierFichier = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_copy
    Call FileCopy(v_nomfich_src, v_nomfich_dest)
    On Error GoTo 0
    
    FICH_CopierFichier = P_OK
    Exit Function
    
err_copy:
    MsgBox "Impossible de copier " & v_nomfich_src & " dans " & v_nomfich_dest & "." & vbcr & vbLf & "Err=" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Fichier (FICH_CopierFichier)"
    On Error GoTo 0
    FICH_CopierFichier = P_ERREUR
    Exit Function
    
End Function

' Copie le fichier v_nomfich_src dans v_nomfich_dest
Public Function FICH_CopierFichierNoMess(ByVal v_nomfich_src As String, _
                                         ByVal v_nomfich_dest As String) As Integer

    If Not FICH_FichierExiste(v_nomfich_src) Then
        MsgBox "Le fichier " & v_nomfich_src & " n'existe pas.", vbCritical + vbOKOnly, "Fichier (FICH_CopierFichier)"
        FICH_CopierFichierNoMess = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_copy
    Call FileCopy(v_nomfich_src, v_nomfich_dest)
    On Error GoTo 0
    
    FICH_CopierFichierNoMess = P_OUI
    Exit Function
    
err_copy:
    On Error GoTo 0
    FICH_CopierFichierNoMess = P_NON
    Exit Function
    
End Function

' Création du répertoire v_nomrep en créant toute l'arborescence
' si nécessaire avec
'   confirmation quand le répertoire existe déjà
'   confirmation à chaque création d'un répertoire
Public Function FICH_CreerRepComp(ByVal v_nomrep As String, _
                                  ByVal v_bconfirm_si_existe As Boolean, _
                                  ByVal v_bconfirm As Boolean) As Integer

    Dim n As Integer, I As Integer, reponse As Integer, nmin As Integer, j As Integer
    Dim sdir As String, s As String
    
    On Error GoTo lab_existe_pas
    s = Dir$(v_nomrep, vbDirectory)
    On Error GoTo 0
    
    ' Le répertoire existe déjà
    If s <> "" Then
        If v_bconfirm_si_existe Then
            reponse = MsgBox("Le répertoire '" & v_nomrep & "' existe déjà." & vbcr & vbLf & "Confirmez-vous votre choix ?", vbQuestion + vbYesNo, "")
            If reponse = vbNo Then
                FICH_CreerRepComp = P_NON
            Else
                FICH_CreerRepComp = P_OUI
            End If
        Else
            FICH_CreerRepComp = P_OUI
        End If
        Exit Function
    End If
    
    n = STR_GetNbchamp(v_nomrep, "\")
    If n = 0 Then
        FICH_CreerRepComp = P_ERREUR
        Exit Function
    End If
    If left$(v_nomrep, 2) = "\\" Then
        nmin = 2
    Else
        nmin = 0
    End If
    ' Recherche le dernier répertoire de la chaine existant
    For I = n - 2 To nmin Step -1
        sdir = ""
        For j = 0 To I
            sdir = sdir + STR_GetChamp(v_nomrep, "\", j) + "\"
        Next j
'        sdir = left$(sdir, Len(sdir) - 1)
        On Error GoTo lab_err_drive
        s = Dir$(sdir, vbDirectory)
        On Error GoTo 0
        If s <> "" Then
            Exit For
        End If
    Next I
    ' Création des répertoires
    For j = I + 1 To n - 2
        sdir = sdir + STR_GetChamp(v_nomrep, "\", j)
        On Error GoTo lab_err_drive
        s = Dir$(sdir, vbDirectory)
        On Error GoTo 0
        If s = "" Then
            If v_bconfirm Then
                reponse = MsgBox("Confirmez-vous la création du répertoire '" & sdir & "'", vbQuestion + vbYesNo, "")
                If reponse = vbNo Then
                    FICH_CreerRepComp = P_NON
                    Exit Function
                End If
            End If
            On Error GoTo err_mkdir
            MkDir sdir
            On Error GoTo 0
        End If
    Next j
    
    If v_bconfirm Then
        reponse = MsgBox("Confirmez-vous la création du répertoire '" & v_nomrep & "'", vbQuestion + vbYesNo, "")
        If reponse = vbNo Then
            FICH_CreerRepComp = P_NON
            Exit Function
        End If
    End If
    On Error GoTo err_mkdir
    MkDir v_nomrep
    On Error GoTo 0
    
    FICH_CreerRepComp = P_OUI
    Exit Function

err_mkdir:
    MsgBox "Erreur création du répertoire '" + sdir & "'" & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "Fichier (FICH_CreerRepComp)"
    FICH_CreerRepComp = P_ERREUR
    Exit Function

lab_err_drive:
    MsgBox "Impossible de tester l'existance de '" & sdir & "'.", vbCritical + vbOKOnly, "FICH_CreerRepComp"
    FICH_CreerRepComp = P_NON
    Exit Function

lab_existe_pas:
    MsgBox "Impossible de tester l'existance de '" & v_nomrep & "'.", vbCritical + vbOKOnly, "FICH_CreerRepComp"
    FICH_CreerRepComp = P_NON
    Exit Function

End Function

' Efface le fichier v_nomfich
Public Function FICH_EffacerFichier(ByVal v_nomfich As String, _
                                    ByVal v_ya_mess As Boolean) As Integer

    On Error GoTo err_kill
    Call Kill(v_nomfich)
    On Error GoTo 0
    
    FICH_EffacerFichier = P_OK
    Exit Function
    
err_kill:
    If v_ya_mess Then
        MsgBox "Impossible d'effacer " & v_nomfich & "." & vbcr & vbLf & "Err=" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Fichier (FICH_EffacerFichier)"
    End If
    On Error GoTo 0
    FICH_EffacerFichier = P_ERREUR
    Exit Function
    
End Function

' Efface le répertoire v_nomrep et toute son arborescence
Public Sub FICH_EffacerRep(ByVal v_nomrep As String)

    Dim s As String, tbl_fich() As String
    Dim n As Integer, I As Integer
    
    On Error GoTo lab_existe_pas
    s = Dir$(v_nomrep + "\*.*", vbDirectory)
    On Error GoTo 0
    
    n = -1
    Do While s <> ""
        If s <> "." And s <> ".." Then
            n = n + 1
            ReDim Preserve tbl_fich(n) As String
            tbl_fich(n) = s
        End If
        s = Dir$
    Loop
    
    For I = 0 To n
        If FICH_EstRepertoire(v_nomrep + "\" + tbl_fich(I), False) Then
            On Error GoTo err_kill
            Call Kill(v_nomrep + "\" + tbl_fich(I) + "\*.*")
            On Error GoTo 0
            Call FICH_EffacerRep(v_nomrep + "\" + tbl_fich(I))
        End If
    Next I
    
    On Error GoTo err_kill
    Call Kill(v_nomrep + "\*.*")
    On Error GoTo 0
    
    On Error GoTo err_kill
    Call RmDir(v_nomrep)
    On Error GoTo 0
    
    Exit Sub
    
err_kill:
    Resume Next
    
lab_existe_pas:
    MsgBox "Impossible d'accéder à '" & v_nomrep & "'.", vbCritical + vbOKOnly, "FICH_EffacerRep"
    Exit Sub
    
End Sub

' Retourne True si v_nomfich est un répertoire
'          False sinon
Public Function FICH_EstRepertoire(ByVal v_nomfich As String, _
                                   ByVal v_affmess As Boolean) As Boolean
            
    If v_nomfich = "" Then
        FICH_EstRepertoire = False
        Exit Function
    End If
    
    On Error GoTo lab_existe_pas
    If (GetAttr(v_nomfich) And vbDirectory) = vbDirectory Then
        FICH_EstRepertoire = True
    Else
        FICH_EstRepertoire = False
    End If
    On Error GoTo 0
    Exit Function
    
lab_existe_pas:
    If v_affmess Then
        MsgBox "Impossible d'accéder à '" & v_nomfich & "'.", vbCritical + vbOKOnly, "FICH_EstRepertoire"
    End If
    FICH_EstRepertoire = False

End Function

' Retourne FICH_FICHIER si v_nomfich est un fichier
'          FICH_REP     si                  répertoire
'          P_NON        si           n'existe pas
Public Function FICH_EstFichierOuRep(ByVal v_nomfich As String) As Integer
            
    If v_nomfich = "" Then
        FICH_EstFichierOuRep = P_NON
        Exit Function
    End If
    
    On Error GoTo lab_existe_pas
    If (GetAttr(v_nomfich) And vbDirectory) = vbDirectory Then
        FICH_EstFichierOuRep = FICH_REP
    ElseIf Dir(v_nomfich) <> "" Then
        FICH_EstFichierOuRep = FICH_FICHIER
    Else
        FICH_EstFichierOuRep = P_NON
    End If
    On Error GoTo 0
    Exit Function
    
lab_existe_pas:
'    MsgBox "Impossible d'accéder à '" & v_nomfich & "'.", vbCritical + vbOKOnly, "FICH_EstFichierOuRep"
    FICH_EstFichierOuRep = P_NON

End Function

Public Function FICH_FichierDateTime(ByVal v_nomfich As String) As String

    On Error GoTo lab_err
    FICH_FichierDateTime = FileDateTime(v_nomfich)
    On Error GoTo 0
    Exit Function
    
lab_err:
    FICH_FichierDateTime = ""
    Exit Function

End Function

Public Function FICH_FichierExiste(ByVal v_nomfich As String) As Boolean
            
    If v_nomfich = "" Then
        FICH_FichierExiste = False
        Exit Function
    End If
    
    On Error GoTo lab_existe_pas
    If Dir$(v_nomfich) <> "" Then
        FICH_FichierExiste = True
    Else
        FICH_FichierExiste = False
    End If
    On Error GoTo 0
    Exit Function
    
lab_existe_pas:
    FICH_FichierExiste = False

End Function

Public Function FICH_Locker(ByVal v_fd As Integer, _
                            ByVal v_libfich As String) As Boolean
    
    Dim reponse As Integer
    
lab_lock:
    On Error GoTo err_lock
    Lock #v_fd
    On Error GoTo 0
    
    FICH_Locker = True
    Exit Function
    
err_lock:
    reponse = MsgBox("Le fichier " & v_libfich & " est verrouillé." & vbCrLf & vbCrLf & "Voulez-vous réessayer ?", vbYesNo + vbQuestion, "")
    If reponse = vbYes Then
        Call SYS_Sleep(1)
        GoTo lab_lock
    End If
    FICH_Locker = False
 
End Function

' Tente d'ouvrir le fichier v_nomfich et affecte le file
' descripteur correspondant r_fd
Public Function FICH_OuvrirFichier(ByVal v_nomfich As String, _
                                   ByVal v_mode As Integer, _
                                   ByRef r_fd As Integer) As Integer

    On Error GoTo err_open
    r_fd = FreeFile
    If v_mode = FICH_LECTURE Then
        Open v_nomfich For Input As r_fd
    Else
        Open v_nomfich For Append As r_fd
    End If
    On Error GoTo 0
    
    FICH_OuvrirFichier = P_OK
    Exit Function
    
err_open:
    MsgBox "Erreur Open " & v_nomfich & vbcr & vbLf & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "Fichier (FICH_OuvrirFichier)"
    FICH_OuvrirFichier = P_ERREUR
    Exit Function
    
End Function

' Renomme le fichier v_nomfich_src en v_nomfich_dest
Public Function FICH_RenommerFichier(ByVal v_nomfich_src As String, _
                                     ByVal v_nomfich_dest As String) As Integer

    On Error GoTo err_rename
    Name v_nomfich_src As v_nomfich_dest
    On Error GoTo 0
    
    FICH_RenommerFichier = P_OK
    Exit Function
    
err_rename:
    MsgBox "Impossible de renommer '" & v_nomfich_src & "' en '" & v_nomfich_dest & "'." & vbcr & vbLf & "Err=" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Fichier (FICH_RenommerFichier)"
    On Error GoTo 0
    FICH_RenommerFichier = P_ERREUR
    Exit Function
    
End Function

' Renomme le fichier v_nomfich_src en v_nomfich_dest
Public Function FICH_RenommerFichierNoMess(ByVal v_nomfich_src As String, _
                                           ByVal v_nomfich_dest As String) As Integer

    On Error GoTo err_rename
    Name v_nomfich_src As v_nomfich_dest
    On Error GoTo 0
    
    FICH_RenommerFichierNoMess = P_OK
    Exit Function
    
err_rename:
    On Error GoTo 0
    FICH_RenommerFichierNoMess = P_ERREUR
    Exit Function
    
End Function

' Tranfère le contenu du répertoire v_srcrep dans le répertoire
' v_destrep
Public Function FICH_TrsFichiersRep(ByVal v_srcrep As String, _
                                    ByVal v_destrep As String) As Integer
                                  
    Dim s As String, tbl_fich() As String
    Dim n As Integer, I As Integer
    
    If FICH_CreerRepComp(v_destrep, False, False) = P_ERREUR Then
        FICH_TrsFichiersRep = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo lab_existe_pas
    s = Dir$(v_srcrep + "\*.*", vbDirectory)
    On Error GoTo 0
    
    n = -1
    Do While s <> ""
        If s <> "." And s <> ".." Then
            n = n + 1
            ReDim Preserve tbl_fich(n) As String
            tbl_fich(n) = s
        End If
        s = Dir$
    Loop
    
    For I = 0 To n
        If FICH_EstRepertoire(v_srcrep + "\" + tbl_fich(I), False) Then
            If FICH_TrsFichiersRep(v_srcrep + "\" + tbl_fich(I), v_destrep + "\" + tbl_fich(I)) = P_ERREUR Then
                FICH_TrsFichiersRep = P_ERREUR
                Exit Function
            End If
        Else
            If FICH_CopierFichier(v_srcrep + "\" + tbl_fich(I), v_destrep + "\" + tbl_fich(I)) = P_ERREUR Then
                FICH_TrsFichiersRep = P_ERREUR
                Exit Function
            End If
        End If
    Next I
    
    FICH_TrsFichiersRep = P_OK
    Exit Function
    
lab_existe_pas:
    MsgBox "Impossible d'accéder à '" & v_srcrep & "'.", vbCritical + vbOKOnly, "FICH_TrsFichiersRep"
    FICH_TrsFichiersRep = P_ERREUR
    Exit Function
    
err_mkdir:
    MsgBox "Impossible de créer le répertoire '" & v_destrep & "'.", vbCritical + vbOKOnly, "FICH_TrsFichiersRep"
    FICH_TrsFichiersRep = P_ERREUR

End Function

'******* FONCTIONS EN RAPPORT AVEC LES FORMES ***************

Public Function FRM_AuPremierPlan(hwnd As Long) As Long

    FRM_AuPremierPlan = SetWindowPos(hwnd, HWND_TOP, 100, 0, 0, 0, FLAGS)
    
End Function

Public Sub FRM_CentrerForm(ByRef v_frm As Form)
    
    v_frm.Move (Screen.width - v_frm.width) / 2, (Screen.Height - v_frm.Height) / 2

End Sub

Public Sub FRM_FormEnbas(ByRef v_frm As Form)
    
    v_frm.Move (Screen.width - v_frm.width) / 2, Screen.Height - (v_frm.Height + 255)

End Sub

Public Sub FRM_FormEnhaut(ByRef v_frm As Form)
    
    v_frm.Move (Screen.width - v_frm.width) / 2, 255

End Sub

Public Function FRM_LargeurTexte(ByRef r_frm As Form, _
                                 ByVal v_obj As Object, _
                                 ByVal v_str As String) As Long

    r_frm.FontName = v_obj.FontName
    r_frm.FontSize = v_obj.FontSize
    r_frm.FontBold = v_obj.FontBold
    r_frm.FontItalic = v_obj.FontItalic
    r_frm.FontStrikethru = v_obj.FontStrikethru
    r_frm.FontUnderline = v_obj.FontUnderline
    FRM_LargeurTexte = r_frm.TextWidth(v_str) + 10
    
End Function

Public Sub FRM_ResizeForm(ByRef r_frm As Form, _
                          ByVal v_largeur As Integer, _
                          ByVal v_hauteur As Integer)

    If v_largeur > 0 Then
        r_frm.width = v_largeur
        r_frm.Height = v_hauteur
        Call FRM_CentrerForm(r_frm)
    Else
        r_frm.left = Screen.width
        r_frm.Top = Screen.Height
    End If
    
End Sub

Public Function FRM_EstEnCours(ByRef v_frm As Variant) As Boolean

    If GetActiveWindow() <> v_frm.hwnd Then
        FRM_EstEnCours = False
    Else
        FRM_EstEnCours = True
    End If

End Function

'******* FONCTIONS SYSTEME **********************************

' Lance le programme contenu dans v_scmd en attendant la fin
' d'exécution si v_attendre=True
Public Sub SYS_ExecShell(ByVal v_scmd As String, _
                         ByVal v_attendre As Boolean, _
                         ByVal v_visible As Boolean, _
                         Optional v_mess_err As String)

    Dim pid As Long, hproc As Long, iret As Long

    On Error GoTo err_shell
    If v_visible Then
        pid = Shell(v_scmd, 1)
    Else
        pid = Shell(v_scmd, vbHide)
    End If
    On Error GoTo 0
    
    If v_attendre Then
        hproc = OpenProcess(&H1F0FFF, False, pid)
        If hproc <> 0 Then
            Do
                iret = WaitForSingleObject(hproc, 500)
                DoEvents
            Loop Until iret = 0
            CloseHandle hproc
        Else
            MsgBox "Erreur OpenProcess renvoie 0 pour le pid " & pid, vbCritical + vbOKOnly, "Common (Cm_ExecShell)"
            GoTo lab_fin
        End If
    End If
    
lab_fin:
    Exit Sub
    
err_shell:
    Call MsgBox("Impossible d'exécuter : " & v_scmd & vbcr & vbLf & Err.Number & " " & Err.Description & " " & v_mess_err, vbExclamation + vbOKOnly, "")
    Exit Sub
    
End Sub

' Renvoie le nom du poste Windows
Public Function SYS_GetComputerName() As String

    Dim s As String
    Dim l As Long
    
    s = Space(512)
    l = Len(s)
    If CBool(GetComputerName(s, l)) Then
        SYS_GetComputerName = left$(s, l)
    Else
        SYS_GetComputerName = ""
    End If
    
End Function

' Renvoie le nom du user Windows
Public Function SYS_GetUserName() As String

    Dim s As String
    Dim l As Long
    
    s = Space(512)
    l = Len(s)
    If CBool(GetUserName(s, l)) Then
        SYS_GetUserName = Mid$(s, 1, InStr(1, s, Chr$(0)) - 1)
    Else
        SYS_GetUserName = ""
    End If
    
End Function

' Retourne la valeur du registre correspondant à la
' section v_section pour la clé v_cle
Public Function SYS_GetRegistre(ByVal v_section As String, _
                                ByVal v_cle As String) As String
    
    SYS_GetRegistre = GetSetting(App.ProductName, _
                                   v_section, _
                                   v_cle, _
                                   "")

End Function

' Ecrit dans le registre à la section v_section la clé v_cle
' de valeur v_sval
Public Sub SYS_PutRegistre(ByVal v_section As String, _
                           ByVal v_cle As String, _
                           ByVal v_sval As String)
    
    Call SaveSetting(App.ProductName, _
                      v_section, _
                      v_cle, _
                      v_sval)

End Sub

Public Function SYS_GetImpSystem() As String

    Dim chemin As String, sret As String
    Dim nc As Long
    
    chemin = String$(260, 0)
    chemin = left$(chemin, GetWindowsDirectory(chemin, Len(chemin))) + "\win.ini"
    sret = String$(255, 0)
    nc = GetPrivateProfileString("windows", "device", "", sret, 255, chemin)
    sret = left$(sret, nc)
    nc = InStr(sret, ",")
    SYS_GetImpSystem = left(sret, nc - 1)
    
End Function

' Retourne la valeur dans le fichier v_nomfich
' correspondant à la section v_section pour la clé v_cle
Public Function SYS_GetIni(ByVal v_section As String, _
                           ByVal v_cle As String, _
                           ByVal v_nomfich As String) As Variant

    Dim pos As Integer
    Dim retour As Long
    Dim buf As String * 512
    
    retour = GetPrivateProfileString(v_section, _
                                     v_cle, _
                                     "", _
                                     buf, _
                                     512, _
                                     v_nomfich)
    pos = InStr(buf, Chr(0))
    If pos > 0 Then
        SYS_GetIni = left$(buf, pos - 1)
    Else
        SYS_GetIni = ""
    End If
    
End Function

' Ecrit dans le fichier v_nomfich à la section v_section
' la clé v_cle de valeur v_sval
Public Sub SYS_PutIni(ByVal v_section As String, _
                      ByVal v_cle As String, _
                      ByVal v_sval As String, _
                      ByVal v_nomfich As String)

    Dim lret As Long
    
    lret = WritePrivateProfileString(v_section, _
                                     v_cle, _
                                     v_sval, _
                                     v_nomfich)

End Sub

Public Sub SYS_SetImpSystem(ByVal v_devicename As String)

    Dim chemin As String, sret As String
    Dim nc As Long
    
    chemin = String$(260, 0)
    chemin = left$(chemin, GetWindowsDirectory(chemin, Len(chemin))) + "\win.ini"
    sret = String$(255, 0)
    nc = GetPrivateProfileString("Devices", v_devicename, "", sret, 255, chemin)
    sret = left$(sret, nc)
    WritePrivateProfileString "windows", "device", v_devicename & "," & sret, chemin
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0, "windows"

End Sub

' Attente de v_delai milli-secondes
Public Sub SYS_Sleep(ByVal v_delai As Integer)

    Sleep (v_delai)
    
End Sub

Public Sub SYS_StartProcess(ByVal sFile As String, _
                            Optional ByVal sParameters As String = vbNullString)

    ShellExecute 0&, "open", sFile, sParameters, vbNullString, 1&
    
End Sub

Public Function SYS_Ya_PrevInstance() As Boolean

    Dim titre As String
    Dim hndw As Long
    
    If App.PrevInstance Then
        On Error GoTo lab_fin
        titre = App.Title
        App.Title = titre & "2"
        hndw = FindWindow("ThunderRT6Main", titre)
        hndw = GetWindow(hndw, GW_HWNDPREV)
        Call OpenIcon(hndw)
        Call SetForegroundWindow(hndw)
        SYS_Ya_PrevInstance = True
    Else
        SYS_Ya_PrevInstance = False
    End If
    Exit Function
    
lab_fin:
    SYS_Ya_PrevInstance = False

End Function

'**************** FONCTIONS EN RAPPORT AVEC LES DATES ***********

' retourne la premiere date à partir de 'v_date' qui tombe un 'v_jour'
Public Function DATE_PremierJour(ByVal v_date As Date, _
                                 ByVal v_jour As Integer) As Date

    ' Lundi = jour 0
    Do While Weekday(v_date, vbMonday) <> v_jour + 1
        v_date = v_date + 1
    Loop
    
    DATE_PremierJour = v_date
    
End Function

Public Function DATE_ToRFC822(ByVal v_date As Date)
    
    Dim tblDate(5)
    Dim tblWeekDayName
    Dim tblMonthName
    
    If Not IsDate(v_date) Then
        DATE_ToRFC822 = ""
        Exit Function
    End If
    
    tblWeekDayName = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    tblMonthName = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    tblDate(0) = tblWeekDayName(Weekday(v_date, 2) - 1) & ","
    If (Len(Day(v_date)) = 1) Then
        tblDate(1) = "0" & Day(v_date)
    Else
        tblDate(1) = Day(v_date)
    End If
    tblDate(2) = tblMonthName(Month(v_date) - 1)
    tblDate(3) = Year(v_date)
    tblDate(4) = TimeValue(v_date)
    tblDate(5) = "+0200"
    
    DATE_ToRFC822 = Join(tblDate, Space(1))
    
End Function

Public Function DATE_ToStrCalendrier(ByVal v_sdate As String) As String

    Dim s As String
    Dim jj As Integer, mm As Integer
    
    Do
        s = STR_DateTosDate(v_sdate)
        If s = "" Then
            mm = Mid$(v_sdate, 4, 2)
            If mm < 1 Or mm > 12 Then
                v_sdate = left$(v_sdate, 2) + "/12/" + Mid$(v_sdate, 7)
            Else
                jj = left$(v_sdate, 2)
                jj = jj - 1
                v_sdate = Format(jj, "00") + "/" + Mid$(v_sdate, 4)
            End If
        End If
    Loop Until s <> ""
    
    DATE_ToStrCalendrier = v_sdate
    
End Function

Public Function DATE_Incrementer(ByVal v_date As Date, _
                                 ByVal v_speriode As String) As Date

    Dim sdate As String, sper As String
    Dim nb As Integer, nbj As Integer, jj As Integer, mm As Integer, AA As Integer
    
    If v_speriode = "" Then
        DATE_Incrementer = v_date
        Exit Function
    End If
    
    nb = left$(v_speriode, Len(v_speriode) - 1)
    sper = Right$(v_speriode, 1)
    
    sdate = Format(v_date, "dd/mm/yyyy")
    jj = left$(sdate, 2)
    mm = Mid$(sdate, 4, 2)
    AA = Right$(sdate, 4)
    Select Case sper
    Case "J"
        DATE_Incrementer = v_date + nb
        Exit Function
    Case "S"
        nbj = nb * 7
        DATE_Incrementer = v_date + nbj
        Exit Function
    Case "M"
        mm = mm + nb
        While mm > 12
            AA = AA + 1
            mm = mm - 12
        Wend
    Case "A"
        AA = AA + nb
    End Select
    
    sdate = Format(jj, "00") + "/" + Format(mm, "00") + "/" + Format(AA, "0000")
    sdate = DATE_ToStrCalendrier(sdate)
    DATE_Incrementer = CDate(sdate)
    
End Function
                        
' Retourne le nombre de jours du mois v_mm de l'année v_aa
Public Function DATE_NbjoursMois(ByVal v_mm As Integer, _
                                 ByVal v_aa As Integer) As Integer

    Select Case v_mm
    Case 1, 3, 5, 7, 8, 10, 12
        DATE_NbjoursMois = 31
    Case 2
        If v_aa Mod 4 Then
            DATE_NbjoursMois = 28
        Else
            DATE_NbjoursMois = 29
        End If
    Case Else
        DATE_NbjoursMois = 30
    End Select
    
End Function

' Retourne le prix HT correspondant au prix TTC v_prixttc
' avec une tva v_tva
Public Function PRIX_TTCtoHT(ByVal v_prixttc As Double, _
                             ByVal v_tva As Double) As Double
                            
    If v_tva > 0 Then
        PRIX_TTCtoHT = (v_prixttc / (100 + v_tva)) * 100
    Else
        PRIX_TTCtoHT = v_prixttc
    End If

End Function

' Retourne le prix v_prix après avoir déduit la remise v_remise
Public Function PRIX_PrixRemisé(ByVal v_prix As Double, _
                                ByVal v_remise As Double) As Double
                            
    If v_remise > 0 Then
        PRIX_PrixRemisé = v_prix - ((v_prix * v_remise) / 100)
    Else
        PRIX_PrixRemisé = v_prix
    End If

End Function

' Retourne le prix TTC correspondant au prix HT v_prixht
' avec une tva v_tva
Public Function PRIX_HTtoTTC(ByVal v_prixht As Double, _
                             ByVal v_tva As Double) As Double
                            
    If v_tva > 0 Then
        PRIX_HTtoTTC = v_prixht + ((v_prixht * v_tva) / 100)
    Else
        PRIX_HTtoTTC = v_prixht
    End If

End Function

Public Function CM_LoadPicture(ByVal nomfic As String) As Picture

    On Error GoTo err_img
    Set CM_LoadPicture = LoadPicture(nomfic)
    Exit Function

err_img:
    MsgBox "Impossible de charger l'image " & nomfic & "." & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Set CM_LoadPicture = LoadPicture("")
    
End Function

Public Function CM_UboundL(ByRef tbl() As Long) As Long

    On Error GoTo err_tbl_vide
    CM_UboundL = UBound(tbl)
    On Error GoTo 0
    Exit Function
    
err_tbl_vide:
    On Error GoTo 0
    CM_UboundL = -1
    Exit Function

End Function

'******* FONCTIONS CHIANE DE CARACTERES *********************

' Retourne l'entier v_ent (hhmm ex: 800 pour 8h) en chaine
' sous la forme hh:mm
Public Function STR_EntierToSheure(ByVal v_ent As Integer) As String

    If v_ent = -1 Then
        STR_EntierToSheure = ""
        Exit Function
    End If
    
    STR_EntierToSheure = Format(v_ent / 100, "#00") + ":" + Format(v_ent Mod 100, "#00")
    
End Function
                                
Public Function STR_StrToBool(ByVal v_str As String) As Boolean

    On Error GoTo lab_test2
    STR_StrToBool = CBool(v_str)
    On Error GoTo 0
    Exit Function
    
lab_test2:
    If UCase(v_str) = "FAUX" Then
        STR_StrToBool = False
    Else
        STR_StrToBool = True
    End If
        
End Function

' Retourne le v_nochp ieme champ dans la chaine v_str
' ayant v_sep comme caractère séparateur de champ
Public Function STR_GetChamp(ByVal v_str As String, _
                             ByVal v_sep As String, _
                             ByVal v_nochp As Integer) As String

    Dim I As Integer, pos As Integer
    Dim s As String
    
    s = v_str
    For I = 0 To v_nochp - 1
        pos = InStr(s, v_sep)
        If pos = 0 Then
            STR_GetChamp = ""
            Exit Function
        End If
        s = Mid$(s, pos + Len(v_sep))
    Next I
    
    pos = InStr(s, v_sep)
    If pos = 0 Then
        STR_GetChamp = s
    Else
        STR_GetChamp = left$(s, pos - 1)
    End If

End Function

' Retourne le nombre de champs de la chaine v_str
' ayant v_sep comme caractère séparateur de champ
Public Function STR_GetNbchamp(ByVal v_str As String, _
                               ByVal v_sep As String) As Integer

    Dim n As Integer, pos As Integer
    Dim s As String
    
    s = v_str
    If s = "" Then
        STR_GetNbchamp = 0
        Exit Function
    End If
    n = 0
    Do
        pos = InStr(s, v_sep)
        If pos = 0 Then
            If Len(s) > 0 Then n = n + 1
            STR_GetNbchamp = n
            Exit Function
        End If
        n = n + 1
        s = Mid$(s, pos + Len(v_sep))
    Loop
    
End Function

' Retourne la chaine v_sheure sous la forme hh:mm
' en entier (hhmm ex: 08:00 devient 800)
Public Function STR_SHeureToEntier(ByVal v_sheure As String) As Integer
                               
    If v_sheure = "" Then
        STR_SHeureToEntier = -1
        Exit Function
    End If
    
    STR_SHeureToEntier = (Mid$(v_sheure, 1, 1) * 1000) + _
                         (Mid$(v_sheure, 2, 1) * 100) + _
                         (Mid$(v_sheure, 4, 1) * 10) + _
                          Mid$(v_sheure, 5, 1)

End Function

Public Function STR_Incrementer(ByVal v_str As String) As String

    Dim sformat As String
    Dim n As Integer
    
    sformat = String$(Len(v_str), "0")
    n = CInt(v_str) + 1
    STR_Incrementer = Format(n, sformat)

End Function

Public Function STR_Crypter(ByVal v_str As String) As String
    
    Dim strTemp As String '*** contiendra la résultat du cryptage caractère par caractère.
    Dim I As Long '*** compteur de caractères du texte à protéger
    Dim q As Integer '*** stockage de la valeur ASCII
    Dim pt As String * 1 '*** texte normal à protéger
    Dim ct As String * 1 '*** texte chiffré
    Dim z As Integer '*** valeur du cycle de rotation
    Dim Step As Integer '*** pas d'incrémentation négatif ou positif
    Dim cycle As Integer
    
    cycle = 7
    
    z = cycle
    Step = -1
    For I = 1 To Len(v_str)
        pt = Mid(v_str, I, 1)
        q = Asc(pt)
        Select Case q
        Case Asc("A") To Asc("Z") '***Majuscules
            q = q + z
            If q > Asc("Z") Then
                q = 64 + (q - Asc("Z"))
            End If
            ct = Chr(q)
        Case Asc("a") To Asc("z") '***Minuscules
            q = q + z
            If q > Asc("z") Then
                q = 96 + (q - Asc("z"))
            End If
            ct = Chr(q)
        Case Else   '***Ponctuation et chiffres:
            ct = Chr(q + z)
        End Select
        z = z + Step
        If z < 0 Then '***Démarre l'incrémentation positive
            z = 1
            Step = 1
        End If
        If z > cycle Then '***Démarre l'incrémentation négative
            z = cycle - 1
            Step = -1
        End If
        strTemp = strTemp & ct
    Next I
    
    STR_Crypter = strTemp
    
End Function

Public Function STR_Crypter_New(ByVal v_str As String) As String
    
    Dim strTemp As String '*** contiendra la résultat du cryptage caractère par caractère.
    Dim I As Long '*** compteur de caractères du texte à protéger
    Dim q As Integer '*** stockage de la valeur ASCII
    Dim pt As String * 1 '*** texte normal à protéger
    Dim ct As String * 1 '*** texte chiffré
    Dim z As Integer '*** valeur du cycle de rotation
    Dim Step As Integer '*** pas d'incrémentation négatif ou positif
    Dim cycle As Integer
    
    cycle = 7
    
    z = cycle
    Step = -1
    For I = 1 To Len(v_str)
        pt = Mid(v_str, I, 1)
        q = Asc(pt)

        If q <= 128 Then
            q = q + z
            If q >= 128 Then
                q = q - 127
            End If
        Else
            q = q + z
        End If
        ct = Chr(q)
        
        z = z + Step
        If z < 0 Then '***Démarre l'incrémentation positive
            z = 1
            Step = 1
        End If
        If z > cycle Then '***Démarre l'incrémentation négative
            z = cycle - 1
            Step = -1
        End If
        strTemp = strTemp & ct
    Next I
    
    STR_Crypter_New = strTemp
    
End Function

Public Function STR_CrypterNombre(ByVal v_str As String) As String
                            
    Dim s_code As String, sc As String, str As String
    Dim I As Integer, pos As Integer
    
    s_code = "aqwZSXedcRFVtgbYHNujkIOPml"
    
    str = ""
    For I = 1 To Len(v_str)
        sc = Mid$(v_str, I, 1)
        pos = Int(sc) + I - 1
        str = str + Mid$(s_code, pos + 1, 1)
    Next I
    
    STR_CrypterNombre = str
    
End Function

Public Function STR_Decrypter(ByVal v_str As String) As String
    
    Dim strTemp As String
    Dim I As Long
    Dim ct As String * 1 '*** texte chiffré
    Dim pt As String * 1 '*** texte normal
    Dim q As Integer
    Dim z As Integer
    Dim Step As Integer
    Dim cycle As Integer
    
    cycle = 7
    
    z = cycle
    Step = -1
    For I = 1 To Len(v_str)
        ct = Mid(v_str, I, 1)
        q = Asc(ct)
        Select Case q
        Case Asc("A") To Asc("Z")
           '***Majuscules
           q = q - z
           If q < Asc("A") Then
              q = Asc("Z") - (64 - q)
           End If
           strTemp = strTemp & Chr(q)
        Case Asc("a") To Asc("z")
           '***Minuscules
           q = q - z
           If q < Asc("a") Then
              q = Asc("z") - (96 - q)
           End If
           strTemp = strTemp & Chr(q)
        Case Else '***Ponctuation et chiffres:
           strTemp = strTemp & Chr(q - z)
        End Select
        z = z + Step
        If z < 0 Then
            z = 1
            Step = 1
        End If
        If z > cycle Then
            z = cycle - 1
            Step = -1
        End If
    Next I
    
    STR_Decrypter = strTemp
    
End Function

Public Function STR_Decrypter_New(ByVal v_str As String) As String
    
    Dim strTemp As String
    Dim I As Long
    Dim ct As String * 1 '*** texte chiffré
    Dim pt As String * 1 '*** texte normal
    Dim q As Integer
    Dim z As Integer
    Dim Step As Integer
    Dim cycle As Integer
    
    cycle = 7
    
    z = cycle
    Step = -1
    For I = 1 To Len(v_str)
        ct = Mid(v_str, I, 1)
        q = Asc(ct)

        If q <= 128 Then
            q = q - z
            If q < 0 Then
                q = 127 - (0 - q)
            End If
        Else
            q = q - z
        End If
        strTemp = strTemp & Chr(q)

        z = z + Step
        If z < 0 Then
            z = 1
            Step = 1
        End If
        If z > cycle Then
            z = cycle - 1
            Step = -1
        End If
    Next I
    
    STR_Decrypter_New = strTemp
    
End Function

Public Function STR_DecrypterNombre(ByVal v_str As String) As String
                            
    Dim s_code As String, sc As String, str As String
    Dim I As Integer, pos As Integer
    
    s_code = "aqwZSXedcRFVtgbYHNujkIOPml"
    
    str = ""
    For I = 1 To Len(v_str)
        sc = Mid$(v_str, I, 1)
        pos = InStr(s_code, sc) - I
        str = str & pos
    Next I
    
    STR_DecrypterNombre = str
    
End Function

Public Function STR_ComparerPeriode(ByVal v_sper1 As String, _
                                    ByVal v_sper2 As String) As Integer
                    
    Dim su1 As String, su2 As String
    Dim nbj1 As Integer, nbj2 As Integer
    
    su1 = Right$(v_sper1, 1)
    su2 = Right$(v_sper2, 1)
    Select Case su1
    Case "J"
        nbj1 = CInt(left$(v_sper1, Len(v_sper1) - 1))
    Case "S"
        nbj1 = CInt(left$(v_sper1, Len(v_sper1) - 1)) * 7
    Case "M"
        nbj1 = CInt(left$(v_sper1, Len(v_sper1) - 1)) * 31
    Case "A"
        nbj1 = CInt(left$(v_sper1, Len(v_sper1) - 1)) * 365
    End Select
    Select Case su2
    Case "J"
        nbj2 = CInt(left$(v_sper2, Len(v_sper2) - 1))
    Case "S"
        nbj2 = CInt(left$(v_sper2, Len(v_sper2) - 1)) * 7
    Case "M"
        nbj2 = CInt(left$(v_sper2, Len(v_sper2) - 1)) * 31
    Case "A"
        nbj2 = CInt(left$(v_sper2, Len(v_sper2) - 1)) * 365
    End Select
    If nbj1 < nbj2 Then
        STR_ComparerPeriode = -1
    ElseIf nbj1 > nbj2 Then
        STR_ComparerPeriode = 1
    Else
        STR_ComparerPeriode = 0
    End If
    
End Function

Private Function STR_DateTosDate(ByVal v_sdate As String) As String

    Dim mdate As Date
    
    On Error GoTo err_date
    mdate = CDate(v_sdate)
    On Error GoTo 0
    STR_DateTosDate = v_sdate
    Exit Function

err_date:
    On Error GoTo 0
    STR_DateTosDate = ""
    Exit Function
    
End Function

' Initialise le tableau de chaine r_tabstr() avec les
' chaines se trouvant dans v_str séparées par des TAB
Public Function STR_Decouper(ByVal v_str As String, _
                             ByRef r_tabstr() As String)

    Dim I As Integer, pos As Integer
    Dim s As String
    
    s = v_str
    pos = InStr(s, vbTab)
    I = 0
    While pos > 0
        ReDim Preserve r_tabstr(I) As String
        r_tabstr(I) = left$(s, pos - 1)
        s = Mid$(s, pos + 1)
        pos = InStr(s, vbTab)
        I = I + 1
    Wend
    ReDim Preserve r_tabstr(I) As String
    r_tabstr(I) = s
    
End Function

Public Function STR_EstAlpha(ByVal v_str As String) As Boolean

    Dim s1 As String
    Dim I As Integer
    
    For I = 1 To Len(v_str)
        s1 = UCase(Mid$(v_str, I, 1))
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÉÈÇ", s1) = 0 Then
            STR_EstAlpha = False
            Exit Function
        End If
    Next I
    STR_EstAlpha = True
    
End Function

Public Function STR_EstEntier(ByVal v_str As String) As Boolean

    If Not IsNumeric(v_str) Then
        STR_EstEntier = False
        Exit Function
    End If
    If InStr(v_str, ".") > 0 Or InStr(v_str, ",") > 0 Then
        STR_EstEntier = False
        Exit Function
    End If
    
    STR_EstEntier = True
    
End Function

Public Function STR_EstEntierPos(ByVal v_str As String) As Boolean

    If Not STR_EstEntier(v_str) Then
        STR_EstEntierPos = False
        Exit Function
    End If
    If InStr(v_str, "-") > 0 Then
        STR_EstEntierPos = False
        Exit Function
    End If
    
    STR_EstEntierPos = True
    
End Function

Public Function STR_EstEntierNeg(ByVal v_str As String) As Boolean

    If Not STR_EstEntier(v_str) Then
        STR_EstEntierNeg = False
        Exit Function
    End If
    If InStr(v_str, "-") = 0 Then
        STR_EstEntierNeg = False
        Exit Function
    End If
    
    STR_EstEntierNeg = True
    
End Function

Public Function STR_EstDecimalPos(ByVal v_str As String) As Boolean

    If Not IsNumeric(v_str) Then
        STR_EstDecimalPos = False
        Exit Function
    End If
    If InStr(v_str, "-") > 0 Then
        STR_EstDecimalPos = False
        Exit Function
    End If
    
    STR_EstDecimalPos = True
    
End Function

Public Function STR_EstDecimalNeg(ByVal v_str As String) As Boolean

    If Not IsNumeric(v_str) Then
        STR_EstDecimalNeg = False
        Exit Function
    End If
    If InStr(v_str, "-") = 0 Then
        STR_EstDecimalNeg = False
        Exit Function
    End If
    
    STR_EstDecimalNeg = True
    
End Function

Public Function STR_EstPonctuation(ByVal v_str As String) As Boolean

    Dim s1 As String, stmp As String
    
    stmp = v_str
    Do While Len(stmp) > 0
        s1 = UCase(left$(stmp, 1))
        If InStr("'- ", s1) = 0 Then
            STR_EstPonctuation = False
            Exit Function
        End If
        stmp = Mid$(stmp, 2)
    Loop
    STR_EstPonctuation = True
    
End Function

Public Function STR_GarderChiffre(ByVal v_str As String) As String

    Dim I As Integer
    Dim stmp As String, str2 As String
    
    str2 = ""
    For I = 1 To Len(v_str)
        stmp = Mid$(v_str, I, 1)
        If InStr("0123456789", stmp) > 0 Then
            str2 = str2 & stmp
        End If
    Next I
    STR_GarderChiffre = str2
    
End Function

Public Function STR_Phonet(ByVal v_str As String) As String

    Dim str As String
    Dim reponse As Integer
    
    str = v_str
    str = STR_Remplacer(str, "à", "a")
    str = STR_Remplacer(str, "â", "a")
    str = STR_Remplacer(str, "ä", "a")
    str = STR_Remplacer(str, "é", "e")
    str = STR_Remplacer(str, "è", "e")
    str = STR_Remplacer(str, "ê", "e")
    str = STR_Remplacer(str, "ë", "e")
    str = STR_Remplacer(str, "î", "i")
    str = STR_Remplacer(str, "ï", "i")
    str = STR_Remplacer(str, "ô", "o")
    str = STR_Remplacer(str, "ö", "o")
    str = STR_Remplacer(str, "ù", "u")
    str = STR_Remplacer(str, "û", "u")
    str = STR_Remplacer(str, "ü", "u")
    str = STR_Remplacer(str, "ç", "c")
    str = UCase(str)
   
    STR_Phonet = str
    
End Function

Public Function STR_Prix(ByVal v_prix As Double) As String

    If v_prix > 0 Then
        STR_Prix = STR_SupprimerBlancDeb(Format(v_prix, "### ### ##0.00"))
    ElseIf v_prix < 0 Then
        STR_Prix = "-" & STR_SupprimerBlancDeb(Format(-v_prix, "### ### ##0.00"))
    Else
        STR_Prix = "0"
    End If
    
End Function

Public Sub STR_PutChamp(ByRef r_str As Variant, _
                        ByVal v_sep As String, _
                        ByVal v_nochp As Integer, _
                        ByVal v_val As Variant)

    Dim I As Integer, pos As Integer, vrai_pos As Integer
    Dim s As String, sD As String
    
    s = r_str
    vrai_pos = 0
    For I = 0 To v_nochp - 1
        pos = InStr(s, v_sep)
        If pos = 0 Then Exit Sub
        s = Mid$(s, pos + 1)
        vrai_pos = vrai_pos + pos
    Next I
    
    pos = InStr(s, v_sep)
    If pos = 0 Then
        sD = ""
    Else
        sD = Mid$(s, pos)
    End If
    
    s = left$(r_str, vrai_pos) & v_val & sD
    r_str = s
    
End Sub

Public Function STR_Remplacer(ByVal v_str As String, _
                              ByVal v_str_a_remplacer As String, _
                              ByVal v_new_str As String) As String

    Dim start As Integer, pos As Integer, len_str_a_remplacer As Integer
    Dim stmp As String, s_in As String, s_out As String
    
    s_in = v_str
    s_out = ""
    start = 1
    len_str_a_remplacer = Len(v_str_a_remplacer)
    pos = InStr(start, s_in, v_str_a_remplacer)
    Do While pos <> 0
        If start < pos Then
            stmp = Mid$(s_in, start, pos - start)
        Else
            stmp = ""
        End If
        stmp = stmp + v_new_str
        s_out = s_out + stmp
        start = pos + len_str_a_remplacer
        pos = InStr(start, s_in, v_str_a_remplacer)
    Loop
    STR_Remplacer = s_out + Mid$(s_in, start)

End Function

Public Function STR_RemplacerSeqDebFin(ByVal v_str As String, _
                                       ByVal v_strdeb As String, _
                                       ByVal v_strfin As String, _
                                       ByVal v_strempl As String) As String
                                       
    Dim s_out As String
    Dim pos As Integer, pos1 As Integer, pos2 As Integer
    
    s_out = ""
    pos = 1
    While pos <= Len(v_str)
        pos1 = InStr(pos, v_str, v_strdeb)
        If pos1 = 0 Then
            s_out = s_out + Mid$(v_str, pos)
            pos = Len(v_str) + 1
        Else
            pos2 = InStr(pos1 + 1, v_str, v_strfin)
            If pos2 = 0 Then
                s_out = s_out + Mid$(v_str, pos)
                pos = Len(v_str) + 1
            Else
                If pos <> pos1 Then
                    s_out = s_out + Mid$(v_str, pos, pos1 - pos)
                End If
                s_out = s_out + v_strempl
                pos = pos2 + 1
            End If
        End If
    Wend
    
    STR_RemplacerSeqDebFin = s_out
    
End Function

Public Function STR_SupprimerBlancDeb(ByVal v_str As String) As String

    Dim s As String
    
    s = v_str
    While left$(s, 1) = " "
        s = Mid$(s, 2)
    Wend
    STR_SupprimerBlancDeb = s

End Function

Public Function STR_SupprimerBlancFin(ByVal v_str As String) As String

    Dim s As String
    
    s = v_str
    While Right$(s, 1) = " "
        s = left$(s, Len(s) - 1)
    Wend
    STR_SupprimerBlancFin = s

End Function

Public Function STR_SupprimerChamp(ByVal v_str As String, _
                                   ByVal v_sep As String, _
                                   ByVal v_pos As Integer) As String
                                
    Dim s As String
    Dim n As Integer, I As Integer
    
    n = STR_GetNbchamp(v_str, v_sep)
    If n <= 1 Then
        STR_SupprimerChamp = ""
        Exit Function
    End If
    
    s = ""
    For I = 0 To v_pos - 1
        s = s + STR_GetChamp(v_str, v_sep, I) + v_sep
    Next I
    For I = v_pos + 1 To n - 1
        s = s + STR_GetChamp(v_str, v_sep, I) + v_sep
    Next I
    
    STR_SupprimerChamp = s
    
End Function

Public Function STR_LaisserUnSeulBlanc(ByVal v_str As Variant) As String

    Dim deja_bl As Boolean
    Dim s1 As Variant, s2 As Variant, sc As String
    Dim pos As Long
    
    s1 = Trim(v_str)
    s2 = ""
    
    Do
        pos = InStr(s1, "  ")
        If pos > 0 Then
            s2 = left$(s1, pos) + Mid$(s1, pos + 2)
            s1 = s2
        End If
    Loop Until pos = 0
    
    STR_LaisserUnSeulBlanc = s1
    Exit Function
    
    deja_bl = False
    Do While Len(s1)
        sc = left$(s1, 1)
        If sc = " " Then
            If Not deja_bl Then
                deja_bl = True
                s2 = s2 + sc
            End If
        ElseIf sc <> " " Then
            s2 = s2 + sc
            deja_bl = False
        End If
        s1 = Mid$(s1, 2)
    Loop
    
    STR_LaisserUnSeulBlanc = s2
    
End Function

Public Function STR_Supprimer(ByVal v_str As String, _
                              ByVal v_str_a_sup As String) As String

    Dim pos As Integer
    Dim stmp As String
    
    stmp = v_str
    pos = 1
    Do While pos > 0
        pos = InStr(stmp, v_str_a_sup)
        If pos > 0 Then
            stmp = left$(stmp, pos - 1) + Mid$(stmp, pos + Len(v_str_a_sup))
        End If
    Loop
    STR_Supprimer = stmp
    
End Function
