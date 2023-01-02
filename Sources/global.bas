Attribute VB_Name = "Module1"
Option Explicit

Public p_NumSite As String

Public Type TB_VAL_OBJET
    numobj As Integer
    valeur As String
    PourSQL As String
    NumFiltre As Long
    numchp As Integer
    typchp As String
    nb As Integer
End Type
Public Type TB_RESULT_SERV
    numsrv As Integer
    libsrv As String
    valeurs() As TB_VAL_OBJET
    nombre As Integer
    niveau As Integer
    est_donnée As Boolean
    srvNumPere As Integer
    PourSQL As String
End Type
Public P_tb_serv() As TB_RESULT_SERV
Public New_tb_serv() As TB_RESULT_SERV

Public Type TB_RESULT_ENTITE
    NumEntite As Integer
    LibEntite As String
    valeurs() As TB_VAL_OBJET
    nombre As Double
    niveau As Integer
    NiveauSH As String ' Hier ou Structure
    est_donnée As Boolean
    EntNumPere As Integer
    PourSQL As String
    Pour_Cnd_RP As String
    lstCnd As String
    NumFiltre As Long
End Type
Public P_UB_tbListeVals As Integer
Public Type TB_LISTE_VALS
    ListeNum As String
    VC_Num As String
End Type
Public P_tbListeVals() As TB_LISTE_VALS
Public P_Faire_tbListeVals As Boolean
Public P_tb_Entite() As TB_RESULT_ENTITE
Public P_bool_tb_Entite_ya_valeurs As Boolean
Public P_tb_Une_Entite() As TB_RESULT_ENTITE
Public New_tb_Entite() As TB_RESULT_ENTITE

Public Type TB_RESULT_HIERARCHIE
    numH As Integer
    libH As String
    valeurs() As TB_VAL_OBJET
    nombre As Integer
    niveau As Integer
    HPere As Integer
End Type
Public P_tb_hierar() As TB_RESULT_HIERARCHIE
Public New_tb_hierar() As TB_RESULT_HIERARCHIE

Public Type TB_RESULT_LISTE
    numvc As Integer
    libvc As String
    valeurs() As TB_VAL_OBJET
    nombre As Integer
    niveau As Integer
End Type
Public P_tb_liste() As TB_RESULT_LISTE
Public New_tb_liste() As TB_RESULT_LISTE

Public Type TB_RESULT_COLONNES
    valeur As String
    chpnum As Integer
    une_vc As String
    valtot As String
    leX As Integer
    nbCols As Integer
    leY As Integer
    lib As String
    libPere As String
    niveau As Integer
    nbtot As Integer
End Type
Public P_total_colonnes() As TB_RESULT_COLONNES
Public P_total_colonnes_Tmp() As TB_RESULT_COLONNES
Public P_bool_tbColonnes As Boolean

Public p_XHG As String, p_YHG As String, p_xBD As String, p_YBD As String

Public b_Chargement_Termine As Boolean
Public p_chemin_fichier_liens As String
Public p_numdoc_liens As String
Public p_nomdocument_encours As String
Public p_demander_titre As Boolean
Public p_requete_encours As String
Public p_nummodele_encours As Integer
Public p_numdoc_encours As Integer

Public p_VersionExcel As String
Public p_ExtensionXls As String
Public p_PointExtensionXls As String
Public Exc_obj_Version As String

Public grpnum As Boolean

Public p_ModeAssistant As Boolean

Public p_smtp_adrsrv As String

Public p_dansExcel As Boolean
Public p_dansGrid As Boolean

Public p_bool_ModeDebug As Boolean

Public p_version_KaliRP As String
Public P_MODE_DEBUG As Boolean

Public p_LibLienDétail As String

Public p_ANC_numfiltre_encours As Integer
Public p_ANC_numindice_encours As Integer
Public p_numfiltre_encours As Integer
Public p_numindice_encours As Integer
Public p_numfor_encours As Integer
Public p_MenFLigCol As String
Public p_ChpType As String

Public p_derchamp As Long
Public p_derannée As String

Public Const IMG_SOMME = 2
Public Const IMG_CHAMP = 5
Public Const IMG_BOULE = 8
Public Const IMG_BOULEBC = 9
Public Const IMG_BOULEBF = 10
Public Const IMG_CHP_DET = 11
Public Const IMG_CHP_GOMME = 12
Public Const IMG_CHP_LOUPER = 14
Public Const IMG_CHP_LOUPEB = 13
Public Const IMG_BOULEBC_PLUS = 15
Public Const IMG_BOULEBF_PLUS = 16
Public Const IMG_BOULE_ERREUR = 17
Public Const IMG_SQL_SELECT_C = 22
Public Const IMG_SQL_SELECT_CPLUS = 23
Public Const IMG_SQL_SELECT_F = 24
Public Const IMG_SQL_SELECT_FPLUS = 25
Public Const IMG_RESULTAT_RAPPORT = 26

Public p_TbFenetres()

' Pour ImgTypChp
Public Const IMG_LOUPE = 1
Public Const IMG_TYPECHP_SERVICE = 0
Public Const IMG_TYPECHP_FONCTION = 1
Public Const IMG_TYPECHP_SELECT = 2
Public Const IMG_TYPECHP_HIERARCHIE = 3
Public Const IMG_TYPECHP_TEXT = 4
Public Const IMG_TYPECHP_CHECK = 5
Public Const IMG_TYPECHP_RADIO = 6
Public Const IMG_TYPECHP_DATE = 7
Public Const IMG_TYPECHP_ENTIER = 8

' Pour le mode Détail de champs
Public p_LeXMaxPourGrdCell As Long
Public p_LeYMaxPourGrdCell As Long
Public p_LeIndexFenetreExcel As Integer
Public p_LeTypeTitreOuChamp As String
Public p_FaireHyperLienListeChamp As Boolean
Public p_MettreCommentListeChamp As Boolean
Public p_LeX_PourHyperlienG As Integer
Public p_LeX_PourHyperlienD As Integer
Public p_LeY_PourHyperlien As Integer
Public p_LeIndexFeuille_PourHyperlien As Integer

Public p_bPlusDeQuestion As Boolean

' pour savoir si le tableau est vide
Public p_bool_tbl_cond As Boolean
Public p_bool_tbl_condCHP As Boolean
Public p_bool_tbl_fichExcel As Boolean
Public p_bool_tbl_cell As Boolean
Public p_bool_tbl_fenExcel As Boolean
Public p_bool_forcer_vider_temp As Boolean  ' forcer la réinitialisation du Temp à partir du modèle si modification des fenêtres
Public p_bool_tbl_rdoF As Boolean
Public p_bool_tbl_rdoL As Boolean
Public p_bool_tbl_Demande As Boolean
Public p_bool_tbl_FichExcelOuverts As Boolean
Public p_bool_tbl_diff As Boolean

' pour les fenetres de détail
Public p_bool_tbl_detail As Boolean
Public Type G_STRUCT_DETAIL
    fornumG As Long
    donnumG As Long
    fornumD As Long
    donnumD As Long
End Type
Public p_tbl_detail() As G_STRUCT_DETAIL

Public p_Mode_FctTrace As Boolean
Public p_Chemin_FichierTrace As String

Public p_boolRetournerAuParam As Boolean
Public p_bool_Faire_VerifSauve As Boolean

Public p_TraitPublier As String

Public p_Appel_Création_Nouveau_Modele As Boolean

' pour effectuer une simulation d'une cellule
Public p_Simul_IFen As Integer
Public p_Simul_ITab As Integer

Public p_nomBDD_ODBC As String
Public p_nomBDD_SERVEUR As String

' numéro des fenêtres à rafraichir
Public p_Version As String
Public p_Imax As String
Public p_StrVersion As String

Public p_ListeRafraichirFenetre As String
Public p_ListeRafraichirTropX As Integer
Public p_ListeRafraichirTropY As Integer

Public p_NumTemp As Integer

Public p_i_tabExcel_pour_Copie As Integer
Public p_i_tabExcel As Integer
Public p_i_tabExcel_à_Relier As Integer
Public p_i_feuilleExcel As Integer

Public p_Derniere_MenFormeNonListe As String
Public p_Derniere_MenFormeListe As String

Public p_BoolMettreComment As Boolean

Public Const P_SUPER_UTIL = 1

Public g_modeSQL_LIB As String

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public p_bfaire_RowColChange As Boolean

Public p_NumForm As Integer
Public p_nummodele As Integer
Public p_NumUtil As Integer
Public p_SuperUser As Integer
Public p_param_supplementaires As Boolean
Public p_CodeUtil As String
Public p_tabdoc_present() As Long

Global Exc_obj As Excel.Application
Global Exc_wrk As Excel.Workbook
Public Const ColHightLigth = &H8080FF
Public Const ColLowLigth = &H8000000A
Public v_drive As String

' Imprimante PostScript déclarée : on peut faire du pdf
Public p_GererPDF As Boolean
Public p_imp_postscript As String
Public p_emplacement_ghostscript As String
Public p_maj_liens_OOO As Integer
Public p_numlabo As Long

Public p_appli_kalidoc As Long

' Indique si la conversion HTML est active ou non
Public p_ConvHtmlActif As Boolean

Public p_tblu() As Long
Public p_tblu_sel() As Long
Public p_siz_tblu As Long
Public p_siz_tblu_sel As Long
'Public p_nblicmax As Integer

Public Const P_DODS_RESP_NUMUTIL = 0
Public Const P_DODS_RESP_PRINCIPAL = 1
Public Const P_DODS_RESP_PRMAUTOR = 2
Public Const P_DODS_RESP_CRAUTOR = 3
Public Const P_DODS_RESP_REMPLACE = 4
Public Const P_DODS_RESP_INFORME = 5

' Pour P_AfficherArborescenceDoc
Public Type G_STRUCT_DOS
    numDos As Long
    ordre As Long
    titre As String
    numpere As Long
End Type

' Tableau des diffusions
Public Type SDIFFUSION
    nomdoc As String
    CheminDoc As String
    NumDest As Integer
    nomdest As String
    numdoc As Integer
    Diffusé As Boolean
    DiffàFaire As String
End Type
Public p_tbl_diff() As SDIFFUSION

' Tableau des diffusions
Public Type SINDEX
    NumFiltre As Integer
    NumIndice As Integer
    NumIndex As Integer
End Type
Public p_tbl_index() As SINDEX

' Chemins où se trouvent l'appli et ses fichiers
'Public p_chemin_appli As String
'Public p_nomini As String
'Public p_chemin_modele As String
'Public p_chemin_archive As String
'Public p_cheminkalidoc_serveur As String
'Public p_est_le_serveur As Boolean
Public p_NumDocs As Long

Public Type FichExcelOuverts
    FichFullname As String
    FichName As String
    FichModifié As Boolean
    FichVisible As Boolean
    FichàSauver As Boolean
End Type
Public p_tbl_FichExcelOuverts() As FichExcelOuverts
Public p_tbl_FichExcelPublier() As FichExcelOuverts

Private Type CELL
    CellFeuille As Integer
    CellX As Integer
    CellY As Integer
    CellTag As String
    CellLink As String
    CellPortee As String
    cellXPère As Integer
    cellYPère As Integer
    cellSQL As String
    cellNumFiltre As String
End Type
Public tbl_cell() As CELL

Private Type SELECTION
    SelFenNum As Integer
    SelX As Integer
    SelY As Integer
    SelXD As Integer
    SelYD As Integer
    SelXF As Integer
    SelYF As Integer
    Sel_ItabExcel As Integer
End Type
Public tb_Selection() As SELECTION

Public p_Indice_Grid_ChpCond As Integer

Public p_estV4 As Boolean
Public mode_Sites As Boolean

Public Type SFICH_PARAM_EXCEL
    CmdType As String
    CmdFenNum As String
    CmdX As String
    CmdY As String
    CmdFormNum As String
    CmdFormIndice As Integer
    CmdChpNum As Integer
    CmdCondition As String
    CmdConditionSQL As String
    cmdTypeChp As String
    CmdMenFormeChp As String
    CmdLstFen As String
    CmdLstDest As String
    CmdTitreDoc As String
    CmdMenFormeDoc As String
    CmdChpIndice As String
    CmdChpGridChargé As String
    CmdChpGridIndice As Integer
    CmdChpSQL As String
    CmdChpRelierà As String
    CmdNiveauRelier As String
    CmdX_Debut As String
    CmdY_Debut As String
    CmdX_Fin As String
    CmdY_Fin As String
End Type
Public tbl_fichExcel() As SFICH_PARAM_EXCEL

Public Type LISTE_ENTITE
    EntFNum As Integer
    EntFNumPere As Integer
    EntFOrdre As Integer
    EntFNiveau As Integer
    EntFLibLong As String
    EntFLibCourt As String
    EntFNiveauSH As String
    EntFRetenue As Boolean
    EntFArbor As String
End Type
Public Type LISTES_ENTITES
    EntPType As String
    EntPNum As Integer
    EntPDates As String
    EntPNom As String
    EntiChpNum As Integer
    EntPEntites() As LISTE_ENTITE
End Type
Public tbl_LesListes_Entites() As LISTES_ENTITES
Public p_bool_tbl_LesListes_Entites As Boolean

Public Type SFEN_EXCEL
    FenNumSave As Integer
    FenNum As Integer
    FenNom As String
    FenLoad As Boolean
    FenColMax As Integer
    FenRowMax As Integer
    FenModif As Boolean
    FenItbl_fichExcel As Integer
End Type
Public tbl_fenExcel() As SFEN_EXCEL

' pour les GRIDs
Public Const GrdForm_FF_Num = 0
Public Const GrdForm_FF_Image = 1
Public Const GrdForm_FF_Lib = 2
Public Const GrdForm_FF_Titre = 3
Public Const GrdForm_FF_NumIndice = 4

Public Type RDOF
    RDOF_num As Integer
    RDOF_rdoresultset As rdoResultset
    RDOF_sql As String
    RDOF_FormNum As String
    RDOF_FormNumS As String
    RDOF_FormIndice As String
    RDOF_etat As String
    RDOF_QuestionsFait As Boolean
    RDOF_QuestionsSQL As String
    RDOF_Q_RP As String
    RDOF_Q_FR As String
    RDOF_AussiQuestionsFait As String
    RDOF_AussiQuestionsSQL As String
    RDOF_Aussi_Q_RP As String
    RDOF_AussiChpNum As String
    RDOF_AussiChpType As String
    RDOF_AussiFctValid As String
    RDOF_Aussi_iDem As String
End Type
Public tbl_rdoF() As RDOF

Public Type RDOL
    RDOL_num As Integer
    RDOL_Opérateur As String
    RDOL_sqlenFrancais As String
    RDOL_sqlPasFrancais As String
    RDOL_sqlenSQL As String
    RDOL_fornum As String
    RDOL_FormIndice As String
    RDOL_DéjàPenCompte As Boolean
End Type
Public tbl_rdoL() As RDOL

Public Type SCOND_PARAM
    CondNumFiltre As Integer
    CondFormIndice As Integer
    CondLigneDansGrid As Integer
    CondString As String
    CondOper As String
    CondType As String
    CondFctValid As String
    CondBoolDetail As Boolean
    CondPasFrancais As String
    CondenFrancais As String
    CondenSQL As String
    CondligGrdCond As Integer
End Type
Public tbl_cond() As SCOND_PARAM

' Pour les conditions ajoutées sur un champ
Public Type SCOND_CHAMP
    CondChpIndiceGrid As Integer    ' indice du grid des conditions
    CondChpCndFormNum As String     ' numéro du formulaire
    CondChpChpFormIndice As Integer ' numéro de l'indice du filtre
    CondChpITabFichExcel As Integer ' Indice dans Tab_FichExcel
    CondChpCndChpNum As String      ' numéro du Champ
    CondChpCndChpOrdre As String    ' ordre pour champ détail
    CondChpCndOper As String        ' opérateur
    CondChpCndVal As String         ' type du champ
    CondChpCndBoolDetail As String  ' champ détaillé ?
    CondChpCndPasFrancais As String ' condition presque SQL
    CondChpCndenFrancais As String    ' fonction Français
    CondChpCndenSQL As String       ' condition vraie SQL
    CondChpCndOrigine As String     ' condition Stockée
End Type
Public tbl_condChp() As SCOND_CHAMP

' Pour les questions à poser à chaque fois
Public Type DEMANDE_SQL
    DemandChpNum As Integer
    DemandFormInd As Integer
    DemandGlobale As String
    DemandForNum As Integer
    DemandFFNum As Integer
    DemandChpStr As String
    DemandChpStrPlus As String
    DemandForStr As String
    DemandType As String
    DemandValeursPossibles As Integer
    DemandFctValid As String
    DemandPasFrancais As String
    DemandenFrancais As String
    DemandenSQL As String
    DemandFait As Boolean
    DemandFouA As String
    DemandAussiBool As Boolean
    DemandAussiStr As String
End Type
Public tbl_Demande() As DEMANDE_SQL

' CndPourLiens => numfiltre # numchp # Condition1 § Condition2
Public Type COND_POUR_LIEN
    CPL_ff_num As Integer
    CPL_I_TabExcel As Integer
    CPL_chp_num As Integer
    CPL_cond_RP As String
    CPL_cond_SQL As String
    CPL_cond_FR As String
    CPL_cond_Type As String
    CPL_cond_I As Integer
End Type
Public p_tblCondPourLien() As COND_POUR_LIEN

Public Const Public_Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public p_Excel_Decimal_Separator As String

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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DM_WIDTH = &H80000
Private Const DM_HEIGHT = &H100000
Private Const WM_DEVMODECHANGE = &H1B
Private Const HWND_BROADCAST = &HFFFF&
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


Public Function estV4()
    Dim Version As String
    
    Call Odbc_RecupVal("select v_kaliweb from version", Version)
    estV4 = Mid(Version, 1, 1) = "4"
End Function



Public Function Rep_Documents(ByVal sCle As String, ByVal AncCle As String, ByRef NewCle As String) As String
    Dim lret As Long, IDL As ITEMIDLIST, sPath As String
    Dim msg As String, s As String, sC As String

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
        If AncCle <> "" Then MsgBox "Mauvaise syntaxe pour " & sCle & " => " & AncCle
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
    
    sC = "CSIDL_PERSONAL"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_PERSONAL, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "CSIDL_LOCAL_APPDATA"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_LOCAL_APPDATA, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "CSIDL_COMMON_DOCUMENTS"
    lret = SHGetSpecialFolderLocation(100&, CSIDL_COMMON_DOCUMENTS, IDL)
    sPath = String$(512, Chr$(0))
    lret = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    s = left$(sPath, InStr(sPath, Chr$(0)) - 1)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "USERPROFILE"
    s = Environ$(sC)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "TEMP"
    s = Environ$(sC)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "APPDATA"
    s = Environ$(sC)
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
    sC = "KaliDoc"
    s = p_chemin_appli
    Call CL_AddLigne(sC & vbTab & s, 0, sC, True)
    
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

Public Function P_Exc_DecimalSeparator() As Boolean
    
    On Error GoTo Err_Excel_Separator
    p_Excel_Decimal_Separator = Exc_obj.DecimalSeparator
    GoTo LabOK
Err_Excel_Separator:
    p_Excel_Decimal_Separator = "."
    Resume Next
LabOK:
    On Error GoTo 0
End Function



Public Sub StartProcess(ByVal sFile As String, Optional ByVal sParameters As String = vbNullString)
    Dim ret As Integer
    
    ret = ShellExecute(0&, "open", sFile, sParameters, vbNullString, 1&)
End Sub

Public Sub MetResolutionEcran(sgWidth As Single, sgHeight As Single)
    Dim blTMP As Boolean
    Dim lgTMP As Long
    lgTMP = 0
    Do
        blTMP = EnumDisplaySettings(0, lgTMP, dmEcran)
        lgTMP = lgTMP + 1
    Loop Until Not blTMP
    dmEcran.dmFields = DM_WIDTH Or DM_HEIGHT
    dmEcran.dmPelsWidth = sgWidth
    dmEcran.dmPelsHeight = sgHeight
    lgTMP = ChangeDisplaySettings(dmEcran, 0)
    Call SendMessage(HWND_BROADCAST, WM_DEVMODECHANGE, 0, 0)
End Sub

Public Sub GetResolutionEcran(ByRef r_sgWidth As Single, ByRef r_sgHeight As Single)
    Dim blTMP As Boolean
    Dim lgTMP As Long
    lgTMP = 0
    Do
        blTMP = EnumDisplaySettings(0, lgTMP, dmEcran)
        lgTMP = lgTMP + 1
    Loop Until Not blTMP
    dmEcran.dmFields = DM_WIDTH Or DM_HEIGHT
    r_sgWidth = dmEcran.dmPelsWidth
    r_sgHeight = dmEcran.dmPelsHeight
    'lgTMP = ChangeDisplaySettings(dmEcran, 0)
    'Call SendMessage(HWND_BROADCAST, WM_DEVMODECHANGE, 0, 0)
End Sub

Public Function ScreenResolution(ByRef r_Largeur_Ecran As Single, ByRef r_Hauteur_Ecran As Single) As String
    
    r_Largeur_Ecran = GetSystemMetrics(SM_CXSCREEN)
    r_Hauteur_Ecran = GetSystemMetrics(SM_CYSCREEN)
    ScreenResolution = "Vidéo " & r_Largeur_Ecran & " x  " & r_Hauteur_Ecran
End Function


Public Function Appel_Aide()
    Dim CheminFichierAide As String
    Dim bExisteEnLocal As Boolean
    Dim strHTTP As String
    Dim CheminServeur As String, CheminLocal As String, nomIn_Chemin As String, nomIn_Fichier As String
    Dim nomIn_Extension As String, nomInCpy As String, Session As String
    Dim iret As Integer
    Dim FichServeur As String, FichLocal As String
    Dim liberr As String

    CheminFichierAide = p_CheminRapportType & "/Aide/Aide.pdf"
    If FICH_FichierExiste(CheminFichierAide) Then
        bExisteEnLocal = True
        P_FctOuvrirFichier (CheminFichierAide)
    Else
        MsgBox "Impossible de charger le fichier " & CheminFichierAide
    End If

End Function

Public Function P_FctOuvrirFichier(ByVal v_CheminFichier As String)
    
    If FICH_FichierExiste(v_CheminFichier) Then
        StartProcess v_CheminFichier
    Else
        MsgBox "Fichier " & v_CheminFichier & " introuvable"
    End If

End Function


Public Function P_AfficherArborescenceDoc(ByRef v_tv As TreeView, _
                                          ByVal v_numdos As Long, _
                                          ByVal v_img_dos As Long, _
                                          ByVal v_img_dos_sel As Long, _
                                          ByVal v_expand As Boolean) As Integer

    Dim sql As String, docs_titre As String, slien As String
    Dim trouve As Boolean, fracine As Boolean, encore As Boolean
    Dim I As Integer, n As Integer, mode As Integer
    Dim sav_numdos As Long, numdocs As Long
    Dim ordre As Long, numlien As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, nd2 As Node
    Dim tbldos() As G_STRUCT_DOS
    
    n = -1
    sav_numdos = v_numdos
    
    encore = True
    While encore
        trouve = True
        On Error GoTo lab_no_dos
        Set nd = v_tv.Nodes("S" & v_numdos)
        On Error GoTo 0
        If trouve Then
            encore = False
        Else
            n = n + 1
            ReDim Preserve tbldos(n) As G_STRUCT_DOS
            tbldos(n).numDos = v_numdos
            sql = "select DS_DONum, DS_SLien, DS_Titre, DS_Numpere, DS_Ordre" _
                & " from Dossier" _
                & " where DS_Num=" & v_numdos
            If Odbc_RecupVal(sql, numdocs, slien, tbldos(n).titre, v_numdos, tbldos(n).ordre) = P_ERREUR Then
                P_AfficherArborescenceDoc = P_ERREUR
                Exit Function
            End If
            If slien <> "" Then
                numlien = Mid$(slien, 2)
                Select Case left$(slien, 1)
                Case "O"
                    sql = "select DO_Titre" _
                        & " from Documentation" _
                        & " where DO_Num=" & numlien
                Case "S"
                    sql = "select DS_Titre" _
                        & " from Dossier" _
                        & " where DS_Num=" & numlien
                Case "D"
                    sql = "select D_Titre" _
                        & " from Document" _
                        & " where D_Num=" & numlien
                End Select
                If Odbc_RecupVal(sql, tbldos(n).titre) = P_ERREUR Then
                    P_AfficherArborescenceDoc = P_ERREUR
                    Exit Function
                End If
            End If
            If v_numdos = 0 Then
                encore = False
            End If
        End If
    Wend
    
    If v_numdos = 0 Then
        If TV_NodeExiste(v_tv, "O" & numdocs, nd) = P_OUI Then
            fracine = False
        Else
            fracine = True
        End If
    Else
        Set nd = v_tv.Nodes("S" & v_numdos)
        fracine = False
    End If
    ' On redébobine les dossiers
    For I = n To 0 Step -1
        On Error GoTo lab_no_dos
        trouve = True
        Set ndp = v_tv.Nodes("S" & tbldos(I).numDos)
        On Error GoTo 0
        If Not trouve Then
            If fracine Then
                mode = tvwChild
                If v_tv.Nodes.Count > 0 Then
                    If v_tv.Nodes(1).Root.Children > 0 Then
                        Set nd = v_tv.Nodes(1).Root
                        Do
                            If nd.tag > tbldos(I).ordre Then
                                mode = tvwPrevious
                                Exit Do
                            End If
                        Loop Until Not TV_NodeNext(nd)
                    End If
                End If
                If mode = tvwPrevious Then
                    Set nd = v_tv.Nodes.Add(nd, mode, "S" & tbldos(I).numDos, tbldos(I).titre, v_img_dos, v_img_dos_sel)
                Else
                    Set nd = v_tv.Nodes.Add(, mode, "S" & tbldos(I).numDos, tbldos(I).titre, v_img_dos, v_img_dos_sel)
                End If
                nd.tag = tbldos(I).ordre
            Else
                mode = tvwChild
                If nd.Children > 0 Then
                    Set nd2 = nd.Child
                    For n = 1 To nd.Children
                        If nd2.tag > tbldos(I).ordre Then
                            Set nd = nd2
                            mode = tvwPrevious
                            Exit For
                        End If
                        Set nd2 = nd2.Next
                    Next n
                End If
                Set nd = v_tv.Nodes.Add(nd, mode, "S" & tbldos(I).numDos, tbldos(I).titre, v_img_dos, v_img_dos_sel)
                nd.tag = tbldos(I).ordre
            End If
        Else
            Set nd = ndp
        End If
'        nd.Sorted = True
        If v_expand Then nd.Expanded = True
        fracine = False
    Next I
    
    P_AfficherArborescenceDoc = P_OK
    Exit Function

lab_no_dos:
    trouve = False
    Resume Next
    
End Function

Public Function Public_FichiersExcelOuverts(ByRef r_tbl_FichExcel() As FichExcelOuverts, v_Trait As String, v_chemin As String, v_visible As Boolean, v_à_Sauver As Boolean)
    Dim LaUbound As Integer
    Dim I As Integer
    Dim bDéjà As Boolean
    Dim strFichG As String, strFichd As String
    
    FctTrace ("Début Public_FichiersExcelOuverts")
    LaUbound = 0
    On Error GoTo Faire
    LaUbound = UBound(r_tbl_FichExcel(), 1) + 1
    For I = 0 To LaUbound
        strFichG = Replace(UCase(r_tbl_FichExcel(I).FichFullname), "\", "$")
        strFichG = Replace(strFichG, "/", "$")
        strFichd = Replace(UCase(v_chemin), "\", "$")
        strFichd = Replace(strFichd, "/", "$")
        If strFichG = strFichd Then
        'If r_tbl_FichExcel(i).FichFullname = v_Chemin Then
            bDéjà = True
            LaUbound = I
            Exit For
        End If
    Next I
Faire:
    On Error GoTo 0
    If Not bDéjà Then
        ReDim Preserve r_tbl_FichExcel(LaUbound)
        p_bool_tbl_FichExcelOuverts = True
    End If
    r_tbl_FichExcel(LaUbound).FichàSauver = v_à_Sauver
    r_tbl_FichExcel(LaUbound).FichVisible = v_visible
    r_tbl_FichExcel(LaUbound).FichFullname = v_chemin
    r_tbl_FichExcel(LaUbound).FichModifié = False
    FctTrace ("Après Public_FichiersExcelOuverts")
End Function

Public Function FctPoserQuestion(ByVal v_Trait As String, v_FF_Num As Integer, ByVal v_FF_Indice As Integer, ByVal v_i_tbl_RDOF As Integer, v_i_tbl_fichExcel As Integer, ByRef r_bPlusDeQuestion)
    Dim iD As Integer, nbdem As Integer
    Dim I As Integer
    Dim sql As String, rs As rdoResultset
    Dim nb As Integer, sqltmp As String
    Dim ForNum As Integer, chpnum As Integer
    Dim Frm As Form, ret As String
    Dim j As Integer
    Dim s As String
    Dim param1 As String, param2 As String, param3 As String
    Dim nomChp As String, oper As String, opSQL As String
    Dim ValChp As String
    Dim lib As String, iret As Integer
    Dim TagOper As String, TagVal As String
    Dim uneVal As String
    Dim nbQ As Integer
    Dim strQ As String
    Dim strQ_RDOF As String
    Dim bDemander As Boolean
    Dim NumForm As Long
    Dim NumFiltre As Long
    Dim leTitre As String
    Dim sQ As String
    Dim sC As String
    Dim ff_num As Integer, ff_indice As Integer, i_F As Integer
    Dim jj As Integer
    Dim nomS As String
    Dim ubCond As Integer
    Dim sRemp As String
    Dim sQ_Oper As String
    Dim bConcerne As Boolean
    Dim sqlRet As String
    Dim stmp As String
    Dim lstnum As Integer
    Dim i_tbl_RDOF As Long
    Dim laS As String
    Dim strserv As String
    Dim strfonction As String
    Dim IàDem As Integer
    Dim xID As Integer
    Dim iSc As Integer
    Dim sDem As String
    Dim iDD As Integer
    Dim sQAussi As String
    Dim n As Integer
    Dim unService As String
    Dim sql_PasF As String
    Dim strServices As String, opS As String
    Dim StrFct As String
    Dim opF As String
    Dim uneFct As String
    Dim sQAussi2 As String
    Dim sQ2 As String
    Dim sQTotal As String
    Dim sQObliger As String
    Dim s1 As String, s2 As String, s3 As String, s4 As String
    Dim nomformS As String
    Dim numfor As String
    
    ubCond = 0
    ' Voir s'il y a des questions a poser
    On Error GoTo err_TabDem
    nbQ = 0
    nbdem = UBound(tbl_Demande)
    GoTo SuiteDem
err_TabDem:
    Resume Fin
SuiteDem:
    On Error GoTo 0
    For iD = 0 To nbdem
        If InStr("LO*FO*FI", tbl_Demande(iD).DemandGlobale) = 0 Then
            tbl_Demande(iD).DemandGlobale = "LO"
        End If
    Next iD
Début:
    sQ = ""
    sQObliger = ""
    sQ2 = ""
    sQAussi = ""
    strQ = ""
    sQTotal = ""
    nbQ = 0
    
    For iD = 0 To nbdem
        bDemander = False
        If tbl_Demande(iD).DemandChpNum = -1 Then   ' supprimé
        Else
            ' si la question concerne un formulaire qui est celui du champ traité => on demande
            If v_Trait = "PARAM" Then
                ' retrouver le bon v_i_tbl_fichExcel
                sql = "select for_num,ff_fornums from formulaire,filtreform where formulaire.for_num = filtreform.ff_fornum " & " and filtreform.ff_num = " & p_numfiltre_encours
                NumFiltre = p_numfiltre_encours
            Else
                sql = "select for_num,ff_fornums from formulaire,filtreform where formulaire.for_num = filtreform.ff_fornum " & " and filtreform.ff_num = " & tbl_fichExcel(v_i_tbl_fichExcel).CmdFormNum
                NumFiltre = tbl_fichExcel(v_i_tbl_fichExcel).CmdFormNum
            End If
            Call Odbc_RecupVal(sql, NumForm, nomformS)
            ' on demande si LO et cet indice
            ' ou FI et meme filtre
            ' ou FO et meme formulaire
            'If tbl_Demande(iD).DemandGlobale = "FO" And InStr(nomformS & "*", NumForm) > 0 Then
            If tbl_Demande(iD).DemandGlobale = "FO" And tbl_Demande(iD).DemandForNum = NumForm Then
                bDemander = True
            ElseIf tbl_Demande(iD).DemandGlobale = "FI" And tbl_Demande(iD).DemandFFNum = NumFiltre Then
                bDemander = True
            ' demander si même indice (locale)
            ElseIf tbl_Demande(iD).DemandGlobale = "LO" Then
                If v_FF_Num & "_" & v_FF_Indice = tbl_Demande(iD).DemandFFNum & "_" & tbl_Demande(iD).DemandFormInd Then
                    'ElseIf tbl_Demande(iD).DemandFormInd = p_numindice_encours And tbl_Demande(iD).DemandForNum = p_numfor_encours Then
                    bDemander = True
                End If
            'ElseIf tbl_Demande(iD).DemandAussiBool Then
            '    If InStr(tbl_Demande(iD).DemandAussiStr, v_FF_Num & ":" & v_FF_Indice & ":") > 0 Then
            '        bDemander = True
            '    End If
            End If
            If v_Trait = "PARAM" Then
                If bDemander Then
                    If InStr(sQ, iD & ";") = 0 Then
                        sQ = sQ & iD & ";"
                    End If
                    If InStr(sQ2, iD & ";") = 0 Then
                        sQ2 = sQ2 & iD & ";"
                    End If
                End If
            Else
                If bDemander Then
                    If Not tbl_Demande(iD).DemandFait Then
                        If InStr(sQ, iD & ";") = 0 Then
                            sQ = sQ & iD & ";"
                        End If
                    End If
                    If InStr(sQ2, iD & ";") = 0 Then
                        sQ2 = sQ2 & iD & ";"
                    End If
                End If
            End If
        End If
    Next iD
    
    Call P_MAJ("")
    
    ' à ce niveau, sQ contient celles à demander, sQ2 contient celles qui le concernent en direct (demandées ou non)
    ' Voir dans tbl_rdoF si d'autre condition concernent ce filtre
    For iDD = 0 To nbdem
        If tbl_Demande(iDD).DemandChpNum >= 0 Then
            If tbl_Demande(iDD).DemandAussiStr <> "" Then
                For j = 0 To STR_GetNbchamp(tbl_Demande(iDD).DemandAussiStr, ";")
                    s = STR_GetChamp(tbl_Demande(iDD).DemandAussiStr, ";", j)
                    If s <> "" Then
                        If STR_GetChamp(s, ":", 0) = v_FF_Num And STR_GetChamp(s, ":", 1) = v_FF_Indice Then
                            ' c'est le bon filtre pour ce champ
                            i_tbl_RDOF = -1
                            For jj = 0 To UBound(tbl_rdoF())
                                If tbl_rdoF(jj).RDOF_num = v_FF_Num Then
                                    If tbl_rdoF(jj).RDOF_FormIndice = v_FF_Indice Then
                                        i_tbl_RDOF = jj
                                        For xID = 0 To STR_GetNbchamp(tbl_rdoF(jj).RDOF_Aussi_iDem, "¤")
                                            s1 = STR_GetChamp(tbl_rdoF(jj).RDOF_Aussi_iDem, "¤", xID)
                                            If s1 <> "" Then
                                                If s1 = iDD Then
                                                    If InStr(sQAussi, iDD & ";") = 0 Then
                                                        ' on l'ajoute si elle ni est pas déja (cas du global FO)
                                                        If InStr(sQAussi, iDD) = 0 Then
                                                            sQAussi = sQAussi & iDD & ";"
                                                            bDemander = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next xID
                                    End If
                                End If
                            Next jj
                        End If
                    End If
                Next j
            End If
        End If
    Next iDD
    
    If sQ2 <> "" Then
        For I = 0 To STR_GetNbchamp(sQ2, ";")
            If STR_GetChamp(sQ2, ";", I) <> "" Then
                iD = STR_GetChamp(sQ2, ";", I)
                sql = "select * from formetapechp where forec_num = " & tbl_Demande(iD).DemandChpNum
                If Odbc_SelectV(sql, rs) <> P_ERREUR Then
                    If Not rs.EOF Then
                        If tbl_Demande(iD).DemandChpStr = "" Then
                            tbl_Demande(iD).DemandChpStr = rs("forec_nom")
                        End If
                        TagOper = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 1)
                        TagOper = STR_GetChamp(TagOper, ":", 1)
                        TagVal = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 2)
                        TagVal = STR_GetChamp(TagVal, ":", 1)
                        If rs("forec_type") = "SELECT" Or rs("forec_type") = "CHECK" Or rs("forec_type") = "RADIO" Then
                            If TagVal <> "" Then
                                lstnum = rs("forec_valeurs_possibles")
                                TagVal = RecupCndLstVal(TagVal)
                                If Len(TagVal) > 100 Then
                                    TagVal = left(TagVal, 100) & " ..."
                                End If
                            End If
                        ElseIf rs("forec_type") = "HIERARCHIE" Then
                            Dim lst_nom  As String
                            If TagVal <> "" Then
                                n = STR_GetNbchamp(TagVal, ";")
                                TagVal = Replace(TagVal, "_DET", "")
                                TagVal = Replace(TagVal, "M", "")
                                If Mid(TagVal, Len(TagVal), 1) = ";" Then
                                    TagVal = Mid(TagVal, 1, Len(TagVal) - 1)
                                End If
                                sql = "select hvc_nom from hierarvalchp where hvc_num in (" & Replace(TagVal, ";", ",") & ")"
                                Call Odbc_SelectV(sql, rs)
                                TagVal = ""
                                While Not rs.EOF
                                    TagVal = TagVal + rs("hvc_nom").Value + " - "
                                    rs.MoveNext
                                Wend
                                rs.Close
                            End If
                        ElseIf rs("forec_fctvalid") = "%NUMSERVICE" Then
                            strServices = ""
                            opS = ""
                            If TagVal = "N0" Then ' Tout le site
                                strServices = "Tout le site"
                            Else
                                For j = 0 To STR_GetNbchamp(TagVal, ";")
                                    unService = STR_GetChamp(TagVal, ";", j)
                                    unService = Replace(unService, "S", "")
                                    unService = Replace(unService, ";", "")
                                    If unService <> "" Then
                                        If IsNumeric(unService) Then
                                            Call P_RecupSrvNom(unService, strserv)
                                            strServices = strServices & opS & strserv
                                        Else
                                            MsgBox "Service " & unService & " invalide"
                                            strServices = strServices & opS & ""
                                        End If
                                        opS = " OU "
                                    End If
                                Next j
                            End If
                            TagVal = strServices
                        ElseIf rs("forec_fctvalid") = "%NUMFCT" Then
                            StrFct = ""
                            opF = ""
                            For j = 0 To STR_GetNbchamp(TagVal, ";")
                                uneFct = STR_GetChamp(TagVal, ";", j)
                                uneFct = Replace(uneFct, "F", "")
                                uneFct = Replace(uneFct, ";", "")
                                If uneFct <> "" Then
                                    If IsNumeric(uneFct) Then
                                        Call P_RecupNomFonction(uneFct, strfonction)
                                        StrFct = StrFct & opF & strfonction
                                    Else
                                        MsgBox "Fonction " & uneFct & " invalide"
                                        StrFct = StrFct & opF & ""
                                    End If
                                    opF = " OU "
                                End If
                            Next j
                            TagVal = StrFct
                        End If
                        
                        If TagOper <> "" Or TagVal <> "" Then
                            If nbQ > 0 Then strQ = strQ & "," & Chr(13) & Chr(10) ' & Chr(13) & Chr(10)
                            strQ = strQ & tbl_Demande(iD).DemandChpStr & " " & tbl_Demande(iD).DemandChpStrPlus ' rs("Forec_Label")
                            strQ = strQ & " " & TagOper & " " & TagVal
                            nbQ = nbQ + 1
                        Else
                            If InStr(sQObliger, iD & ";") = 0 Then
                                If nbQ > 0 Then strQ = strQ & "," & Chr(13) & Chr(10) ' & Chr(13) & Chr(10)
                                strQ = strQ & tbl_Demande(iD).DemandChpStr & " " & tbl_Demande(iD).DemandChpStrPlus ' rs("Forec_Label")
                                sQObliger = sQObliger & iD & ";"
                                nbQ = nbQ + 1
                            End If
                        End If
                    End If
                    ' pour mettre le numero du champ
                    s1 = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 0)
                    If STR_GetNbchamp(s1, ":") = 2 Then
                        s2 = STR_GetChamp(s1, ":", 0) & ":" & tbl_Demande(iD).DemandChpNum & ":" & STR_GetChamp(s1, ":", 1) & "¤"
                        s2 = s2 & STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "|", 1) & "¤"
                        s2 = s2 & STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "|", 2) & "¤"
                    Else
                        s2 = tbl_Demande(iD).DemandPasFrancais
                    End If
                    nb = STR_GetNbchamp(tbl_Demande(iD).DemandPasFrancais, "¤")
                    If nb = 3 Then
                        tbl_Demande(iD).DemandPasFrancais = tbl_Demande(iD).DemandPasFrancais & "¤"
                    End If
                    Call AjouterCndPourLiens("Demande", iD, tbl_Demande(iD).DemandChpNum, tbl_Demande(iD).DemandFFNum, tbl_Demande(iD).DemandenSQL, s2, strQ)
                    'rs.Close
                End If
            End If
        Next I
    End If
    
    If sQAussi <> "" Then
        For I = 0 To STR_GetNbchamp(sQAussi, ";")
            If STR_GetChamp(sQAussi, ";", I) <> "" Then
                iD = STR_GetChamp(sQAussi, ";", I)
                For jj = 0 To UBound(tbl_rdoF())
                    If tbl_rdoF(jj).RDOF_num = v_FF_Num Then
                        If tbl_rdoF(jj).RDOF_FormIndice = v_FF_Indice Then
                            If tbl_rdoF(jj).RDOF_Aussi_iDem <> "" Then
                                For xID = 0 To STR_GetNbchamp(tbl_rdoF(jj).RDOF_Aussi_iDem, "¤")
                                    s1 = STR_GetChamp(tbl_rdoF(jj).RDOF_Aussi_iDem, "¤", xID)
                                    If s1 <> "" And InStr(sQ2, s1 & ";") = 0 Then
                                        If s1 = iD Then
                                            sql = "select * from formetapechp where forec_num = " & STR_GetChamp(tbl_rdoF(jj).RDOF_AussiChpNum, "¤", xID)
                                            If Odbc_SelectV(sql, rs) <> P_ERREUR Then
                                                If Not rs.EOF Then
                                                    TagOper = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 1)
                                                    TagOper = STR_GetChamp(TagOper, ":", 1)
                                                    TagVal = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 2)
                                                    TagVal = STR_GetChamp(TagVal, ":", 1)
                                                    If rs("forec_type") = "SELECT" Or rs("forec_type") = "CHECK" Or rs("forec_type") = "RADIO" Then
                                                        If TagVal <> "" Then
                                                            TagVal = RecupCndLstVal(TagVal)
                                                            If Len(TagVal) > 100 Then
                                                                TagVal = left(TagVal, 100) & " ..."
                                                            End If
                                                        End If
                                                    ElseIf rs("forec_type") = "HIERARCHIE" Then
                                                        n = STR_GetNbchamp(TagVal, ";")
                                                        If TagVal <> "" Then
                                                            TagVal = Replace(TagVal, "_DET", "")
                                                            TagVal = Replace(TagVal, "V", "")
                                                            If Mid(TagVal, Len(TagVal), 1) = ";" Then
                                                                TagVal = Mid(TagVal, 1, Len(TagVal) - 1)
                                                            End If
                                                            sql = "select hvc_nom from hierarvalchp where hvc_num in (" & Replace(TagVal, ";", ",") & ")"
                                                            Call Odbc_SelectV(sql, rs)
                                                            TagVal = ""
                                                            While Not rs.EOF
                                                                TagVal = TagVal + rs("hvc_nom").Value + " - "
                                                                rs.MoveNext
                                                            Wend
                                                            rs.Close
                                                        End If
                                                    ElseIf rs("forec_fctvalid") = "%NUMSERVICE" Then
                                                        strServices = ""
                                                        opS = ""
                                                        For j = 0 To STR_GetNbchamp(TagVal, ";")
                                                            unService = STR_GetChamp(TagVal, ";", j)
                                                            unService = Replace(unService, "S", "")
                                                            unService = Replace(unService, ";", "")
                                                            If unService <> "" Then
                                                                If IsNumeric(unService) Then
                                                                    Call P_RecupSrvNom(unService, strserv)
                                                                    strServices = strServices & opS & strserv
                                                                Else
                                                                    MsgBox "Service " & unService & " invalide"
                                                                    strServices = strServices & opS & ""
                                                                End If
                                                                opS = " OU "
                                                            End If
                                                        Next j
                                                        TagVal = strServices
                                                    End If
                                                    If TagOper <> "" Or TagVal <> "" Then
                                                        If nbQ > 0 Then strQ = strQ & "," & Chr(13) & Chr(10) ' & Chr(13) & Chr(10)
                                                        strQ = strQ & rs("forec_nom") & "  " & TagOper & " " & TagVal
                                                    End If
                                                    s2 = STR_GetChamp(tbl_rdoF(jj).RDOF_AussiQuestionsSQL, "¤", xID)
                                                    s2 = STR_GetChamp(tbl_rdoF(jj).RDOF_Aussi_Q_RP, "§", xID)
                                                    If Mid(s2, 1, 4) = "TBL_" Then
                                                        If InStr(sQObliger, iD & ";") = 0 Then
                                                            sQObliger = sQObliger & iD & ";"
                                                        End If
                                                    End If
                                                    nbQ = nbQ + 1
                                                End If
                                            End If
                                            rs.Close
                                        End If
                                    End If
                                Next xID
                            End If
                        End If
                    End If
                Next jj
            End If
        Next I
    End If
    
    If nbQ = 0 Then
        iret = 2
    Else
        If v_Trait = "PARAM" Then
            iret = 2
        ElseIf sQObliger <> "" Then
            iret = 2
        ElseIf Not r_bPlusDeQuestion Then
            ' trouver le titre
            leTitre = ""
            lib = ""
            For I = 0 To PiloteExcelBis.grdForm.Rows - 1
                If tbl_rdoF(v_i_tbl_RDOF).RDOF_num = PiloteExcelBis.grdForm.TextMatrix(I, GrdForm_FF_Num) Then
                    If tbl_rdoF(v_i_tbl_RDOF).RDOF_FormIndice = PiloteExcelBis.grdForm.TextMatrix(I, GrdForm_FF_NumIndice) Then
                        leTitre = vbCrLf & "-> filtre : " & PiloteExcelBis.grdForm.TextMatrix(I, GrdForm_FF_Lib) & " " & PiloteExcelBis.grdForm.TextMatrix(I, GrdForm_FF_Titre)
                        Exit For
                    End If
                End If
            Next I
            
            If p_nomdocument_encours <> "" Then
                lib = "Constitution du document " & p_nomdocument_encours & " " & vbCrLf
                lib = lib & "----------------------------------------------------- " & vbCrLf
            End If
            If nbQ = 1 Then
                lib = lib & "Confirmez vous la condition d'extraction pour " & leTitre & vbCrLf
                lib = lib & "----------------------------------------------------- " & vbCrLf
                lib = lib & strQ
            Else
                lib = lib & "Confirmez vous les " & nbQ & " conditions d'extraction pour " & leTitre & vbCrLf
                lib = lib & "--------------------------------------------------------------- " & vbCrLf
                lib = lib & strQ
            End If
            lib = lib & vbCrLf & "--------------------------------------------------------------- " & Chr(13) & Chr(10)
            'lib = lib & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Oui : ces conditions seront appliquées à ce filtre"
            'lib = lib & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Non : je souhaite modifier les conditions d'extraction"
            'lib = lib & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Annuler : ces conditions seront appliquées à tous les filtres suivants"
            Exc_obj.Visible = False
            ReDim tbl_libelle(3) As String
            ReDim tbl_tooltip(3) As String
            tbl_libelle(0) = "&OUI pour ce filtre"
            tbl_libelle(1) = "O&UI pour tous les filtres"
            tbl_libelle(2) = "&Modifier"
            tbl_libelle(3) = "&Abandonner"
            tbl_tooltip(0) = "Ces conditions seront appliquées à ce filtre"
            tbl_tooltip(1) = "Ces conditions seront appliquées à tous les filtres suivants"
            tbl_tooltip(2) = "Je souhaite modifier les conditions d'extraction"
            Set Frm = Com_Message
            iret = Com_Message.AppelFrm(lib, _
                                      "", _
                                      tbl_libelle(), _
                                      tbl_tooltip())
            Set Frm = Nothing
            If iret = 3 Then
                FctPoserQuestion = "ANNULER"
                Exit Function
            End If
'            iret = MsgBox(lib, vbDefaultButton1 + vbQuestion + vbYesNoCancel, "Critères d'extraction")
            If iret = 1 Then
                r_bPlusDeQuestion = True
            End If
        End If
    End If
        
    If iret = 2 Then
        ' Pour ce filtre, cet indice, voir si des questions sont à poser dans tbl_demande
        sQTotal = ""
        sQ = sQ & ";" & sQ2 & ";" & sQObliger & ";" & sQAussi
        For I = 0 To STR_GetNbchamp(sQ, ";")
            s = STR_GetChamp(sQ, ";", I)
            If s <> "" Then
                If InStr(sQTotal, s) = 0 Then
                    sQTotal = sQTotal & s & ";"
                End If
            End If
        Next I
        Set Frm = PrmFormAction
        ret = PrmFormAction.AppelFrm(v_FF_Num, v_FF_Indice, leTitre, sQTotal, v_Trait)
        Set Frm = Nothing
        If sQObliger <> "" And v_Trait <> "PARAM" Then
            ' on refait une passe pour contrôler
            ret = MsgBox("Vous souhaitez annuler l'extraction", vbQuestion + vbYesNo + vbDefaultButton2, "")
            If ret = vbYes Then
                FctPoserQuestion = "ANNULER"
                Exit Function
            Else
                GoTo Début
            End If
        End If
        Call PiloteExcelBis.verifSiSauve
    End If
    
    If v_Trait = "PARAM" Then
        ' on a terminé
        Exit Function
    End If
    
    sqlRet = ""
    sql_PasF = ""
    opSQL = ""
    For I = 0 To STR_GetNbchamp(sQ2, ";")
        If STR_GetChamp(sQ2, ";", I) <> "" Then
            IàDem = STR_GetChamp(sQ2, ";", I)
            ' Est il concerné en direct ?
            Call FaitCondQuestion(tbl_Demande(IàDem).DemandPasFrancais, tbl_Demande(IàDem).DemandType, tbl_Demande(IàDem).DemandChpNum, sqlRet, sql_PasF, opSQL)
            Call AjouterCndPourLiens("Demande", IàDem, tbl_Demande(IàDem).DemandChpNum, tbl_Demande(IàDem).DemandFFNum, tbl_Demande(IàDem).DemandenSQL, tbl_Demande(IàDem).DemandPasFrancais, tbl_Demande(IàDem).DemandenFrancais)
        End If
    Next I
    For I = 0 To STR_GetNbchamp(sQAussi, ";")
        If STR_GetChamp(sQAussi, ";", I) <> "" Then
            IàDem = STR_GetChamp(sQAussi, ";", I)
            For n = 0 To UBound(tbl_rdoF())
                If n = v_i_tbl_RDOF Then
                    s = tbl_rdoF(n).RDOF_Aussi_iDem
                    If s <> "" Then
                        For j = 0 To STR_GetNbchamp(s, "¤")
                            s4 = STR_GetChamp(s, "¤", j)
                            If s4 <> "" And InStr(sQ2, s4 & ";") = 0 Then
                                If s4 = IàDem Then
                                    s1 = STR_GetChamp(tbl_rdoF(n).RDOF_Aussi_Q_RP, "§", j)   ' STR_GetChamp(tbl_rdoF(n).RDOF_AussiQuestionsSQL, "¤", j)
                                    s2 = STR_GetChamp(tbl_rdoF(n).RDOF_AussiChpType, "¤", j)
                                    s3 = STR_GetChamp(tbl_rdoF(n).RDOF_AussiChpNum, "¤", j)
                                    Call FaitCondQuestion(s1, s2, s3, sqlRet, sql_PasF, opSQL)
                                End If
                            End If
                        Next j
                    End If
                End If
            Next n
        End If
    Next I
    
    FctPoserQuestion = sqlRet
    
    If v_Trait <> "PARAM" Then
        tbl_rdoF(v_i_tbl_RDOF).RDOF_QuestionsFait = True
        tbl_rdoF(v_i_tbl_RDOF).RDOF_QuestionsSQL = sqlRet
        tbl_rdoF(v_i_tbl_RDOF).RDOF_Q_RP = sql_PasF
    End If
    
Fin:
    On Error GoTo 0
End Function

Public Function AjouterCndPourLiens(ByVal LeType As String, ByVal I As Long, chpnum As Integer, ByVal FFNum As Integer, ByVal sql As String, ByVal SQL_RP As String, ByVal SQL_FR As String)
    Dim ub1 As Integer, ub2 As Integer
    Dim leI As Integer
    Dim j As Integer, deja As Boolean
    Dim nb As Integer
    Dim str_SQL As String, str_FR As String, str_RP As String
    Dim s1 As String, s2 As String
        
    'Debug.Print leType & " " & i & " " & chpnum & " " & SQL_RP
    ' format : CHP:1067:e1_date_integration¤OP:SU¤VAL:01/01/2010¤* (le * indique toutes les fenetres)
    nb = STR_GetNbchamp(SQL_RP, "§")
    For leI = 0 To nb
        str_RP = STR_GetChamp(SQL_RP, "§", leI)
        If str_RP <> "" Then
            Call PiloteExcelBis.FaitConditionChamp(str_RP, str_SQL, str_FR)
            On Error GoTo Suite1
            ub1 = 0
            ub1 = UBound(p_tblCondPourLien) + 1
            GoTo Suite1
Suite1:
            On Error GoTo 0
            deja = False
            If ub1 > 0 Then
                For j = 0 To UBound(p_tblCondPourLien)
                    If p_tblCondPourLien(j).CPL_cond_Type = LeType And p_tblCondPourLien(j).CPL_I_TabExcel = I Then
                        deja = True
                        Exit For
                    End If
                Next j
            End If
            If Not deja Or LeType = "RDOF" Then
                ReDim Preserve p_tblCondPourLien(ub1)
            Else
                ub1 = j
            End If
            s1 = STR_GetChamp(SQL_RP, "¤", 0)
            If STR_GetNbchamp(s1, ":") = 2 Then
                s2 = STR_GetChamp(s1, ":", 0) & ":" & chpnum & ":" & STR_GetChamp(s1, ":", 1) & "¤"
                s2 = s2 & STR_GetChamp(SQL_RP, "¤", 1) & "¤"
                SQL_RP = s2 & STR_GetChamp(SQL_RP, "¤", 2) & "¤"
            End If
            p_tblCondPourLien(ub1).CPL_chp_num = chpnum
            p_tblCondPourLien(ub1).CPL_ff_num = FFNum
            p_tblCondPourLien(ub1).CPL_cond_SQL = str_SQL
            p_tblCondPourLien(ub1).CPL_cond_FR = str_FR
            p_tblCondPourLien(ub1).CPL_cond_RP = Replace(str_RP, "|", "¤")
            p_tblCondPourLien(ub1).CPL_cond_Type = LeType
            p_tblCondPourLien(ub1).CPL_I_TabExcel = I
        End If
    Next leI
End Function

Public Function AjouteOuRemplace(v_str As String, v_iDem As String, v_car As String)
    Dim I As Integer
    Dim sOut As String
    
    For I = 0 To STR_GetNbchamp(v_str, v_car) - 1
        If STR_GetChamp(v_str, v_car, I) = v_iDem Then
            AjouteOuRemplace = I
            Exit Function
        End If
    Next I
    AjouteOuRemplace = -1
End Function

Public Function AjouteCondition(v_str As String, v_cnd As String, v_car As String, v_i As Integer)
    Dim I As Integer
    Dim sOut As String
    
    If v_i = -1 Then    ' ajout
        sOut = v_str & IIf(v_str <> "", v_car, "") & v_cnd
    Else                ' remplace
        For I = 0 To STR_GetNbchamp(v_cnd, v_car)
            sOut = sOut & IIf(I > 0, v_car, "")
            If I = v_i Then
                sOut = sOut & v_cnd
            Else
                sOut = sOut & STR_GetChamp(v_str, v_car, I)
            End If
        Next I
    End If
    AjouteCondition = sOut
End Function

Public Function EnleveCondition(v_cnd As String, v_car As String, v_i As Integer)
    Dim I As Integer
    Dim sOut As String
    
    For I = 0 To STR_GetNbchamp(v_cnd, v_car) - 1
        sOut = sOut & IIf(I > 0, v_car, "")
        If I <> v_i Then
            sOut = sOut & STR_GetChamp(v_cnd, v_car, I)
        End If
    Next I
    EnleveCondition = sOut
End Function

Public Function P_MAJ(v_Trait As String)
    Dim iDem As Integer, nbdem As Integer, iD As Integer, idLig As Integer
    Dim nb As Integer
    Dim iSc As Integer
    Dim i_tbl_RDOF As Integer
    Dim s As String, sql As String
    Dim I As Integer
    Dim rs As rdoResultset
    Dim nomAnc As String
    Dim nomS As String
    Dim i_F As Integer
    Dim j As Integer
    Dim ret As Integer
    Dim z As Integer
    Dim laS As String
    Dim sqlRet As String, opSQL As String
    Dim ff_num As Long, ff_indice As Long
    Dim chpnum As String
    Dim chp_nom As String
    Dim chp_type As String, chp_fctvalid As String
    Dim s1 As String, s2 As String, s3 As String
    Dim s4 As String
    Dim laSPF As String, laSF As String
    Dim uCount As Integer
    
    uCount = 0
    On Error Resume Next
    uCount = UBound(P_tb_conditions)
    'Erase P_tb_conditions
    ' voir les tbl_demande faites (completes)
    For I = 0 To UBound(tbl_Demande())
        If tbl_Demande(I).DemandChpNum >= 0 Then
            s = tbl_Demande(I).DemandPasFrancais
            If InStr(s, "¤") = 0 Then
                ' à mettre au format : CHP:3:e1_service¤OP:¤NUMSERVICE:S3;
                s = "CHP:" & tbl_Demande(I).DemandChpNum & ":"
                sql = "Select forec_nom,forec_type,forec_fctvalid from formetapechp where forec_num=" & tbl_Demande(I).DemandChpNum
                Call Odbc_RecupVal(sql, chp_nom, chp_type, chp_fctvalid)
                s = s & chp_nom & "¤" & "OP:¤"
                If chp_type = "TEXT" Then
                    If chp_fctvalid = "%NUMSERVICE" Then
                        s = s & "NUMSERVICE"
                    ElseIf chp_fctvalid = "%NUMFCT" Then
                        s = s & "NUMFCT"
                    ElseIf InStr(chp_fctvalid, "%DATE") > 0 Then
                        s = s & "DATE"
                    End If
                    s = s & ":"
                End If
            End If
            ff_num = tbl_Demande(I).DemandFFNum
            s = reformer_BCR(ff_num, s)
            'Debug.Print i & " *************************"
            'Debug.Print i & " " & s
            s1 = STR_GetChamp(STR_GetChamp(s, "¤", 0), ":", 2)  ' nom chp
            s2 = STR_GetChamp(STR_GetChamp(s, "¤", 1), ":", 1)  ' operateur
            s3 = STR_GetChamp(STR_GetChamp(s, "¤", 2), ":", 1)  ' valeurs
            s4 = STR_GetChamp(s, "¤", 3)                        ' fenetres
            If s1 <> "" And s2 <> "" And s3 <> "" Then
                tbl_Demande(I).DemandFait = True
                '  vers CHP:86:e1_datedecl¤OP:COMPRIS¤VAL:01/01/2013 31/12/2013¤*
                If InStr(tbl_Demande(I).DemandFctValid, "%DATE") > 0 Then
                    s = "CHP:" & tbl_Demande(I).DemandChpNum & ":" & s1 & "¤OP:" & s2 & "¤DATE:" & s3 & "¤" & s4
                Else
                    s = "CHP:" & tbl_Demande(I).DemandChpNum & ":" & s1 & "¤OP:" & s2 & "¤VAL:" & s3 & "¤" & s4
                End If
                s = Replace(s, ":V", ":")
                s = Replace(s, ";V", ";")
                nb = STR_GetNbchamp(s, "¤")
                If nb = 3 Then
                    s = s & "¤"
                End If
                'Debug.Print s
                Call PiloteExcelBis.FaitConditionChamp(s, laSPF, laSF)
                Call AjouterCndPourLiens("Demande", -1 * I, tbl_Demande(I).DemandChpNum, tbl_Demande(I).DemandFFNum, laSPF, s, laSF)
                ReDim Preserve P_tb_conditions(uCount)
                P_tb_conditions(uCount).titre = laSF
                uCount = uCount + 1
                'Debug.Print laSPF
                'Debug.Print laSF
                tbl_Demande(I).DemandenFrancais = laSF
                tbl_Demande(I).DemandenSQL = laSPF
            End If
        End If
    Next I
    GoTo suite
    If v_Trait <> "" Then
        For j = 0 To STR_GetNbchamp(v_Trait, ";")
            s1 = STR_GetChamp(v_Trait, ";", j)
            If s1 <> "" Then
                For i_F = 0 To UBound(tbl_rdoF())
                    s2 = tbl_rdoF(i_F).RDOF_Aussi_iDem
                    If s2 <> "" Then
                        For I = 0 To STR_GetNbchamp(s2, "¤")
                            s3 = STR_GetChamp(s2, "¤", I)
                            If s3 = s1 Then
                                ' on ne remplace que celui là
                                s4 = tbl_rdoF(i_F).RDOF_Aussi_iDem
                                tbl_rdoF(i_F).RDOF_Aussi_iDem = EnleveCondition(s4, "¤", I)
                                
                                s4 = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL
                                tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = EnleveCondition(s4, "¤", I)
                                ' tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = Replace(s4, STR_GetChamp(s4, "¤", i), "TBL_DEMANDE:" & s1)
                                
                                s4 = tbl_rdoF(i_F).RDOF_AussiQuestionsFait
                                tbl_rdoF(i_F).RDOF_AussiQuestionsFait = EnleveCondition(s4, "¤", I)
                                ' tbl_rdoF(i_F).RDOF_AussiQuestionsFait = Replace(s4, STR_GetChamp(s4, "¤", i), "F")
                                
                                s4 = tbl_rdoF(i_F).RDOF_AussiChpNum
                                tbl_rdoF(i_F).RDOF_AussiChpNum = EnleveCondition(s4, "¤", I)
                                'tbl_rdoF(i_F).RDOF_AussiChpNum = Replace(s4, STR_GetChamp(s4, "¤", i), "")
                                s4 = tbl_rdoF(i_F).RDOF_AussiChpType
                                tbl_rdoF(i_F).RDOF_AussiChpType = EnleveCondition(s4, "¤", I)
                                'tbl_rdoF(i_F).RDOF_AussiChpType = Replace(s4, STR_GetChamp(s4, "¤", i), "")
                                s4 = tbl_rdoF(i_F).RDOF_AussiFctValid
                                tbl_rdoF(i_F).RDOF_AussiFctValid = EnleveCondition(s4, "¤", I)
                                'tbl_rdoF(i_F).RDOF_AussiFctValid = Replace(s4, STR_GetChamp(s4, "¤", i), "")
                            End If
                        Next I
                    End If
                Next i_F
            End If
        Next j
    End If

    'For i = 0 To UBound(tbl_Demande())
    '    s = i & " fait=" & tbl_Demande(i).DemandFait & " " & tbl_Demande(i).DemandForNum & "_" & tbl_Demande(i).DemandFFNum & "_" & tbl_Demande(i).DemandFormInd & " sql=" & tbl_Demande(i).DemandenSQL & " aussi=" & tbl_Demande(i).DemandAussiStr
    '    'Debug.Print s
    'Next i
    'For i = 0 To UBound(tbl_rdoF())
    '    s = i & " " & tbl_rdoF(i).RDOF_num & "_" & tbl_rdoF(i).RDOF_FormIndice & " RDOF_QuestionsSQL=" & tbl_rdoF(i).RDOF_QuestionsSQL & " ->RDOF_sql=" & tbl_rdoF(i).RDOF_sql
    '    'Debug.Print s
    '    s = " QaussiSQL=" & tbl_rdoF(i).RDOF_AussiQuestionsSQL & "-> idem=" & tbl_rdoF(i).RDOF_Aussi_iDem & " " & tbl_rdoF(i).RDOF_AussiChpNum & " chpnum=" & tbl_rdoF(i).RDOF_AussiChpNum
    '    'Debug.Print s
    'Next i
    
suite:
    nbdem = UBound(tbl_Demande)
    For iDem = 0 To nbdem
        If tbl_Demande(iDem).DemandChpNum = -1 Then GoTo Lab_Next_iDem

        If tbl_Demande(iDem).DemandAussiBool Then
            If tbl_Demande(iDem).DemandAussiStr <> "" Then
                ' ex : 183:1:3170;69:1:137;   ff_num  ff_indice   forec_num
                For iSc = 0 To STR_GetNbchamp(tbl_Demande(iDem).DemandAussiStr, ";")
                    s = STR_GetChamp(tbl_Demande(iDem).DemandAussiStr, ";", iSc)
                    If s <> "" Then
                        ' on applique à l'autre (RDO_F)
                        chpnum = STR_GetChamp(s, ":", 2)
                        ' remplacer le champ
                        ff_num = STR_GetChamp(s, ":", 0)
                        ff_indice = STR_GetChamp(s, ":", 1)
                        For i_F = 0 To UBound(tbl_rdoF())
                            If tbl_rdoF(i_F).RDOF_num = ff_num Then
                                If tbl_rdoF(i_F).RDOF_FormIndice = ff_indice Then
                                    Call Odbc_RecupVal("select forec_nom from formetapechp where forec_num=" & tbl_Demande(iDem).DemandChpNum, nomAnc)
                                    Call Odbc_RecupVal("select forec_nom from formetapechp where forec_num=" & chpnum, nomS)
                                    s4 = iDem
                                    'Debug.Print tbl_rdoF(i_F).RDOF_AussiQuestionsSQL
                                    ret = AjouteOuRemplace(tbl_rdoF(i_F).RDOF_Aussi_iDem, s4, "¤")      ' -1 on ajoute sinon on remplace
                                    tbl_rdoF(i_F).RDOF_Aussi_iDem = AjouteCondition(tbl_rdoF(i_F).RDOF_Aussi_iDem, s4, "¤", ret)
                                    
                                    s = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL
                                    If Not tbl_Demande(iDem).DemandFait Then
                                        'tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & "TBL_DEMANDE:" & iDem
                                        tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiQuestionsSQL, "TBL_DEMANDE:" & iDem, "¤", ret)
                                        tbl_rdoF(i_F).RDOF_Aussi_Q_RP = AjouteCondition(tbl_rdoF(i_F).RDOF_Aussi_Q_RP, "TBL_DEMANDE:" & iDem, "§", ret)
                                    Else
                                        'tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & Replace(tbl_Demande(iDem).DemandenSQL, nomAnc, nomS)
                                        tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiQuestionsSQL, Replace(tbl_Demande(iDem).DemandenSQL, nomAnc, nomS), "¤", ret)
                                        tbl_rdoF(i_F).RDOF_Aussi_Q_RP = AjouteCondition(tbl_rdoF(i_F).RDOF_Aussi_Q_RP, Replace(tbl_Demande(iDem).DemandPasFrancais, nomAnc, nomS), "§", ret)
                                    End If
                                    tbl_rdoF(i_F).RDOF_AussiQuestionsFait = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiQuestionsFait, IIf(tbl_Demande(iDem).DemandFait, "T", "F"), "¤", ret)
                                    tbl_rdoF(i_F).RDOF_AussiChpNum = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiChpNum, chpnum, "¤", ret)
                                    tbl_rdoF(i_F).RDOF_AussiChpType = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiChpType, tbl_Demande(iDem).DemandType, "¤", ret)
                                    tbl_rdoF(i_F).RDOF_AussiFctValid = AjouteCondition(tbl_rdoF(i_F).RDOF_AussiFctValid, tbl_Demande(iDem).DemandFctValid, "¤", ret)
                                    Exit For
                                End If
                            End If
                        Next i_F
                    End If
                Next iSc
            End If
        End If
Lab_Next_iDem:
    Next iDem
    
    'Debug.Print "FIN P_MAJ"
    'For i = 0 To UBound(tbl_rdoF())
    '    s = i & " " & tbl_rdoF(i).RDOF_num & "_" & tbl_rdoF(i).RDOF_FormIndice & " RDOF_QuestionsSQL=" & tbl_rdoF(i).RDOF_QuestionsSQL & " ->sql=" & tbl_rdoF(i).RDOF_sql
    '    Debug.Print s
    '    s = " QaussiSQL=" & tbl_rdoF(i).RDOF_AussiQuestionsSQL & "-> idem=" & tbl_rdoF(i).RDOF_Aussi_iDem & " " & tbl_rdoF(i).RDOF_AussiChpNum & " chpnum=" & tbl_rdoF(i).RDOF_AussiChpNum
    '    debug.Print s
    'Next i
End Function


Public Function FaitCondQuestion(ByVal v_DemandenSQL As String, ByVal v_DemandType As String, ByVal v_DemandChpNum As String, ByRef r_sqlRet, ByRef r_sql_PasF, ByRef r_opSQL)
    Dim param1 As String, param2 As String, param3 As String, sqltmp As String
    Dim nb As Integer, I As Integer
    Dim LeType As String
    Dim fctvalid As String
    Dim d1 As String, d2 As String
    Dim s As String
    Dim nomChp As String, nomS As String, oper As String, ValChp  As String, uneVal  As String, stmp As String, laS As String
    
        If v_DemandenSQL = "" Then Exit Function
        param1 = STR_GetChamp(v_DemandenSQL, "¤", 0)
        param2 = STR_GetChamp(v_DemandenSQL, "¤", 1)
        param3 = STR_GetChamp(v_DemandenSQL, "¤", 2)
        LeType = v_DemandType
        nomChp = STR_GetChamp(param1, ":", 1)
        If Odbc_RecupVal("select forec_nom, forec_type, forec_fctvalid from formetapechp where forec_num=" & v_DemandChpNum, nomS, LeType, fctvalid) <> P_ERREUR Then
            nomChp = nomS
        End If
        oper = STR_GetChamp(param2, ":", 1)
        'If Oper = "<" Then Oper = "<="
        'If Oper = ">" Then Oper = ">="
        ValChp = STR_GetChamp(param3, ":", 1)
        If LeType = "SELECT" Or LeType = "RADIO" Or LeType = "CHECK" Or LeType = "HIERARCHIE" Then
            If LeType = "HIERARCHIE" Then
                ValChp = Replace(ValChp, "_DET", "")
            End If
            r_sqlRet = r_sqlRet & r_opSQL & "("
            nb = STR_GetNbchamp(ValChp, ";")
            sqltmp = ""
            For I = 0 To nb - 1
                uneVal = STR_GetChamp(ValChp, ";", I)
                If oper = "=" Then
                    If I > 0 Then sqltmp = sqltmp & " Or "
                    If uneVal = "<NR>" Then
                        sqltmp = sqltmp & "(" & nomChp & " = '' or " & nomChp & " is null " & ")"
                    Else
                        sqltmp = sqltmp & "(" & nomChp & " like '%" & uneVal & ";%'" & ")"
                    End If
                ElseIf oper = "!" Then
                    If I > 0 Then sqltmp = sqltmp & " And "
                    If uneVal = "<NR>" Then
                        sqltmp = sqltmp & "(" & nomChp & " != '' or not " & nomChp & " is null " & ")"
                    Else
                        sqltmp = sqltmp & "(" & nomChp & " not like '%" & uneVal & ";%'" & ")"
                    End If
                End If
            Next I
            r_sqlRet = r_sqlRet & sqltmp & ")"
        Else
            'MsgBox " a voir"
            If STR_GetChamp(param3, ":", 0) = "DATE" Then
                If oper = "!" Then oper = "<>"
                If oper = "COMPRIS" Then
                    d1 = STR_GetChamp(ValChp, " ", 0)
                    d2 = STR_GetChamp(ValChp, " ", 1)
                    s = "( to_date(" & nomChp & ",'dd/mm/YYYY') " & ">=" & " '" & Mid(d1, 7, 4) & "-" & Mid(d1, 4, 2) & "-" & Mid(d1, 1, 2) & "'"
                    s = s & " And to_date(" & nomChp & ",'dd/mm/YYYY') " & "<=" & " '" & Mid(d2, 7, 4) & "-" & Mid(d2, 4, 2) & "-" & Mid(d2, 1, 2) & "')"
                    r_sqlRet = r_sqlRet & r_opSQL & s
                Else
                    r_sqlRet = r_sqlRet & r_opSQL & "( to_date(" & nomChp & ",'dd/mm/YYYY') " & oper & " '" & Mid(ValChp, 7, 4) & "-" & Mid(ValChp, 4, 2) & "-" & Mid(ValChp, 1, 2) & "')"
                End If
            ElseIf STR_GetChamp(param3, ":", 0) = "NUMSERVICE" Then
                'BoolLeDetail = False
                '
                Dim s_SQL As String, s_SQLB As String, s_F As String
                If oper = "" Then
                    oper = "="
                End If
                'BoolLeDetail = IIf(InStr(ValChp, "_DET") > 0, True, False)
                ValChp = Replace(ValChp, "S", "")
                r_sqlRet = r_sqlRet & r_opSQL & "("
                nb = STR_GetNbchamp(ValChp, ";")
                sqltmp = ""
                stmp = ""
                If ValChp = "N0" Then
                    sqltmp = sqltmp & "(" & nomChp & " like '" & "%#%'" & ")"
                Else
                    For I = 0 To nb - 1
                        uneVal = STR_GetChamp(ValChp, ";", I)
                        If uneVal <> "" Then
                            laS = FaitLstService(uneVal, uneVal)
                        End If
                        stmp = stmp & ";" & laS
                    Next I
                    nb = STR_GetNbchamp(stmp, ";")
                    For I = 0 To nb - 1
                        uneVal = STR_GetChamp(stmp, ";", I)
                        If uneVal <> "" Then
                            If oper = "=" Then
                                If sqltmp <> "" Then sqltmp = sqltmp & " Or " ' Ancienne version
                                If uneVal = "<NR>" Then
                                    sqltmp = sqltmp & "(" & nomChp & " = '' or " & nomChp & " is null " & ")"
                                Else
                                    sqltmp = sqltmp & "(" & nomChp & " like '" & uneVal & "#%'" & ")"
                                    'sqltmp = sqltmp & IIf(sqltmp = "", "", ",") & UneVal
                                End If
                            ElseIf oper = "!" Then
                                If I > 0 Then sqltmp = sqltmp & " And "
                                If uneVal = "<NR>" Then
                                    sqltmp = sqltmp & "(" & nomChp & " != '' or not " & nomChp & " is null " & ")"
                                Else
                                    sqltmp = sqltmp & "(" & nomChp & " not like '%" & uneVal & ";%'" & ")"   ' Ancienne version
                                    'sqltmp = sqltmp & IIf(sqltmp = "", "", ",") & UneVal
                                End If
                            End If
                        End If
                    Next I
                End If
                'If UneVal <> "<NR>" Then
                '    If sqltmp <> "" Then
                '        sqltmp = nomChp & " in (" & sqltmp & ")"
                '    End If
                'End If
                r_sqlRet = r_sqlRet & sqltmp & ")"
            ElseIf STR_GetChamp(param3, ":", 0) = "NUMFCT" Then
                sqltmp = ""
                ValChp = Replace(ValChp, "F", "")
                r_sqlRet = r_sqlRet & r_opSQL & "("
                nb = STR_GetNbchamp(ValChp, ";")
                For I = 0 To nb - 1
                    uneVal = STR_GetChamp(ValChp, ";", I)
                    If uneVal <> "" Then
                        If oper = "=" Then
                            If sqltmp <> "" Then sqltmp = sqltmp & " Or "
                            If uneVal = "<NR>" Then
                                sqltmp = sqltmp & "(" & nomChp & " = '' or " & nomChp & " is null " & ")"
                            Else
                                sqltmp = sqltmp & "(" & nomChp & " like '" & uneVal & "#%'" & ")"
                            End If
                        ElseIf oper = "!" Then
                            If I > 0 Then sqltmp = sqltmp & " And "

                            If uneVal = "<NR>" Then
                                sqltmp = sqltmp & "(" & nomChp & " != '' or not " & nomChp & " is null " & ")"
                            Else
                                sqltmp = sqltmp & "(" & nomChp & " not like '%" & uneVal & ";%'" & ")"
                            End If
                        End If
                    End If
                Next I
                r_sqlRet = r_sqlRet & sqltmp & ")"
            Else
                If oper = "!" Then oper = "<>"
                r_sqlRet = r_sqlRet & r_opSQL & "(" & nomChp & " " & oper & " '" & ValChp & "')"
            End If
        End If
        r_opSQL = " And "
        r_sql_PasF = r_sql_PasF & v_DemandenSQL & "§"
End Function
Private Function FaitLstService(v_numsrv, r_lstSrv)
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from service where srv_numpere=" & v_numsrv
    Call Odbc_SelectV(sql, rs)
    While Not rs.EOF
        r_lstSrv = r_lstSrv & ";" & rs("srv_num")
        Call FaitLstService(rs("srv_num"), r_lstSrv)
        rs.MoveNext
    Wend
    FaitLstService = r_lstSrv
End Function

Public Function P_Transforme(v_LaCond As String, v_LaValChp As String, v_LaFctValid As String, ByVal v_StrSQLBasic As String, ByRef r_StrSQLCond As String, ByRef r_StrCondF As String, ByRef r_BoolDetail As Boolean)
    Dim sql As String, rs As rdoResultset
    Dim nb As Integer
    Dim I As Integer
    Dim j As Integer, op As String, CondOut As String
    Dim UnItem As String
    Dim s As String
    Dim LaCondOut As String
    Dim BoolDetail As Boolean
    Dim LeOp As String
    Dim StrSQLBasic As String
    
    On Error GoTo Fin
    If v_LaFctValid = "%NUM" Then
        v_LaValChp = ""
    End If
    If v_LaValChp <> "" Then
        nb = STR_GetNbchamp(v_LaCond, ";")
        If v_LaCond = "R;" Or v_LaCond = "R" Then
            LaCondOut = " est renseigné"
        ElseIf v_LaCond = "NR;" Or v_LaCond = "NR" Then
            LaCondOut = " n'est pas renseigné"
        Else
            LaCondOut = v_LaCond
            LeOp = ""
            For I = 0 To nb - 1
                s = Trim(STR_GetChamp(v_LaCond, ";", I))
                sql = "select * from valchp where vc_num=" & s
                If Odbc_SelectV(sql, rs) <> P_ERREUR Then
                    LaCondOut = Replace(LaCondOut, s & ";", LeOp & rs("vc_lib"))
                    LeOp = " ou "
                End If
            Next I
        End If
        r_StrSQLCond = v_StrSQLBasic
        r_StrCondF = LaCondOut
        r_BoolDetail = BoolDetail
    ElseIf InStr(v_LaFctValid, "NUMSERVICE") > 0 Then
        s = Replace(v_LaCond, "!!", "")
        s = Replace(s, "DET", "")
        s = Replace(s, ";", "")
        s = Replace(s, "_", "")
        BoolDetail = IIf(InStr(v_LaCond, "DET") > 0, True, False)
        If v_LaCond = "N0_DET" Then ' Tout le site
            s = "R"
        End If
        If s = "R" Or s = "R;" Then
            r_StrCondF = " est renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf s = "NR" Or s = "NR;" Then
            r_StrCondF = " n'est pas renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf IsNumeric(s) Then
            If s = "-2" Then
                r_StrCondF = "Mon Service"
                r_StrSQLCond = v_StrSQLBasic
            Else
                StrSQLBasic = Replace(v_StrSQLBasic, "(", "")
                StrSQLBasic = Replace(StrSQLBasic, ")", "")
                Call P_TransformeItem("S", s, s, r_BoolDetail, op, StrSQLBasic, r_StrSQLCond, r_StrCondF, 0, r_BoolDetail)
                r_StrSQLCond = r_StrSQLCond
            End If
        Else
            MsgBox v_LaCond & " non traité"
        End If
        r_BoolDetail = BoolDetail
    ElseIf InStr(v_LaFctValid, "NUMFCT") > 0 Then
        If v_LaCond = "R;" Or v_LaCond = "R" Then
            r_StrCondF = " est renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf v_LaCond = "NR;" Or v_LaCond = "NR" Then
            r_StrCondF = " n'est pas renseigné"
            r_StrSQLCond = v_StrSQLBasic
        Else
            nb = STR_GetNbchamp(v_LaCond, ";")
            LaCondOut = ""
            LeOp = ""
            For I = 0 To nb - 1
                s = Trim(STR_GetChamp(v_LaCond, ";", I))
                s = Replace(s, "F", "")
                sql = "select * from fcttrav where ft_num=" & s
                If Odbc_SelectV(sql, rs) <> P_ERREUR Then
                    r_StrCondF = r_StrCondF & LeOp & rs("ft_libelle")
                    LeOp = " ou "
                End If
            Next I
            r_StrSQLCond = v_StrSQLBasic
        End If
        r_BoolDetail = BoolDetail
    ElseIf InStr(v_LaFctValid, "%DATE") > 0 Then
        If v_LaCond = "R;" Or v_LaCond = "R" Then
            r_StrCondF = " est renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf v_LaCond = "NR;" Or v_LaCond = "NR" Then
            r_StrCondF = " n'est pas renseigné"
            r_StrSQLCond = v_StrSQLBasic
        Else
            r_StrCondF = v_LaCond
            r_StrSQLCond = v_StrSQLBasic
        End If
        r_BoolDetail = BoolDetail
    ElseIf InStr(v_LaFctValid, "%ENTIER") > 0 Then
        If v_LaCond = "R;" Or v_LaCond = "R" Then
            r_StrCondF = " est renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf v_LaCond = "NR;" Or v_LaCond = "NR" Then
            r_StrCondF = " n'est pas renseigné"
            r_StrSQLCond = v_StrSQLBasic
        Else
            r_StrCondF = v_LaCond
            r_StrSQLCond = v_StrSQLBasic
        End If
        r_BoolDetail = BoolDetail
    ElseIf InStr(v_LaFctValid, "%NUM") > 0 Then
        If v_LaCond = "R;" Or v_LaCond = "R" Then
            r_StrCondF = " est renseigné"
            r_StrSQLCond = v_StrSQLBasic
        ElseIf v_LaCond = "NR;" Or v_LaCond = "NR" Then
            r_StrCondF = " n'est pas renseigné"
            r_StrSQLCond = v_StrSQLBasic
        Else
            r_StrCondF = v_LaCond
            r_StrSQLCond = v_StrSQLBasic
        End If
        r_BoolDetail = BoolDetail
    Else
        MsgBox "Case ?"
    End If
Fin:
    On Error GoTo 0
End Function

Public Function P_TransformeItem(ByVal v_Trait, ByVal v_Item, ByVal v_ItemBase, ByVal v_BoolDetail, ByVal v_Op, ByVal v_StrSQLBasic, ByRef r_StrSQLCond As String, ByRef r_StrCondF As String, ByVal v_niveau As Integer, ByRef r_BoolDetail As Boolean)
    Dim Cnd As String
    Dim sql As String, rs As rdoResultset
    Dim s As String
    
    Cnd = Replace(v_StrSQLBasic, ")", "")
    Cnd = Replace(Cnd, "(", "")
    ' Chercher le service
    sql = "select * from service where srv_num=" & v_Item
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        r_StrCondF = "Erreur SQL " & sql
    ElseIf rs.EOF Then
        If v_niveau = 0 Then
            r_StrCondF = "Service " & v_Item & " inconnu"
        End If
    Else
        If v_niveau = 0 Then r_StrCondF = rs("Srv_Nom")
        r_StrSQLCond = r_StrSQLCond & v_Op & Replace(v_StrSQLBasic, v_ItemBase, v_Item)
        v_Op = " Or "
        If v_BoolDetail Then
            r_BoolDetail = True
            If v_niveau = 0 Then r_StrCondF = r_StrCondF & " (D)"
            ' voir ses fils
            sql = "select * from service where srv_numpere=" & v_Item
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
            End If
            Do While Not rs.EOF
                s = P_TransformeItem(v_Trait, rs("srv_num"), v_ItemBase, v_BoolDetail, v_Op, v_StrSQLBasic, r_StrSQLCond, r_StrCondF, v_niveau + 1, r_BoolDetail)
                rs.MoveNext
            Loop
        End If
    End If
    rs.Close
End Function


Public Function RecupCndLstVal(ByVal TagVal As String)
    Dim sql As String, rs As rdoResultset
    Dim op As String, I As Integer
    Dim s As String
    Dim laSout As String
    
    op = ""
    laSout = ""
    If TagVal <> "" Then
        For I = 0 To STR_GetNbchamp(TagVal, ";")
            s = STR_GetChamp(TagVal, ";", I)
            If s <> "" Then
                If s = "<NR>" Then
                    laSout = laSout & op & "<Non renseignée>"
                    op = " ou "
                Else
                    sql = "select * from valchp where vc_num=" & Replace(s, "V", "")
                    If Odbc_SelectV(sql, rs) <> P_ERREUR Then
                        If Not rs.EOF Then
                            laSout = laSout & op & rs("vc_lib")
                            op = " ou "
                        End If
                    End If
                End If
            End If
        Next I
    End If
    RecupCndLstVal = laSout
End Function
Public Sub Public_VerifOuvrir(ByVal v_chemin As String, ByVal v_visible As Boolean, ByVal v_àSauver As Boolean, ByRef r_tbl_FichExcel() As FichExcelOuverts)
    
    Dim I As Integer
    Dim b_Fichier_Deja_Ouvert  As Boolean
    Dim strFichG As String, strFichd As String
    
    ' vérifier si une application Excel est déjà ouverte
    FctTrace ("Début Public_VerifOuvrir")
    FctTrace (".. 1 avant Excel_Init")
    Call Excel_Init
    FctTrace (".. 2 après Excel_Init")
    ' vérifier si le fichier modèle est ouvert : si non l'ouvrir
    On Error GoTo Err_Excel
Test_Excel:
    v_chemin = Replace(v_chemin, "..", ".")
    b_Fichier_Deja_Ouvert = False
    For I = 1 To Exc_obj.Workbooks.Count
        strFichG = Replace(UCase(Exc_obj.Workbooks(I).FullName), "\", "$")
        strFichG = Replace(strFichG, "/", "$")
        strFichd = Replace(UCase(v_chemin), "\", "$")
        strFichd = Replace(strFichd, "/", "$")
        If strFichG = strFichd Then
            Exc_obj.Workbooks(I).Activate
            If v_visible Then
                Call Excel_MetVisible
                Exc_obj.Visible = v_visible
            Else
                Exc_obj.Visible = False
            End If
            Exit Sub
        End If
    Next I
    FctTrace (".. 3 avant Public_OuvrirModele")
    Public_OuvrirModele v_chemin, v_visible, v_àSauver, r_tbl_FichExcel()
    FctTrace (".. 4 après Public_OuvrirModele")
    FctTrace ("Fin Public_VerifOuvrir")
    Exit Sub
    
Err_Excel:
    'MsgBox Err & " " & Error$
    If Err = 462 Or Err = 91 Then
        If Excel_Init() = P_OK Then
        End If
    End If
    Resume Next
End Sub

Public Function Public_OuvrirModele(ByVal v_chemin As String, ByVal v_visible As Boolean, ByVal v_àSauver As Boolean, ByRef r_tbl_FichExcel() As FichExcelOuverts) As Integer
    
    Dim encore As Boolean
    Dim retour As Integer
    Dim FichierIn As String, cmd As String
    Dim v_chemin_For As String, v_chemin_Fil As String, v_chemin_Excel As String
    
    ' Ouvrir le modele
    FctTrace ("Début Public_OuvrirModele " & v_chemin)
    If FICH_FichierExiste(v_chemin) Then
        FctTrace (".. 1 avant Excel_OuvrirDoc")
        Excel_OuvrirDoc v_chemin, "", Exc_wrk, False
        FctTrace (".. 2 après Excel_OuvrirDoc")
        FctTrace (".. 3 avant Public_FichiersExcelOuverts")
        Call Public_FichiersExcelOuverts(r_tbl_FichExcel(), "VoirExcel", v_chemin, v_visible, v_àSauver)
        FctTrace (".. 4 après Public_FichiersExcelOuverts")
        If v_visible Then
            Call Excel_MetVisible
        Else
            Exc_obj.Visible = False
        End If
        FctTrace ("Fin Public_OuvrirModele " & v_chemin)
        Public_OuvrirModele = P_OK
    Else
        Call MsgBox("Impossible d'ouvrir le fichier " & v_chemin, vbCritical + vbOKOnly, "")
        FctTrace ("Fin Public_OuvrirModele Erreur " & v_chemin)
        Public_OuvrirModele = P_ERREUR
    End If
    
End Function

Public Sub P_SimulMettreChamp(ByVal v_dansGrid As Boolean, ByVal v_dansExcel As Boolean, v_I_TabExcel As Integer, ByRef v_lex As Integer, ByRef v_leY As Integer, v_MenForme As String, v_libelle As String, v_valeur As String, v_idgrid As Integer, v_bool_liste As Boolean, v_SQL As Variant, v_numchp As Integer, v_Epingle As Boolean)
    Dim NomCellDest As String
    Dim V_url As String
    Dim url As String
    Dim util As String
    Dim cnd_sversconf As String
    Dim exc_sheet As Excel.Worksheet
    Dim ANC_bfaire_RowColChange As Boolean
    Dim LaColMax As Integer, LaRowMax As Integer
    Dim s As String
    Dim LesLinks As Hyperlinks
    Dim Padd As String
    Dim Padd2 As String
    Dim ij As Integer
    Dim NumForm As String
    Dim NumFiltre As String
    Dim sSQL As String
    Dim MenFType As String
    Dim bHyperL As Boolean
    Dim iFeuille As Integer
    
    'Debug.Print v_SQL
    On Error GoTo 0
    
    ANC_bfaire_RowColChange = p_bfaire_RowColChange
    'Call FctTrace("Mémoire Totale=" & Exc_obj.MemoryTotal & " Free=" & Exc_obj.MemoryFree & " Used=" & Exc_obj.MemoryUsed)
    p_bfaire_RowColChange = False
    If v_Epingle Then
        Padd = "    "
    Else
        Padd = "   "
    End If
    'If v_lex = 0 Or v_leY = 0 Then
    '    MsgBox "ici"
    'End If
    MenFType = STR_GetChamp(tbl_fichExcel(v_I_TabExcel).CmdMenFormeChp, "#", 1)
    
    If v_dansExcel Then
        If (v_MenForme = "NewFenêtre") Then
            iFeuille = p_LeIndexFenetreExcel
        Else
            iFeuille = v_idgrid
        End If
        Set exc_sheet = Exc_wrk.Sheets(iFeuille)
    End If
    
    If tbl_fenExcel(v_idgrid).FenLoad Then
        LaColMax = PiloteExcelBis.grdCell(v_idgrid).Cols - 1
        LaRowMax = PiloteExcelBis.grdCell(v_idgrid).Rows
    End If
    
    If tbl_fichExcel(v_I_TabExcel).CmdType = "RES" Then
        NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
        If v_dansExcel Then
            exc_sheet.Range(NomCellDest).Value = v_valeur
            Exit Sub
        End If
    End If
    If v_MenForme = "Simple" Then
        Padd = ""
        NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
        If v_dansExcel Then
            Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_libelle, v_SQL, MenFType)
            Exit Sub
        End If
        GoTo Lab_Colonne_Lib
    End If

    If (v_MenForme = "Colonne_Lib" Or v_MenForme = "Colonne_Val" Or v_MenForme = "Colonne_Lib_Val") Then
        If (v_MenForme = "Colonne_Lib") Then  'Colonne_Lib
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Colonne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
Lab_Colonne_Lib:
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_libelle, "", MenFType)
            End If
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & v_libelle
            End If
            If p_TraitPublier = "G" Then
                If Publier.ChkHyperlien.Value = 1 Then
                    bHyperL = True
                End If
            ElseIf p_TraitPublier = "P" Then
                If PiloteExcelBis.ChkHyperlien.Value = 1 Then
                    bHyperL = True
                End If
            End If
            If v_MenForme = "Simple" Then
                p_bfaire_RowColChange = ANC_bfaire_RowColChange
                If bHyperL And v_dansExcel Then
                    'Debug.Print v_SQL
                    P_MetLink v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
                End If
                Exit Sub
            End If
        ElseIf (v_MenForme = "Colonne_Val") Then  'Colonne_Val
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Colonne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_valeur, v_SQL, MenFType)
            End If
            ' mettre seulement la valeur
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                If v_valeur <> "" Then
                    PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & IIf(MenFType = "POURCENT", Replace(v_valeur, ".", ",") & "%", Replace(v_valeur, ".", ","))
                End If
                PiloteExcelBis.grdCell(v_idgrid).col = v_lex
                PiloteExcelBis.grdCell(v_idgrid).row = v_leY
                Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            End If
        ElseIf (v_MenForme = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Colonne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_libelle, "", MenFType)
            End If
            ' mettre seulement la valeur
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & v_libelle
            End If
            NomCellDest = FctFaitNomCellDest((v_lex + 1), v_leY)
            If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex + 1 = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, (v_lex + 1), v_leY, NomCellDest, v_valeur, v_SQL, MenFType)
            End If
            ' mettre seulement la valeur
            If v_dansGrid And v_lex + 1 <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex + 1) = Padd2 & IIf(MenFType = "POURCENT", Replace(v_valeur, ".", ",") & "%", Replace(v_valeur, ".", ","))
                PiloteExcelBis.grdCell(v_idgrid).col = v_lex + 1
                PiloteExcelBis.grdCell(v_idgrid).row = v_leY
                Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            End If
        End If
        
        ' mettre le commentaire
        bHyperL = (p_ModePublication = "Param" And PiloteExcelBis.ChkHyperlien.Value = 1) Or (p_ModePublication = "Publier" And Publier.ChkHyperlien.Value)
        If bHyperL And v_dansExcel Then
            If (v_MenForme = "Colonne_Lib_Val") Then  'sur la colonne des valeurs
                P_MetLink v_idgrid, v_lex + 1, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
            Else
                P_MetLink v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
            End If
        End If
        v_leY = v_leY + 1
    End If
    
    If v_dansExcel And (v_MenForme = "MêmeFenêtre" Or v_MenForme = "NewFenêtre") Then
        Dim TypVal As String
        NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
        If v_SQL <> "DON_NUM" Then TypVal = P_RecupereNomChamp(v_numchp, "type")
        If p_LeTypeTitreOuChamp = "T" Then
            exc_sheet.Range(NomCellDest).Value = v_valeur
            exc_sheet.Range(NomCellDest).WrapText = True
        Else
            If v_SQL = "DON_NUM" Then
                exc_sheet.Cells(v_leY, v_lex).Value = STR_GetChamp(v_valeur, "|", 0)
                If exc_sheet.Cells(v_leY, v_lex).Value > 0 Then
                    Set LesLinks = exc_sheet.Hyperlinks
                    If p_estV4 Then
                        V_url = "V4/kaliform/form_saisie.php%3FV_numfor=" & STR_GetChamp(v_valeur, "|", 1) & "%26V_numdon=" & STR_GetChamp(v_valeur, "|", 0)
                    Else
                        V_url = "form_saisie.php%3FV_numfor=" & STR_GetChamp(v_valeur, "|", 1) & "%26V_numdon=" & STR_GetChamp(v_valeur, "|", 0)
                    End If
                    ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
                    util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
                    
                    If p_S_Vers_Conf <> "" Then
                        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
                    End If
                    url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & V_url
                    LesLinks.Add Anchor:=exc_sheet.Cells(v_leY, v_lex), Address:=url
                End If
            Else
                v_valeur = Replace(v_valeur, """", "'")
                If left$(v_valeur, 1) = "=" Then
                    v_valeur = Mid$(v_valeur, 2)
                End If
                exc_sheet.Range(NomCellDest).Value = v_valeur
                exc_sheet.Range(NomCellDest).WrapText = True
            End If
        End If
        If v_MenForme = "MêmeFenêtre" And p_LeTypeTitreOuChamp = "T" Then
            exc_sheet.Range(NomCellDest).Font.Color = 255   ' rouge
            exc_sheet.Range(NomCellDest).Font.Bold = True
        End If
        ' link pour retour
        If v_MenForme = "NewFenêtre" And p_LeTypeTitreOuChamp = "T" Then
            s = "'" & tbl_fenExcel(p_LeIndexFeuille_PourHyperlien).FenNom & "'!" & FctStrColDest(p_LeX_PourHyperlienG) & p_LeY_PourHyperlien
            If Exc_wrk.Sheets(p_LeIndexFenetreExcel).Cells(v_leY, v_lex).Value > 0 Then
                Set LesLinks = Exc_wrk.Sheets(p_LeIndexFenetreExcel).Hyperlinks
                'LesLinks.Add Anchor:=Exc_wrk.Sheets(p_LeIndexFenetreExcel).Cells(v_leY, v_leX), Address:="", SubAddress:="'Feuille 1'!A1"
                LesLinks.Add Anchor:=Exc_wrk.Sheets(p_LeIndexFenetreExcel).Cells(v_leY, v_lex), Address:="", SubAddress:=s
            End If
        End If
        
        If v_lex <= p_LeXMaxPourGrdCell And v_leY <= p_LeYMaxPourGrdCell Then
            If p_bool_tbl_cell And p_MettreCommentListeChamp Then
                On Error GoTo Suite_CommentListe
                exc_sheet.Range(NomCellDest).Select
                exc_sheet.Range(NomCellDest).AddComment
                exc_sheet.Range(NomCellDest).Comment.Text Text:="Sql=" & Chr(10) & v_SQL
Suite_CommentListe:
            End If
        End If
        v_lex = v_lex + 1
    End If

    If (v_MenForme = "Ligne_Lib" Or v_MenForme = "Ligne_Val" Or v_MenForme = "Ligne_Lib_Val") Then 'en Ligne
        If (v_MenForme = "Ligne_Lib") Then  'Ligne_Lib
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Ligne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_libelle, "", MenFType)
            End If
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & v_libelle
            End If
        ElseIf (v_MenForme = "Ligne_Val") Or (v_MenForme = "MêmeFenêtre") Then  'Ligne_Val
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Colonne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_valeur, v_SQL, MenFType)
            End If
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                If MenFType = "POURCENT" Then
                    PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & Replace(v_valeur, ".", ",") & "%"
                Else
                    PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & Replace(v_valeur, ".", ",")
                End If
                PiloteExcelBis.grdCell(v_idgrid).col = v_lex
                PiloteExcelBis.grdCell(v_idgrid).row = v_leY
                Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            End If
        ElseIf (v_MenForme = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
            NomCellDest = FctFaitNomCellDest(v_lex, v_leY)
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansGrid Then Call P_MettrePicture("Colonne", v_I_TabExcel, v_idgrid, v_leY, v_lex, v_Epingle)
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, v_leY, NomCellDest, v_libelle, "", MenFType)
            End If
            ' mettre seulement la valeur
            If v_dansGrid And v_lex <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY, v_lex) = Padd & v_libelle
            End If
            NomCellDest = FctFaitNomCellDest(v_lex, (v_leY + 1))
            If v_dansGrid Then
                If v_leY = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_lex + 1 = PiloteExcelBis.grdCell(v_idgrid).Cols Then v_dansGrid = False
            End If
            If v_dansExcel Then
                Call ecrire_liens(v_I_TabExcel, iFeuille, v_lex, (v_leY + 1), NomCellDest, v_valeur, v_SQL, MenFType)
            End If
            If v_dansGrid And v_lex + 1 <= LaColMax And v_leY <= LaRowMax Then
                PiloteExcelBis.grdCell(v_idgrid).TextMatrix(v_leY + 1, v_lex) = Padd2 & IIf(MenFType = "POURCENT", Replace(v_valeur, ".", ",") & "%", Replace(v_valeur, ".", ","))
                PiloteExcelBis.grdCell(v_idgrid).col = v_lex
                PiloteExcelBis.grdCell(v_idgrid).row = v_leY + 1
                Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            End If
        End If
        ' mettre le commentaire
Mettre_Le_Commentaire:
        If p_dansGrid Then
            If p_bool_tbl_cell And p_BoolMettreComment Then
                On Error GoTo Suite_CommentL
                If v_dansExcel Then
                    exc_sheet.Range(NomCellDest).Select
                    exc_sheet.Range(NomCellDest).AddComment
                    exc_sheet.Range(NomCellDest).Comment.Text Text:="Sql=" & Chr(10) & v_SQL
Suite_CommentL:
                    'P_MetSQL v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
                End If
            Else
                'P_MetSQL v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
            End If
        End If
        bHyperL = (p_ModePublication = "Param" And PiloteExcelBis.ChkHyperlien.Value = 1) Or (p_ModePublication = "Publier" And Publier.ChkHyperlien.Value)
        If bHyperL Then
            'Call ecrire_liens(iFeuille, v_lex, v_leY, v_valeur, v_SQL, MenFType)
            'If v_dansExcel Then
                'If (v_MenForme = "Ligne_Lib_Val") Then  'sur la ligne des valeurs
                '    P_MetLink v_idgrid, v_lex, v_leY + 1, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
                'Else
                '    P_MetLink v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
                'End If
            'Else
            '    P_MetLink v_idgrid, v_lex, v_leY, v_SQL, tbl_fichExcel(v_I_TabExcel).CmdFormNum
            'End If
        End If
        v_lex = v_lex + 1
    End If
    p_bfaire_RowColChange = ANC_bfaire_RowColChange
End Sub

Public Function ecrire_liens(i_tabExcel, v_iFeuille, v_X, v_Y, v_Cell, v_valeur, v_SQL, MenFType)
    Dim fdliens As Integer
    Dim s As String
    ' i_tabExcel indique le champ d'origine de la donnée
    
    'Debug.Print v_SQL
    If v_Cell <> "" Then
        'Debug.Print i_tabExcel & " " & v_iFeuille & " " & v_X & " " & v_Y & " " & v_Cell & " " & v_SQL
        fdliens = FreeFile
        FICH_OuvrirFichier p_chemin_fichier_liens, FICH_ECRITURE, fdliens
        s = ""
        If v_SQL <> "" Then
            s = P_FaitLink(v_iFeuille, v_Cell, v_SQL)
        End If
        s = v_iFeuille & "|" & v_X & "|" & v_Y & "|" & v_Cell & "|" & v_valeur & "|" & MenFType & "|" & s & "|" & i_tabExcel & "|" & tbl_fichExcel(i_tabExcel).CmdFormNum & "|" & v_SQL & "|" & p_S_Vers_Conf
        'Debug.Print s
        Print #fdliens, s
        Close #fdliens
    End If
End Function

Public Function FctIntColdest(ByVal v_col As String)
    Dim I As Integer, s As String, sret As String
    Dim n As Integer
    Dim iret As Integer
    Dim pos As Integer
    Dim dec As Integer
    
    n = Len(v_col)
    If n = 1 Then
        iret = InStr(Public_Alpha, Mid$(v_col, 1, 1))
    Else
        I = InStr(Public_Alpha, Mid$(v_col, 1, 1))
        iret = I * 26
        I = InStr(Public_Alpha, Mid$(v_col, 2, 1))
        iret = iret + I
    End If
    FctIntColdest = iret
End Function

Public Function FctStrColDest(ByVal v_lex As Integer)
    Dim sret As String
    Dim I As Integer, x As Integer
    Dim j As Integer
    Dim encore As Boolean
    
    sret = ""
    If v_lex <= 26 Then
        sret = Mid$(Public_Alpha, v_lex, 1)
    Else
        encore = True
        While encore
            I = Int(v_lex / 26)
            If v_lex = 26 Then
                sret = "Z"
                v_lex = 0
            ElseIf I > 0 Then
                sret = sret & Mid$(Public_Alpha, I, 1)
                v_lex = v_lex - (26 * I)
            Else
                x = v_lex Mod 26
                sret = sret & Mid$(Public_Alpha, x, 1)
                v_lex = 0
            End If
            If v_lex <= 0 Then
                encore = False
            End If
        Wend
    End If
    FctStrColDest = sret
End Function

Public Function FctFaitNomCellDest(ByVal v_lex As Integer, v_leY)
    Dim sret As String
    Dim s As String
    Dim x As Integer
    Dim I As Integer
    Dim j As Integer
    Dim encore As Boolean
    
    If v_lex = 0 Or v_leY = 0 Then
        FctFaitNomCellDest = ""
        Exit Function
    End If
    sret = ""
    encore = True
    If v_lex <= 26 Then
        FctFaitNomCellDest = Mid$(Public_Alpha, v_lex, 1) & v_leY
        Exit Function
    Else
        While encore
            I = Int(v_lex / 26)
            If v_lex = 26 Then
                sret = "Z"
                v_lex = 0
            ElseIf I > 0 Then
                sret = sret & Mid$(Public_Alpha, I, 1)
                v_lex = v_lex - (26 * I)
            Else
                x = v_lex Mod 26
                sret = sret & Mid$(Public_Alpha, x, 1)
                v_lex = 0
            End If
            If v_lex <= 0 Then
                encore = False
            End If
        Wend
    End If
    sret = sret & v_leY
    
    'If (v_lex <= 26) Then
    '    sret = (Mid$(Public_Alpha, v_lex, 1) & v_leY)
    'Else
        'i = Int(v_lex / 26)
        'sret = Mid$(Public_Alpha, i, 1)
        'x = v_lex Mod 26
        'sret = sret & Mid$(Public_Alpha, x, 1)
        'encore = True
        'i = 1
        'x = 1
        'j = 0
        'Do While encore
        '    s = Mid$(Public_Alpha, i, 1)
        '    If i = 26 Then
        '        j = j + 1
        '        i = 0
        '    End If
        '    If x >= v_leX Then encore = False
        '    i = i + 1
        '    x = x + 1
        'Loop
        'sret = Mid$(Public_Alpha, j, 1) & s & v_leY
    'End If
    FctFaitNomCellDest = sret
End Function

Public Function FctEvalue_VersionExcel()

End Function

Public Function P_MettreLien(v_lex, v_leY, v_I_TabExcel, v_SQL)
    Dim ij As Integer
    Dim sSQL As String
    Dim NumFiltre As String
    Dim laSQL As String
    Dim NumForm As String
    Dim url As String
    Dim sURL As String
    Dim util As String, cnd_sversconf As String
    
    For ij = 0 To UBound(tbl_cell())
        If tbl_cell(ij).CellX = v_lex And tbl_cell(ij).CellY = v_leY Then
            If tbl_cell(ij).CellFeuille = tbl_fichExcel(v_I_TabExcel).CmdFenNum Then
                NumFiltre = tbl_fichExcel(v_I_TabExcel).CmdFormNum
                tbl_cell(ij).cellNumFiltre = NumFiltre
                laSQL = "select for_num from formulaire,filtreform where formulaire.for_num = filtreform.ff_fornum " & " and filtreform.ff_num = " & NumFiltre
                Call Odbc_RecupVal(laSQL, NumForm)
                sSQL = v_SQL
                If InStr(UCase(sSQL), "WHERE") = 0 Then sSQL = sSQL & " Where true"
                sSQL = Mid(sSQL, InStr(UCase(sSQL), "WHERE"))
                sSQL = Replace(UCase(sSQL), "WHERE", "")
                If p_estV4 Then
                    url = "V4/kaliform/filtre_form.php%3FV_numfiltre=" & NumFiltre & "%26V_numfor=" & NumForm & "%26V_etat=2%26V_etattermine=0" & "%26V_typaff=D%26V_quitter=1" & "%26V_RapportType=" & sSQL
                    util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
                    If p_S_Vers_Conf <> "" Then
                        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
                    End If
'                    cnd_sversconf = "&s_vers_conf=_STB"
'                    p_AdrServeur = "192.168.101.24"
'                    util = STR_CrypterNombre(Format(702, "#0000000"))
                    sURL = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & url
                    Exit Function
                Else
                    url = "filtres/liste_form_resp.php%3FV_numfiltre=" & NumFiltre & "%26V_numfor=" & NumForm & "%26V_etat=2%26V_etattermine=0" & "%26V_typaff=D%26V_quitter=1" & "%26V_RapportType=" & sSQL
                End If
                util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
                If p_S_Vers_Conf <> "" Then
                    cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
                End If
                sURL = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & url
                tbl_cell(ij).CellLink = sURL
                Exit For
            End If
        End If
    Next ij

End Function


Public Function P_MettrePicture(v_Trait, v_tbl, v_idgrid, v_row, v_col, v_Epingle)
    Dim Padd1 As String, Padd2 As String
    Dim Padd As String

    Padd1 = "    "
    Padd2 = "   "
    
    If p_ChpType = "ListeChamp" Then
        MsgBox "ici P_MettrePicture"
    End If
    
    If v_row = PiloteExcelBis.grdCell(v_idgrid).Rows Or v_col = PiloteExcelBis.grdCell(v_idgrid).Cols Then
        GoTo LabPasMettre
    End If
    On Error GoTo LabErreur
    PiloteExcelBis.grdCell(v_idgrid).col = v_col
    PiloteExcelBis.grdCell(v_idgrid).row = v_row
    If v_Trait = "Colonne" Then
        If v_Epingle Then
            Padd = Padd1
            If tbl_fichExcel(v_tbl).CmdFormNum = p_numfiltre_encours And tbl_fichExcel(v_tbl).CmdFormIndice = p_numindice_encours Then
                If tbl_fichExcel(v_tbl).CmdChpSQL <> "" Then
                    If tbl_fichExcel(v_tbl).CmdCondition <> "" Then
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_SQL_SELECT_FPLUS).Picture
                    Else
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_SQL_SELECT_F).Picture
                    End If
                ElseIf tbl_fichExcel(v_tbl).CmdCondition <> "" Then
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBF_PLUS).Picture
                Else
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBF).Picture
                End If
            Else
                If tbl_fichExcel(v_tbl).CmdChpSQL <> "" Then
                    If tbl_fichExcel(v_tbl).CmdCondition <> "" Then
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_SQL_SELECT_CPLUS).Picture
                    Else
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_SQL_SELECT_C).Picture
                    End If
                ElseIf tbl_fichExcel(v_tbl).CmdCondition <> "" Then
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBC_PLUS).Picture
                Else
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBC).Picture
                End If
            End If
        Else
            Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            Padd = Padd2
        End If
        If (p_MenFLigCol = "Colonne_Lib_Val") Then  'Colonne_Lib_Val
            PiloteExcelBis.grdCell(v_idgrid).col = v_col + 1
            PiloteExcelBis.grdCell(v_idgrid).row = v_row
            Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
        End If


    ElseIf v_Trait = "Ligne" Then

        PiloteExcelBis.grdCell(v_idgrid).col = v_col
        PiloteExcelBis.grdCell(v_idgrid).row = v_row
        If v_Epingle Then
            Padd = Padd1
            If tbl_fichExcel(v_tbl).CmdFormNum = p_numfiltre_encours And tbl_fichExcel(v_tbl).CmdFormIndice = p_numindice_encours Then
                If p_ChpType = "ListeChamps" Then
                    If InStr(tbl_fichExcel(v_tbl).cmdTypeChp, "Ici") > 0 Then
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_CHP_LOUPER).Picture
                    Else
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_CHP_LOUPER).Picture
                    End If
                Else
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBF).Picture
                End If
            Else
                If p_ChpType = "ListeChamps" Then
                    If InStr(tbl_fichExcel(v_tbl).cmdTypeChp, "Ici") > 0 Then
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_CHP_LOUPEB).Picture
                    Else
                        Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_CHP_LOUPEB).Picture
                    End If
                Else
                    Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULEBC).Picture
                End If
            End If
        Else
            Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
            Padd = Padd2
        End If
        If (p_MenFLigCol = "Ligne_Lib_Val") Then  'Ligne_Lib_Val
            PiloteExcelBis.grdCell(v_idgrid).col = v_col
            PiloteExcelBis.grdCell(v_idgrid).row = v_col + 1
            Set PiloteExcelBis.grdCell(v_idgrid).CellPicture = PiloteExcelBis.imglst.ListImages(IMG_BOULE).Picture
        End If
    Else
        MsgBox "Case ?"
    End If
LabErreur:
    On Error GoTo 0
LabPasMettre:
End Function

Public Function P_ListeValeurs(ByVal v_valeur, ByVal V_ChpNum)
    Dim sql As String, rs As rdoResultset
    Dim ChpType As String
    Dim ValChp As String
    Dim TitreChp As String
    Dim fctvalid As String
    Dim op As String
    Dim I As Integer
    Dim lib As String
    
    sql = "select forec_type, forec_valeurs_possibles, forec_fctvalid from formetapechp where forec_num=" & V_ChpNum
    If Odbc_RecupVal(sql, ChpType, ValChp, fctvalid) = P_ERREUR Then
        P_ListeValeurs = "??? " & sql
        Exit Function
    End If
    If ChpType = "CHECK" Or ChpType = "SELECT" Or ChpType = "RADIO" Then
        sql = "select * from valchp where vc_lvcnum=" & ValChp
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            P_ListeValeurs = "??? " & sql
            Exit Function
        End If
        
        ValChp = ""
        For I = 0 To STR_GetNbchamp(v_valeur, ";")
            If STR_GetChamp(v_valeur, ";", I) <> "" Then
                sql = "select VC_Lib from valchp where vc_num=" & Replace(STR_GetChamp(v_valeur, ";", I), "V", "")
                If Odbc_RecupVal(sql, lib) = P_ERREUR Then
                    P_ListeValeurs = "??? " & sql
                    Exit Function
                Else
                    ValChp = ValChp & op & lib
                    op = " - "
                End If
            End If
        Next I
        P_ListeValeurs = ValChp
    ElseIf ChpType = "HIERARCHIE" Then
        sql = "select * from hierarvalchp where hvc_lvcnum=" & ValChp
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            P_ListeValeurs = "??? " & sql
            Exit Function
        End If
        
        ValChp = ""
        For I = 0 To STR_GetNbchamp(v_valeur, ";")
            If STR_GetChamp(v_valeur, ";", I) <> "" Then
                sql = "select HVC_Nom from hierarvalchp where hvc_num=" & Replace(STR_GetChamp(v_valeur, ";", I), "V", "")
                If Odbc_RecupVal(sql, lib) = P_ERREUR Then
                    P_ListeValeurs = "??? " & sql
                    Exit Function
                Else
                    ValChp = ValChp & op & lib
                    op = " - "
                End If
            End If
        Next I
        P_ListeValeurs = ValChp
    Else
        MsgBox "ici P_ListeValeurs"
        P_ListeValeurs = "??? " & ChpType
    End If
    
End Function

Public Function P_RecupereNomChamp(ByVal V_ChpNum As Integer, v_Trait As String)
    Dim sql As String, ChpNom As String
    Dim rs As rdoResultset
    Dim ChpType As String
    Dim ValChp As String
    Dim TitreChp As String
    Dim Label As String
    Dim formule As String
    Dim fctvalid As String
    
    If Odbc_RecupVal("select forec_formule, forec_label,forec_nom,forec_type,forec_valeurs_possibles, forec_fctvalid from formetapechp where forec_num=" & V_ChpNum, _
                     formule, TitreChp, ChpNom, ChpType, ValChp, fctvalid) = P_ERREUR Then
        ChpNom = "???"
        Exit Function
    End If
    If v_Trait = "label" Then
        P_RecupereNomChamp = TitreChp
    ElseIf v_Trait = "nom" Then
        P_RecupereNomChamp = ChpNom
    ElseIf v_Trait = "type" Then
        P_RecupereNomChamp = ChpType
    ElseIf v_Trait = "valchp" Then
        P_RecupereNomChamp = ValChp
    ElseIf v_Trait = "formule" Then
        P_RecupereNomChamp = formule
    ElseIf v_Trait = "fctvalid" Then
        P_RecupereNomChamp = fctvalid
    End If
End Function


Public Sub P_MetSQL(v_numfeuille, v_X, v_Y, v_SQL, v_NumFiltre)
    Dim ij As Integer
    
    If p_bool_tbl_cell Then
        For ij = 0 To UBound(tbl_cell())
            If tbl_cell(ij).CellFeuille = v_numfeuille Then
                If tbl_cell(ij).CellX = v_X And tbl_cell(ij).CellY = v_Y Then
                    tbl_cell(ij).cellSQL = v_SQL
                    tbl_cell(ij).cellNumFiltre = v_NumFiltre
                    Exit For
                End If
            End If
        Next ij
    End If
End Sub



Public Sub P_MetLink(v_numfeuille, v_X, v_Y, v_SQL, v_NumFiltre)
    Dim ij As Integer
    Dim G As String
    Dim laSQL As String, rs As rdoResultset
    Dim pos As Integer
    Dim Cnd As String
    Dim strX As String
    Dim numfor As String
    Dim V_url As String
    Dim V_urlv As Variant
    Dim util As String
    Dim cnd_sversconf As String
    Dim url As String
    Dim links As Hyperlinks
    Dim valeur As String
    Dim fdliens As Integer
    Dim Gv As Variant
    Dim URLv As Variant
    Dim Cndv As Variant
    
    If p_TraitPublier = "P" Then
        p_chemin_fichier_liens = p_Chemin_Modeles_Local & "\Temp_" & p_nummodele & "_" & p_NumUtil & ".txt"
    End If

    If p_chemin_fichier_liens = "" Then Exit Sub
    If v_SQL = "" Then Exit Sub
    
    'Debug.Print v_SQL
    Gv = STR_GetChamp(v_SQL, "|", 0)
    pos = InStr(Gv, "Where")
    If pos = 0 Then
        Cndv = Gv
    Else
        Cndv = Mid(Gv, pos + 5)
    End If
    laSQL = "select * from formulaire,filtreform where formulaire.for_num = filtreform.ff_fornum " _
        & " and filtreform.ff_num = " & v_NumFiltre
    If Odbc_SelectV(laSQL, rs) = P_ERREUR Then
        Exit Sub
    End If
    numfor = rs("for_num").Value
    rs.Close
    If p_estV4 Then
        V_urlv = "V4/kaliform/filtre_form.php%3FV_numfiltre=" & v_NumFiltre & "%26V_RapportType=" & Cndv
        util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
        cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
        url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & V_urlv
        util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
    Else
        V_urlv = "filtres/liste_form_resp.php%3FV_numfiltre=" & v_NumFiltre & "%26V_numfor=" & numfor & "%26V_etat=2%26V_etattermine=0" & "%26V_typaff=D%26V_quitter=1" & "%26V_RapportType=" & Cndv
        ' Permet d’ouvrir IE en grand avec l’URL indiqué dans la variable ‘url’
        util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
        If p_S_Vers_Conf <> "" Then
            cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
        End If
        URLv = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util & "&V_url=" & V_urlv
    End If
    
    strX = FctStrColDest(v_X)
    
    Dim Faire As Boolean
    If p_dansExcel And Exc_wrk.Sheets(v_numfeuille).Cells(v_Y, v_X).Value > 0 Then
        Faire = True
    ElseIf p_dansGrid Then
        Faire = True
    End If
    If Faire Then
        Set links = Exc_wrk.Sheets(v_numfeuille).Hyperlinks
        On Error GoTo NextErreur
        If p_estV4 Then
            url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util
            url = url & "&V_url=RapportTypeLiens.php%3FstrX=" & strX & "%26strY=" & v_Y & "%26numfeuille=" & v_numfeuille
            url = url & "%26numFichierLiens=" & p_numFichier_Liens & "%26numDocument=" & p_numdoc_encours & "%26numModele=" & p_nummodele_encours
        Else
            url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util
            url = url & "&V_url=RapportTypeLiens.php%3FstrX=" & strX & "%26strY=" & v_Y & "%26numfeuille=" & v_numfeuille
            url = url & "%26numFichierLiens=" & p_numFichier_Liens & "%26numDocument=" & p_numdoc_encours & "%26numModele=" & p_nummodele_encours
        End If
        url = url & "%26V_mode=" & p_TraitPublier
        ' si p_TraitPublier = "P"     ' les liens sont dans /Temp
        If p_TraitPublier = "P" Then
            url = url & "%26FichierLiens=" & "Temp_" & p_nummodele & "/Temp_" & p_NumUtil & ".txt"
        End If
        If p_bool_tbl_cell Then
            For ij = 0 To UBound(tbl_cell())
                If tbl_cell(ij).CellFeuille = v_numfeuille Then
                    If tbl_cell(ij).CellY = PiloteExcelBis.grdCell(v_numfeuille).RowSel Then
                        If tbl_cell(ij).CellX = PiloteExcelBis.grdCell(v_numfeuille).ColSel Then
                            tbl_cell(ij).CellLink = url
                            tbl_cell(ij).cellSQL = v_SQL
                            tbl_cell(ij).cellNumFiltre = v_NumFiltre
                            Exit For
                        End If
                    End If
                End If
            Next ij
        End If
        ' Ecrire dans le fichier du resultat
        fdliens = FreeFile
        FICH_OuvrirFichier p_chemin_fichier_liens, FICH_ECRITURE, fdliens
        V_urlv = v_X & "|" & v_Y & "|" & v_numfeuille & "|" & V_urlv
        
        'Debug.Print V_urlv
        Print #fdliens, V_urlv
        Close #fdliens
        
        On Error GoTo 0
        GoTo Pas_Erreur
NextErreur:
        MsgBox "Erreur en création des liens" & Chr(13) & Chr(10) & "   => " & Error$ & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Les liens ont été déscativés"
        Publier.ChkHyperlien.Value = False      ' Désactiver la mise en place des liens
Pas_Erreur:
    End If
End Sub

Public Function P_FaitLink(v_numfeuille, v_Cell, v_SQL)
    Dim Cndv As Variant
    Dim pos As Integer
    Dim strX As String, cnd_sversconf As String, util As String, url As String
    
    If p_chemin_fichier_liens = "" Then Exit Function
    'pos = InStr(v_SQL, "Where")
    'If pos = 0 Then
    '    Cndv = v_SQL
    'Else
    '    Cndv = Mid(v_SQL, pos + 5)
    'End If
    
    Dim Faire As Boolean
    Dim s As String
    
    If p_dansExcel Then
        Faire = True
    ElseIf p_dansGrid Then
        Faire = True
    End If
    
    If Faire Then
        If p_S_Vers_Conf <> "" Then
            cnd_sversconf = "&s_vers_conf=" & p_S_Vers_Conf
        End If
        util = STR_CrypterNombre(Format(p_NumUtil, "#0000000"))
        If p_estV4 Then
            url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util
            url = url & "&V_url=RapportTypeLiens.php%3FCell=" & v_Cell & "%26numfeuille=" & v_numfeuille
            url = url & "%26numFichierLiens=" & p_numFichier_Liens & "%26numDocument=" & p_numdoc_encours & "%26numModele=" & p_nummodele_encours
        Else
            url = "http://" & p_AdrServeur & "/publiweb/pident.php?in=divers" & cnd_sversconf & "&V_util=" & util
            url = url & "&V_url=RapportTypeLiens.php%3FCell=" & v_Cell & "%26numfeuille=" & v_numfeuille
            url = url & "%26numFichierLiens=" & p_numFichier_Liens & "%26numDocument=" & p_numdoc_encours & "%26numModele=" & p_nummodele_encours
        End If
        url = url & "%26V_mode=" & p_TraitPublier
        ' si p_TraitPublier = "P"     ' les liens sont dans /Temp
        If p_TraitPublier = "P" Then
            url = url + "%26random=" & p_numRandom
        End If
        If p_S_Vers_Conf <> "" Then
            cnd_sversconf = "%26s_vers_conf=" & p_S_Vers_Conf
        End If
        url = url & cnd_sversconf
        'Debug.Print url
        P_FaitLink = url
    End If
End Function

Public Function P_RecupSrvNom(ByVal v_num As Long, _
                              ByRef r_nom As String) As Integer

    Dim sql As String

    sql = "select SRV_Nom from Service" _
        & " where SRV_Num=" & v_num
    If Odbc_RecupVal(sql, r_nom) = P_ERREUR Then
        P_RecupSrvNom = P_ERREUR
        Exit Function
    End If
    
    P_RecupSrvNom = P_OK
    
End Function

Public Function P_RecupPosteNomfct(ByVal v_numposte As Long, _
                                   ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "select FT_Libelle from Poste, FctTrav" _
        & " where PO_Num=" & v_numposte _
        & " and FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupPosteNomfct = P_ERREUR
        Exit Function
    End If
    
    P_RecupPosteNomfct = P_OK
    
End Function

Public Function P_RecupNomFonction(ByVal v_numfct As Long, _
                                   ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "select FT_Libelle from Poste, FctTrav" _
        & " where FT_Num=" & v_numfct
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupNomFonction = P_ERREUR
        Exit Function
    End If
    
    P_RecupNomFonction = P_OK
    
End Function


Public Function P_RecupPosteNom(ByVal v_numposte As Long, _
                                ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "select PO_Libelle from Poste" _
        & " where PO_Num=" & v_numposte
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupPosteNom = P_ERREUR
        Exit Function
    End If
    
    P_RecupPosteNom = P_OK
    
End Function

Public Function P_SaisirUtilIdent(ByVal x As Integer, _
                                  ByVal y As Integer, _
                                  ByVal L As Integer, _
                                  ByVal H As Integer) As Integer

    Dim codutil As String, mpasse As String, sql As String
    Dim deuxieme_saisie As Boolean, bad_util As Boolean
    Dim nb As Integer, reponse As Integer
    Dim lnb As Long, lbid As Long
    Dim rs As rdoResultset
    Dim oMD5 As CMD5
    
    ' Pour le cryptage MD5
    Set oMD5 = New CMD5
    
    nb = 1
    deuxieme_saisie = False
    
    'Saisie du code utilisateur
lab_debut:
    Call SAIS_Init
    Call SAIS_InitOblig(False)
    If deuxieme_saisie Then
        Call SAIS_InitTitreHelp("Confirmez votre mot de passe", "")
        Call SAIS_AddChampComplet("Mot de passe (confirmation)", 15, 15, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
    Else
        Call SAIS_InitTitreHelp("Identification", p_chemin_appli + "\help\kalidoc.chm;demarrage.htm")
        Call SAIS_AddChamp("Code d'accès", 50, 20, SAIS_TYP_TOUT_CAR, False)
        Call SAIS_AddChampComplet("Mot de passe", 15, 15, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
    End If
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
lab_saisie:
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        P_SaisirUtilIdent = P_NON
        Exit Function
    End If
        
    If deuxieme_saisie Then
        If mpasse <> SAIS_Saisie.champs(0).sval Then
            MsgBox "Vous n'avez pas saisi le même mot de passe.", vbOKOnly + vbExclamation, ""
            deuxieme_saisie = False
            GoTo lab_debut
        End If
        ' Maj du mot de passe utilisateur
        sql = "select * from UtilAppli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & UCase(codutil) & "'"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo 0
        If rs.EOF Then GoTo err_no_resultset
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
        If rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "" Then
            rs("UAPP_MotPasse").Value = STR_Crypter(UCase(mpasse))
        ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" Then
            rs("UAPP_MotPasse").Value = STR_Crypter_New(UCase(mpasse))
        ElseIf rs("UAPP_TypeCrypt").Value = "md5" Then
            rs("UAPP_MotPasse").Value = oMD5.MD5(mpasse)
        End If
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
    Else
        codutil = SAIS_Saisie.champs(0).sval
        mpasse = SAIS_Saisie.champs(1).sval
    End If
'MsgBox "attention mode debug"
'p_NumUtil = 948
'codutil = "BAYARD"
'GoTo lab_ok
    If codutil = "ROOT" And mpasse = "007" Then
        p_CodeUtil = "ROOT"
        p_NumUtil = p_SuperUser     'P_SUPER_UTIL
        P_SaisirUtilIdent = P_OUI
        Exit Function
    End If
    
    'Recherche de cet utilisateur
    sql = "select U_Num, UAPP_MotPasse, UAPP_TypeCrypt from Utilisateur, UtilAppli" _
        & " where UAPP_Code='" & UCase(codutil) & "'" _
        & " and UAPP_APPNum=" & p_appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_Num=UAPP_UNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_SaisirUtilIdent = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        bad_util = True
    Else
        If rs("UAPP_MotPasse").Value <> "" Then
            If (rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "") And STR_Decrypter(rs("UAPP_MotPasse").Value) <> UCase(mpasse) Then
                bad_util = True
            ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" And STR_Decrypter_New(rs("UAPP_MotPasse").Value) <> UCase(mpasse) Then
                bad_util = True
            ElseIf rs("UAPP_TypeCrypt").Value = "md5" And rs("UAPP_MotPasse").Value <> oMD5.MD5(mpasse) Then
                bad_util = True
            Else
                p_NumUtil = rs("U_Num").Value
                rs.Close
                GoTo lab_ok
            End If
        Else
            bad_util = False
        End If
        rs.Close
    End If
    If bad_util Then
        MsgBox "Identification inconnue.", vbOKOnly + vbExclamation, ""
        nb = nb + 1
        If nb > 3 Then
            P_SaisirUtilIdent = P_ERREUR
            Exit Function
        End If
        SAIS_Saisie.champs(1).sval = ""
        GoTo lab_saisie
    Else
        deuxieme_saisie = True
        GoTo lab_debut
    End If
    
lab_ok:
    p_CodeUtil = UCase(codutil)
    
    P_SaisirUtilIdent = P_OUI
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultset " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_no_resultset:
    MsgBox "Pas de ligne pour " & sql, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_edit:
    MsgBox "Erreur Edit " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_affecte:
    MsgBox "Erreur Affectation " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_update:
    MsgBox "Erreur Update " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
End Function

Public Function P_SaisirUtilPasswd() As Integer

    Dim mpasse As String, sql As String, lib As String, titre As String
    Dim nb As Integer, etape As Integer
    Dim rs As rdoResultset
    
    nb = 1
    etape = 0
    
    'Saisie du code utilisateur
lab_debut:
    Call SAIS_Init
    'Call SAIS_InitOblig(False)
    If etape = 0 Then
        titre = "Saisissez votre mot de passe actuel"
        lib = "Mot de passe"
    ElseIf etape = 1 Then
        titre = "Saisissez le nouveau mot de passe"
        lib = "Nouveau mot de passe"
    Else
        titre = "Confirmez votre nouveau mot de passe"
        lib = "Mot de passe (confirmation)"
    End If
    Call SAIS_InitTitreHelp(titre, "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call SAIS_AddChampComplet(lib, 15, 15, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
lab_saisie:
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        P_SaisirUtilPasswd = P_NON
        Exit Function
    End If
        
    Select Case etape
    Case 0
        
        sql = "select * from UtilAppli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & p_CodeUtil & "'" _
            & " and (((UAPP_TypeCrypt = 'kalidoc' or UAPP_TypeCrypt = '') AND UAPP_MotPasse='" & UCase(STR_Crypter(SAIS_Saisie.champs(0).sval)) & "') " _
            & " or (UAPP_TypeCrypt = 'kalidocnew' AND UAPP_MotPasse='" & UCase(STR_Crypter_New(SAIS_Saisie.champs(0).sval)) & "'))"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
        On Error GoTo 0
        If rs.EOF Then
            rs.Close
            MsgBox "Mot de passe incorrect.", vbOKOnly + vbExclamation, ""
            Exit Function
        End If
        rs.Close
        etape = 1
        GoTo lab_debut
    Case 1
        mpasse = UCase(SAIS_Saisie.champs(0).sval)
        etape = 2
        GoTo lab_debut
    Case 2
        If mpasse <> UCase(SAIS_Saisie.champs(0).sval) Then
            MsgBox "Vous n'avez pas saisi le même mot de passe.", vbOKOnly + vbExclamation, ""
            etape = 1
            GoTo lab_debut
        End If
        ' Maj du mot de passe utilisateur
        sql = "select * from Utilappli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & p_CodeUtil & "'"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo 0
        If rs.EOF Then GoTo err_no_resultset
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
        rs("UAPP_TypeCrypt").Value = "kalidocnew"
        rs("UAPP_MotPasse").Value = STR_Crypter_New(mpasse)
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
    End Select
    
    P_SaisirUtilPasswd = P_OUI
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultset " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_no_resultset:
    MsgBox "Pas de ligne pour " & sql, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_edit:
    MsgBox "Erreur Edit " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_affecte:
    MsgBox "Erreur Affectation " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_update:
    MsgBox "Erreur Update " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
End Function

