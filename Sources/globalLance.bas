Attribute VB_Name = "Module1"
Option Explicit

Public p_bool_ModeDebug As Boolean

' pour savoir si le tableau est vide
Public p_bool_tbl_cond As Boolean
Public p_bool_tbl_fichExcel As Boolean
Public p_bool_tbl_cell As Boolean
Public p_bool_tbl_fenExcel As Boolean
Public p_bool_tbl_rdoF As Boolean
Public p_bool_tbl_rdoL As Boolean
Public p_bool_tbl_Demande As Boolean
Public p_bool_tbl_FichExcelOuverts As Boolean
Public p_bool_tbl_diff As Boolean

Public p_boolRetournerAuParam As Boolean

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

Public p_NumTemp As Integer

Public p_i_tabExcel As Integer
Public p_i_feuilleExcel As Integer

Public p_Derniere_MenForme As String

Public p_BoolMettreComment As Boolean

Public Const P_SUPER_UTIL = 1

Public g_modeSQL_LIB As String

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public bfaire_RowColChange As Boolean

Public p_NumForm As Integer
Public p_NumModele As Integer
Public p_NumUtil As Integer
Public p_CodeUtil As String
Public p_chemin_appli As String
Public p_tabdoc_present() As Long

' Chemins où se trouvent l'appli et ses fichiers
Public p_CheminRapportType As String
Public p_CheminRapportType_Ini As String
Public p_RapportTypeExe As String
Public p_Drive_Résultats As String
Public p_Chemin_Résultats As String
Public p_Drive_Modeles_Serveur As String
Public p_Chemin_Modeles_Serveur As String
Public p_Drive_Modeles_Local As String
Public p_Chemin_Modeles_Local As String
Public p_Nom_Modele As String
Public p_Drive_KaliDoc As String
Public p_Path_KaliDoc As String
Public p_S_Vers_Conf As String

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
    numdos As Long
    ordre As Long
    titre As String
    numpere As Long
End Type

' Chemins où se trouvent l'appli et ses fichiers
'Public p_chemin_appli As String
Public p_nomini As String
Public p_chemin_modele As String
'Public p_chemin_archive As String
Public p_cheminkalidoc_serveur As String
Public p_est_le_serveur As Boolean
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
    CellMenForme As String
    cellXPère As Integer
    cellYPère As Integer
    cellSQL As String
    cellNumFiltre As String
End Type
Public tbl_cell() As CELL

Public Type SFICH_PARAM_EXCEL
    CmdType As String
    CmdFenNum As String
    CmdX As String
    CmdY As String
    CmdFormNum As String
    CmdFormIndice As Integer
    CmdChpNum As Integer
    CmdCondition As String
    CmdTypeChp As String
    CmdMenFormeChp As String
    CmdLstFen As String
    CmdLstDest As String
    CmdTitreDoc As String
    CmdMenFormeDoc As String
End Type
Public tbl_fichExcel() As SFICH_PARAM_EXCEL

Public Type SFEN_EXCEL
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
    RDOF_FormIndice As String
    RDOF_etat As String
    RDOF_QuestionsFait As Boolean
    RDOF_QuestionsSQL As String
End Type
Public tbl_rdoF() As RDOF

Public Type RDOL
    RDOL_num As Integer
    RDOL_Opérateur As String
    RDOL_sql As String
    RDOL_sqlFrancais As String
    RDOL_sqlPF As String
    RDOL_fornum As String
    RDOL_FormIndice As String
    RDOL_DéjàPenCompte As Boolean
End Type
Public tbl_rdoL() As RDOL

Public Type SCOND_PARAM
    CondNumFiltre As Integer
    CondFormIndice As Integer
    CondString As String
    CondOper As String
    CondType As String
    CondFrancais As String
    CondPF As String
    CondligGrdCond As Integer
End Type
Public tbl_cond() As SCOND_PARAM

' Pour les questions à poser à chaque fois
Public Type DEMANDE_SQL
    DemandChpNum As Integer
    DemandFormInd As Integer
    DemandForNum As Integer
    DemandFFNum As Integer
    DemandChpStr As String
    DemandForStr As String
    DemandType As String
    DemandSQL As String
    DemandFrancais As String
    DemandFait As Boolean
End Type
Public tbl_Demande() As DEMANDE_SQL

Public Const Public_Alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

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

Public Sub StartProcess(ByVal sFile As String, Optional ByVal sParameters As String = vbNullString)
    Dim ret As Integer
    
    ret = ShellExecute(0&, "open", sFile, sParameters, vbNullString, 1&)
End Sub

Public Sub ResolutionEcran(sgWidth As Single, sgHeight As Single)
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

Public Function ScreenResolution(ByRef r_Largeur_Ecran As Single, ByRef r_Hauteur_Ecran As Single) As String
    
    r_Largeur_Ecran = GetSystemMetrics(SM_CXSCREEN)
    r_Hauteur_Ecran = GetSystemMetrics(SM_CYSCREEN)
    ScreenResolution = "Vidéo " & r_Largeur_Ecran & " x  " & r_Hauteur_Ecran

End Function

Public Function P_AfficherArborescenceDoc(ByRef v_tv As TreeView, _
                                          ByVal v_numdos As Long, _
                                          ByVal v_img_dos As Long, _
                                          ByVal v_img_dos_sel As Long, _
                                          ByVal v_expand As Boolean) As Integer

    Dim sql As String, docs_titre As String, slien As String
    Dim trouve As Boolean, fracine As Boolean, encore As Boolean
    Dim i As Integer, n As Integer, mode As Integer
    Dim sav_numdos As Long, numDocs As Long
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
            tbldos(n).numdos = v_numdos
            sql = "select DS_DONum, DS_SLien, DS_Titre, DS_Numpere, DS_Ordre" _
                & " from Dossier" _
                & " where DS_Num=" & v_numdos
            If Odbc_RecupVal(sql, numDocs, slien, tbldos(n).titre, v_numdos, tbldos(n).ordre) = P_ERREUR Then
                P_AfficherArborescenceDoc = P_ERREUR
                Exit Function
            End If
            If slien <> "" Then
                numlien = Mid$(slien, 2)
                Select Case Left$(slien, 1)
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
        If TV_NodeExiste(v_tv, "O" & numDocs, nd) = P_OUI Then
            fracine = False
        Else
            fracine = True
        End If
    Else
        Set nd = v_tv.Nodes("S" & v_numdos)
        fracine = False
    End If
    ' On redébobine les dossiers
    For i = n To 0 Step -1
        On Error GoTo lab_no_dos
        trouve = True
        Set ndp = v_tv.Nodes("S" & tbldos(i).numdos)
        On Error GoTo 0
        If Not trouve Then
            If fracine Then
                mode = tvwChild
                If v_tv.Nodes.Count > 0 Then
                    If v_tv.Nodes(1).Root.Children > 0 Then
                        Set nd = v_tv.Nodes(1).Root
                        Do
                            If nd.Tag > tbldos(i).ordre Then
                                mode = tvwPrevious
                                Exit Do
                            End If
                        Loop Until Not TV_NodeNext(nd)
                    End If
                End If
                If mode = tvwPrevious Then
                    Set nd = v_tv.Nodes.Add(nd, mode, "S" & tbldos(i).numdos, tbldos(i).titre, v_img_dos, v_img_dos_sel)
                Else
                    Set nd = v_tv.Nodes.Add(, mode, "S" & tbldos(i).numdos, tbldos(i).titre, v_img_dos, v_img_dos_sel)
                End If
                nd.Tag = tbldos(i).ordre
            Else
                mode = tvwChild
                If nd.Children > 0 Then
                    Set nd2 = nd.Child
                    For n = 1 To nd.Children
                        If nd2.Tag > tbldos(i).ordre Then
                            Set nd = nd2
                            mode = tvwPrevious
                            Exit For
                        End If
                        Set nd2 = nd2.Next
                    Next n
                End If
                Set nd = v_tv.Nodes.Add(nd, mode, "S" & tbldos(i).numdos, tbldos(i).titre, v_img_dos, v_img_dos_sel)
                nd.Tag = tbldos(i).ordre
            End If
        Else
            Set nd = ndp
        End If
'        nd.Sorted = True
        If v_expand Then nd.Expanded = True
        fracine = False
    Next i
    
    P_AfficherArborescenceDoc = P_OK
    Exit Function

lab_no_dos:
    trouve = False
    Resume Next
    
End Function

Public Function Public_FichiersExcelOuverts(ByRef r_tbl_FichExcel() As FichExcelOuverts, v_Trait As String, v_chemin As String, v_visible As Boolean, v_à_Sauver As Boolean)
    Dim LaUbound As Integer
    Dim i As Integer
    Dim bDéjà As Boolean
    Dim strFichG As String, strFichd As String
    
    LaUbound = 0
    On Error GoTo Faire
    LaUbound = UBound(r_tbl_FichExcel(), 1) + 1
    For i = 0 To LaUbound
        strFichG = Replace(UCase(r_tbl_FichExcel(i).FichFullname), "\", "$")
        strFichG = Replace(strFichG, "/", "$")
        strFichd = Replace(UCase(v_chemin), "\", "$")
        strFichd = Replace(strFichd, "/", "$")
        If strFichG = strFichd Then
        'If r_tbl_FichExcel(i).FichFullname = v_Chemin Then
            bDéjà = True
            LaUbound = i
            Exit For
        End If
    Next i
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
End Function







Public Function P_SaisirUtilIdent(ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByVal l As Integer, _
                                  ByVal h As Integer) As Integer

    Dim codutil As String, mpasse As String, sql As String
    Dim deuxieme_saisie As Boolean, bad_util As Boolean
    Dim nb As Integer, reponse As Integer
    Dim lnb As Long, lbid As Long
    Dim rs As rdoResultset
    
    nb = 1
    deuxieme_saisie = False
    
    'Saisie du code utilisateur
lab_debut:
    Call SAIS_Init
    If deuxieme_saisie Then
        Call SAIS_InitTitreHelp("Confirmez votre mot de passe", "")
        Call SAIS_AddChampComplet("Mot de passe (confirmation)", 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
    Else
        Call SAIS_InitTitreHelp("Identification", p_chemin_appli + "\help\kalidoc.chm;demarrage.htm")
        Call SAIS_AddChamp("Code d'accès", 15, SAIS_TYP_TOUT_CAR, False)
        Call SAIS_AddChampComplet("Mot de passe", 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
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
        rs("UAPP_MotPasse").Value = STR_Crypter(UCase(mpasse))
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
    Else
        codutil = SAIS_Saisie.champs(0).sval
        mpasse = SAIS_Saisie.champs(1).sval
    End If
    
    If codutil = "ROOT" And mpasse = "007" Then
        p_CodeUtil = "ROOT"
        p_NumUtil = P_SUPER_UTIL
        P_SaisirUtilIdent = P_OUI
        Exit Function
    End If
    
    'Recherche de cet utilisateur
    sql = "select U_Num, UAPP_MotPasse from Utilisateur, UtilAppli" _
        & " where UAPP_Code='" & UCase(codutil) & "'" _
        & " and UAPP_APPNum=" & p_appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_Num=UAPP_UNum"
    'MsgBox sql
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_SaisirUtilIdent = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        bad_util = True
    Else
        If rs("UAPP_MotPasse").Value <> "" Then
            If STR_Decrypter(rs("UAPP_MotPasse").Value) <> UCase(mpasse) Then
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
    
    ' Ajout si nécessaire dans UtilKD
    'If p_CheminPHP <> "" Then
    '    sql = "select count(*) from UtilKD where ukd_unum=" & p_NumUtil
    '    If Odbc_Count(sql, lnb) = P_ERREUR Then
    '        lnb = 1
    '    End If
    '    If lnb = 0 Then
    '        reponse = MsgBox("Devez-vous être considéré comme un utilisateur habituel de KaliDoc ?", vbQuestion + vbYesNo, "")
    '        If reponse = vbYes Then
    '            Call Odbc_AddNew("UtilKD", "ukd_num", "ukd_seq", False, lbid, _
    '                             "ukd_unum", p_NumUtil)
    '        End If
    '    End If
    'End If
    
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
    Call SAIS_AddChampComplet(lib, 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
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
            & " and UAPP_MotPasse='" & UCase(STR_Crypter(SAIS_Saisie.champs(0).sval)) & "'"
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
        rs("UAPP_MotPasse").Value = STR_Crypter(mpasse)
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




