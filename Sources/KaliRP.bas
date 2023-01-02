Attribute VB_Name = "KaliRPMain"
Option Explicit

Public P_MODCNUM_ou_MODELE As String

Public p_message As Variant
Public p_message_choisir As Integer
Public p_message_un_par_un As Boolean
Public p_mode_acces As String
Public p_peut_creer As Boolean

Public p_changement_de_champ As Boolean

Public p_MultiSite As Boolean
Public p_scmd As String
Public p_num_ent_juridique As Integer

' Définition des emplacements des paramètres passés à scmd
Public Const p_SCMD_CHEMIN_APPLI = 0
Public Const p_SCMD_TYPE_BASE = 1
Public Const p_SCMD_NOM_BASE = 2
Public Const p_SCMD_CHEMIN_INI = 4
Public Const p_SCMD_PARAM_SUPPLEMENT = 5
' paramètres supplémentaires
Public Const p_SCMD_PARAM_NUMFORM = 0
Public Const p_SCMD_PARAM_NUMUTIL = 1
Public Const p_SCMD_PARAM_NUMFILTRE = 2

' Chemins où se trouvent l'appli et ses fichiers
Public p_chemin_appli As String
Public p_nomini As String
Public p_chemin_modele As String
Public p_chemin_archive As String
Public p_cheminkalidoc_serveur As String
Public p_chemin_RapportType_exe As String
Public p_chemin_RapportType_ini As String
Public p_chemin_Scanx_ini As String
Public p_chemin_Scanx_exe As String
Public p_gerer_btnfusion As Boolean
Public p_est_le_serveur As Boolean
Public p_est_le_convertisseur As Boolean
Public p_Nom_BDD As String

' Chemins où se trouvent l'appli et ses fichiers
Public p_CheminRapportType As String
Public p_CheminRapportType_Ini As String
Public p_RapportTypeExe As String
Public p_Drive_Résultats As String
Public p_Chemin_Résultats As String
Public p_HTTP_Résultats As String
Public p_Chemin_Modeles_Serveur As String
Public p_Chemin_Modeles_Local As String
Public p_CheminDossierTravailLocal As String
Public p_Nom_Modele As String
Public p_Drive_KaliDoc As String
Public p_Path_KaliDoc As String
Public p_cheminKW As String
Public p_S_Vers_Conf As String

Public p_numFichier_Liens As Long
Public p_numRandom As String

Public Type TB_CONDITION_FILTRE
    titre As String
    Condition As String
End Type
Public P_tb_conditions() As TB_CONDITION_FILTRE

' pour compatibilité
Public p_numdocinit As Integer
Public p_CheminDoc As String

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


Sub Main()
    
    Dim scmd As String
    Dim direct As Boolean, fok As Boolean
    Dim i As Integer
    Dim frm As Form
    Dim fp As Integer
    Dim NomFichierOut As String

    If App.PrevInstance Then
        If MsgBox("KaliRP a déjà été lancé." & vbCrLf & "Confirmez-vous quand même le lancement ?", vbQuestion + vbYesNo, "") = vbNo Then
            End
        End If
    End If
    
    p_version_KaliRP = "V 4.19.06B"
    Splash.lblVersion.Caption = p_version_KaliRP
    
    Splash.Show
    DoEvents
    'fermeture au bout de 5 secondes
    Splash.CloseAfter 20
    DoEvents
    
    p_LibLienDétail = "Accès au Détail"
    
    If FICH_FichierExiste(p_chemin_appli & "c:\kalidoc\lance2.exe") Then
        Call MsgBox("Mise à jour du programme 'Lance'à faire.", vbOKOnly + vbInformation, "")
        fok = False
        For i = 1 To 100
            If FICH_EffacerFichier("c:\kalidoc\lance.exe", False) = P_OK Then
                fok = True
                Exit For
            End If
            SYS_Sleep (500)
        Next i
        If fok Then
            fok = False
            For i = 1 To 100
                If FICH_RenommerFichier(p_chemin_appli & "c:\kalidoc\lance2.exe", "c:\kalidoc\Lance.exe") = P_OK Then
                    Call MsgBox("Mise à jour du programme 'Lance' effectuée.", vbOKOnly + vbInformation, "")
                    fok = True
                    Exit For
                End If
                SYS_Sleep (500)
            Next i
        End If
    End If
    
    p_NumUtil = 0
    
    ' Param de l'application
    scmd = Command$
'MsgBox "attention mode debug"
'scmd = "C:\Kalidoc;PG;chpa_C3;0;KaliRP_CHPA.ini"
    'Call MsgBox("scmd=" & scmd, vbOKOnly + vbInformation, "")
    'scmd = "c:\KaliDoc;PG;kalidoc_hmz;0;KaliRP.ini"
    'scmd = "C:/KaliDoc;PG;vm_STBRIEUC;0;vm_STBRIEUC.ini;23|73|4"
    'scmd = "C:/KaliDoc;PG;kalidoc_demo;0;KaliRP_demo.ini;23|73|1;DEBUG;"
    'scmd = "C:/KaliDoc;PG;VM_KALI_STANGELY;0;KaliRP_StAng.ini"
    'scmd = "C:/KaliDoc;PG;VM_KALI_STANGELY;0;KaliRP_StAng.ini"
    'scmd = "C:\Kalitech\KaliDoc_315;PG;kalidoc_eteer;0;KaliRP_eteer.ini"
    'scmd = "C:\Kalitech\KaliDoc_315;PG;kalidoc_hpsj;0;KaliRP_hpsj.ini"
    'scmd = "C:\Kalitech\KaliDoc_315;PG;kalidoc_blois;0;KaliRP_blois.ini"
    'scmd = "C:\Kalitech\KaliDoc_315;PG;CHUB;0;KaliRP_CHUB.ini"
    'scmd = "C:\Kalitech\KaliDoc_315;PG;kalidoc_meudon;0;KaliRP_meudon.ini"
'scmd = "C:\KaliDoc;PG;kalidoc_KALIDEV;0;KaliRP_KALIDEV.ini;0|ROOT|6"
'scmd = "C:\KaliDoc;PG;kalidoc_KALIDEV;0;KaliRP_KALIDEV.ini;0|ROOT|"
'scmd = "c:\kalidoc;PG;pg://dev.dev.kali/dev;0;kalidoc_choix_PG.ini;0|2715|0"
'scmd = "c:\kalidoc;PG;pg://dev.dev.kali/dev;0;kalidoc_choix_PG.ini;0|ROOT|0"
    'scmd = "DEBUG"
    
    If scmd = "DEBUG" Then
        direct = False
        P_MODE_DEBUG = True
        If init_param_debug() = P_ERREUR Then
            End
        End If
    Else
        If init_param_exe(scmd, direct) = P_ERREUR Then
            End
        End If
    End If
    
    p_message = ""
    p_message_choisir = 0
    p_message_un_par_un = False
    
    ' Menu
    Unload Splash
    
    p_scmd = scmd
        
    MenuG.Show
    
End Sub

Public Function MODCNUM_ou_MODELE()
    Dim sql As String, rs As rdoResultset
    
    ' savoir si on est en MODCNUM ou MODELE
    sql = "select column_name from information_schema.columns where table_name='docsnaturemodele' and (column_name='donm_modele' or column_name='donm_modcnum')"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        MODCNUM_ou_MODELE = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        MODCNUM_ou_MODELE = rs(0).Value
        rs.MoveNext
    Wend
End Function
Public Function reformer_BCR(ByVal v_ffnum, ByVal v_bcr)
    Dim i As Integer, UnBcr As String, param1 As String, param2 As String, param3 As String
    Dim ff_fornums As String
    Dim j As Integer
    Dim s As String
    Dim param4 As String, param5 As String
    Dim numchp As Integer, nomChp As String
    Dim sql As String, rs As rdoResultset
    Dim nb As Integer
    
    For i = 0 To STR_GetNbchamp(v_bcr, "§")
        UnBcr = STR_GetChamp(v_bcr, "§", i)
        If UnBcr <> "" Then
            param1 = STR_GetChamp(UnBcr, "¤", 0)
            nb = STR_GetNbchamp(param1, ":")
            If nb = 2 Then
                param2 = STR_GetChamp(UnBcr, "¤", 1)
                param3 = STR_GetChamp(UnBcr, "¤", 2)
                param4 = STR_GetChamp(UnBcr, "¤", 3)
                param5 = STR_GetChamp(UnBcr, "¤", 4)
                nomChp = STR_GetChamp(param1, ":", 1)
                sql = "Select ff_fornums from filtreform where ff_num=" & v_ffnum
                Call Odbc_RecupVal(sql, ff_fornums)
                nb = STR_GetNbchamp(ff_fornums, "*")
                sql = "Select forec_num from formetapechp where forec_nom='" & nomChp & "'"
                If nb = 2 Then
                    sql = sql & " And (forec_fornum=" & STR_GetChamp(ff_fornums, "*", 1) & ")"
                Else
                    sql = sql & " And (forec_fornum=" & STR_GetChamp(ff_fornums, "*", 1) & " OR forec_fornum=" & STR_GetChamp(ff_fornums, "*", 2) & ")"
                End If
                Call Odbc_SelectV(sql, rs)
                If rs.EOF Then
                    MsgBox ("fonction reformer_BCR : Vide pour " & sql)
                    numchp = 0
                Else
                    numchp = rs("forec_num")
                End If
                s = "CHP:" & numchp & ":" & nomChp & "¤" & param2 & "¤" & param3 & "¤" & param4 & "¤" & param5
            Else
                'param3 = STR_GetChamp(UnBcr, "¤", 2)
                's = STR_GetChamp(param3, ":", 0)
                'If s <> "VAL" And s <> "DATE" Then
                '    param2 = STR_GetChamp(UnBcr, "¤", 1)
                '    param3 = "VAL:" & STR_GetChamp(param3, ":", 1)
                '    param4 = STR_GetChamp(UnBcr, "¤", 3)
                '    param5 = STR_GetChamp(UnBcr, "¤", 4)
                '    UnBcr = param1 & "¤" & param2 & "¤" & param3 & "¤" & param4 & "¤" & param5
                'End If
                s = UnBcr
                nb = STR_GetNbchamp(s, "¤")
                If nb = 4 Then
                    s = s & "¤"
                End If
            End If
            reformer_BCR = s
        End If
    Next i
End Function

Private Function Choisir_Ini(ByVal v_p_nomini As String, ByVal v_p_chemin_appli As String) As String
    Dim s As String
    Dim n As Integer
    Dim nomfich As String
    
lab_choix:
    'Call FRM_ResizeForm(Me, 0, 0)
    nomfich = v_p_nomini
    
    If v_p_nomini = "" Then
        nomfich = SYS_GetIni("FICHIER_INI", "DERNIER", v_p_chemin_appli & "\KaliRP_dernier_ini_ouvert.txt")
    End If
    
    Call CL_Init
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("Ouvrir le ini", "", 0, 0, 1500)
    Call CL_AddBouton("ODBC", "", 0, 0, 1500)
    Call CL_AddBouton("Créer à partir de", "", 0, 0, 1800)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    s = Dir$(v_p_chemin_appli & "\kaliRP*.ini")
    n = 0
    While s <> ""
        If UCase(v_p_chemin_appli & "\" & s) = UCase(nomfich) Then
            Call CL_AddLigne(s, n, "", True)
        Else
            Call CL_AddLigne(s, n, "", False)
        End If
        n = n + 1
        s = Dir$()
    Wend
    Call CL_AffiSelFirst
    
    If n = 0 Then
        MsgBox "Aucun fichier ini dans " & v_p_chemin_appli
        End
    End If
    
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 4 Then
        Choisir_Ini = P_NON
        Exit Function
    End If
    
    nomfich = v_p_chemin_appli & "\" & CL_liste.lignes(CL_liste.pointeur).texte
    
    ' Ouvrir
    If CL_liste.retour = 1 Then
        If FICH_FichierExiste(nomfich) Then
            Call SYS_StartProcess(nomfich)
        Else
            MsgBox "Le fichier '" & nomfich & "' n'a pas été trouvé."
        End If
        GoTo lab_choix
    End If
    
    ' ODBC
    If CL_liste.retour = 2 Then
        Call SYS_StartProcess("c:\windows\system32\control", "odbccp32.cpl")
        GoTo lab_choix
    End If
        
    ' Nouveau à partir de
    If CL_liste.retour = 3 Then
        If FICH_FichierExiste(nomfich) Then
            Choisir_Ini = nomfich
        Else
            MsgBox nomfich & " n'existe pas"
        End If
encore:
        s = InputBox("Nom du nouveau .ini", "Créer un .ini", "kaliRP_nouveau.ini")
        If s <> "" Then
            If FICH_FichierExiste(UCase(v_p_chemin_appli & "\" & s)) Then
                MsgBox s & " existe déjà"
                GoTo encore
            Else
                If FICH_CopierFichier(nomfich, v_p_chemin_appli & "\" & s) = P_ERREUR Then
                    MsgBox "Erreur pour copier " & nomfich & " vers " & s
                Else
                    nomfich = v_p_chemin_appli & "\" & s
                    If FICH_FichierExiste(nomfich) Then
                        Call SYS_StartProcess(nomfich)
                    End If
                End If
            End If
        End If
    End If
    
    If FICH_FichierExiste(nomfich) Then
        Choisir_Ini = nomfich
        '
        Call SYS_PutIni("FICHIER_INI", "DERNIER", nomfich, v_p_chemin_appli & "\KaliRP_dernier_ini_ouvert.txt")
        '
    Else
        MsgBox nomfich & " n'existe pas"
    End If
End Function

Public Function P_GetNomUtil(ByVal v_numutil As String)
    Dim rs As rdoResultset, sql As String
    
    sql = "select * from utilisateur where u_num = " & v_numutil
    If v_numutil <> "" Then
        If Odbc_SelectV(sql, rs) <> P_ERREUR Then
            If Not rs.EOF Then
                P_GetNomUtil = LCase(rs("u_prenom")) & " " & UCase(rs("u_nom"))
            Else
                P_GetNomUtil = "Utilisateur n° " & v_numutil
            End If
        End If
    End If
End Function

Public Function FRM_AuPremierPlan(hwnd As Long) As Long

    FRM_AuPremierPlan = SetWindowPos(hwnd, HWND_TOP, 100, 0, 0, 0, FLAGS)
    
End Function

Private Function init_param_debug() As Integer
    
    Dim stype_bdd As String, nom_bdd As String
    Dim ask_enreg As Boolean
    Dim reponse As Integer
        
    p_chemin_appli = "c:\kalidoc"
    'RapportType.MnuCheminAppli.Caption = p_chemin_appli
    
Lab_Choisir_Ini:
    p_chemin_appli = "c:\kalidoc"
    If Environ$("KALIDOC") <> "" Then
        p_chemin_appli = Environ$("KALIDOC")
    End If
    
    ' chercher les .ini
    p_nomini = Choisir_Ini(p_nomini, p_chemin_appli)
    
    If p_nomini = "" Then
        MsgBox "Vous devez choisir un fichier ini"
        GoTo Lab_Choisir_Ini
    End If
    
    'p_nomini = InputBox("Chemin du .ini : ", , "c:\kalidoc\KaliRP.ini")
    'If p_nomini = "" Then
    '    init_param_debug = P_ERREUR
    '    Exit Function
    'End If
    
    ask_enreg = False
    ' Type de base
    stype_bdd = SYS_GetIni("BASE", "TYPE", p_nomini)
    If stype_bdd = "" Then
lab_sais_typb:
        If stype_bdd = "" Then
            init_param_debug = P_ERREUR
            Exit Function
        End If
        If stype_bdd <> "MDB" And stype_bdd <> "PG" Then
            GoTo lab_sais_typb
        End If
        ask_enreg = True
    End If
    ' Nom base
    nom_bdd = SYS_GetIni("BASE", "NOM", p_nomini)
    If nom_bdd = "" Then
        nom_bdd = InputBox("Nom de la base : ", , "c:\kalidoc\kalidoc.mdb")
        If nom_bdd = "" Then
            init_param_debug = P_ERREUR
            Exit Function
        End If
        ask_enreg = True
    End If
    ' Enregistrement des infos base
    If ask_enreg Then
        reponse = MsgBox("Voulez-vous enregistrer les informations saisies ?", vbQuestion + vbYesNo, "")
        If reponse = vbYes Then
            Call SYS_PutIni("BASE", "TYPE", stype_bdd, p_chemin_appli & "\kalidoc.ini")
            Call SYS_PutIni("BASE", "NOM", nom_bdd, p_chemin_appli & "\kalidoc.ini")
        End If
    End If

    ' Connexion à la base
    If Odbc_Init(stype_bdd, nom_bdd, True) = P_ERREUR Then
        init_param_debug = P_ERREUR
        Exit Function
    End If
    p_Nom_BDD = nom_bdd
    
    P_MODCNUM_ou_MODELE = MODCNUM_ou_MODELE()

    Call charger_var_ini
    
    init_param_debug = P_OK
    
End Function



Private Function init_param_exe(ByVal v_scmd As String, _
                                ByRef r_direct As Boolean) As Integer
                                  
    Dim s As String, stype_bdd As String, nom_bdd As String, sql As String
    Dim saction As String, snumdos As String, tbldoscli() As String, snumcli As String
    Dim snumdosp As String, titredos As String, lstresp As String, cmd_direct As String
    Dim cmd_direct_compl As String, prm_direct As String
    Dim etat As Boolean
    Dim nbprm As Integer, n As Integer, i As Integer
    Dim numutil As Long
    Dim frm As Form
    Dim rs As rdoResultset
    Dim strUsage As String
    
    strUsage = "Usage : KaliRP " & p_version_KaliRP & " <Chemin application>;<Type BDD>;<Nom BDD>;<MULT>;[NOM INI];<PARAM SUPPLEMENTAIRES>"
    strUsage = strUsage & vbCrLf & vbCrLf & "<PARAM SUPPLEMENTAIRES> = <Num_Formulaire> | <Num_Utilisateur> | <Num_Modele>"
    strUsage = strUsage & vbCrLf & vbCrLf & "Ex : C:/KaliDoc;PG;kalidoc_demo;0;KaliRP_demo.ini;23|73|1;DEBUG;"
    nbprm = STR_GetNbchamp(v_scmd, ";")
    If nbprm < 4 Then
        ' c:\kalidoc;PG;kalitech;0;kalidoc.ini
        Call MsgBox(strUsage & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    
    ' 1- Chemin appli
    p_chemin_appli = STR_GetChamp(v_scmd, ";", p_SCMD_CHEMIN_APPLI)
    If p_chemin_appli = "" Then
        Call MsgBox("<Chemin application> est vide." & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCrLf & vbCrLf _
                    & strUsage, vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
'MsgBox p_chemin_appli
    ' 2- Type de base
    stype_bdd = STR_GetChamp(v_scmd, ";", p_SCMD_TYPE_BASE)
    If stype_bdd <> "PG" And stype_bdd <> "MDB" Then
        Call MsgBox("Type de Base incorrect : " & stype_bdd, vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
'MsgBox stype_bdd
    ' 3- Nom de la base
    p_nomBDD_ODBC = STR_GetChamp(v_scmd, ";", p_SCMD_NOM_BASE)
    If p_nomBDD_ODBC = "" Then
        Call MsgBox("Pas de nom de base.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
'MsgBox p_nomBDD_ODBC
    ' 5- .ini
    p_nomini = ""
    If nbprm > 4 Then
        p_nomini = STR_GetChamp(v_scmd, ";", p_SCMD_CHEMIN_INI)
    End If
    If p_nomini = "" Then
        p_nomini = "KaliRP.ini"
    End If
    p_nomini = p_chemin_appli & "\" & p_nomini
'MsgBox p_nomini
    
    '' Connexion à la base
    'If Odbc_Init(stype_bdd, nom_bdd) = P_ERREUR Then
   '     init_param_exe = P_ERREUR
   '     Exit Function
   ' End If
   ' p_nom_bdd = nom_bdd
    
    Call charger_var_ini
        
    r_direct = False
    
lab_fin:
    init_param_exe = P_OK
    
End Function
    

Private Sub charger_var_ini()

    Dim s As String, sprinc As String, sdate As String
    Dim trouve As Boolean, effacer As Boolean
    Dim lcr As Long
    
    ' Chaine d'initialisation de la connexion
    s = SYS_GetIni("BASE", "STR_INIT", p_nomini)
    If s <> "" Then
        Call Odbc_RecupVal("select kd_system('" & Replace(s, "\", "\\") & "')", lcr)
        'Call MsgBox("lcr:" & lcr)
    End If

    sdate = SYS_GetIni("LOGICIEL", "LAST_CONNEXION", p_nomini)
    If sdate = "" Then
        effacer = True
    Else
        If SAIS_CtrlChamp(sdate, SAIS_TYP_DATE) Then
            If CDate(sdate) < Date Then
                effacer = True
            Else
                effacer = False
            End If
        Else
            effacer = True
        End If
    End If
    If effacer Then
        ' Vidage du répertoire tmp
        Call FICH_EffacerFichier(p_chemin_appli + "\tmp\*.*", False)
    End If
    Call SYS_PutIni("LOGICIEL", "LAST_CONNEXION", Format(Date, "dd/mm/yyyy"), p_nomini)
    
    p_est_le_serveur = False
    p_est_le_convertisseur = False
    
lab_fin_mess:

    Exit Sub
    
lab_no_docappli:
    
End Sub

