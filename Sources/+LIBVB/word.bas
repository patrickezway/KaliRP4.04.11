Attribute VB_Name = "Mword"
Option Explicit

' Mode de travail de Word_Fusionner
Public Const WORD_IMPRESSION = 0
Public Const WORD_VISU = 1
Public Const WORD_MODIF = 2
Public Const WORD_CREATE = 3

' Ce qu'il y a à faire au début de l'appel à Word_Fusionner
Public Const WORD_DEB_CROBJ = 0
Public Const WORD_DEB_OUVDOC = 1
Public Const WORD_DEB_RIEN = 2

' Ce qu'il y a à faire à la fin de l'appel à Word_Fusionner
Public Const WORD_FIN_FERMDOC = 0
Public Const WORD_FIN_RAZOBJ = 1
Public Const WORD_FIN_RIEN = 2

Public Excel_Doc As Excel.Workbook
Public Excel_Obj As Excel.Application
Public Excel_EstActif As Boolean

Public Word_Doc As Word.Document
Public Word_Obj As Word.Application
Public Word_EstActif As Boolean

' Pour Word_CreerModele
Public Word_CrMod_stypedoc As String
Public Word_CrMod_Num As String
Public Word_CrMod_chemin As String

Public Type WORD_SSIGNET
    nom As String
    indice As Integer
End Type
Public Word_tblsignet() As WORD_SSIGNET
Public Word_nbsignet As Integer

' Pour Word_Fusionner
Private g_garder_bookmark As Boolean
Private Type W_STBLCHP
    nombk As String
    exist As Integer
End Type
' Valeurs de exist
Private Const W_AEVALUER = 0
Private Const W_NON = 1
Private Const W_OUI = 2


' Le fichier doit être en local dans tous les cas
Public Function Word_AfficherDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByVal v_fimprime As Boolean, _
                               ByVal v_fmodif As Boolean, _
                               ByVal v_voir_signets As Boolean, _
                               ByVal v_nomdata As String) As Integer

    Dim s As String, nombk As String, nomdoc As String, nomdot As String
    Dim fexist As Boolean, a_redim As Boolean
    Dim I As Integer, j As Integer, fd As Integer, n As Integer, pos As Integer
    Dim imode As Integer
    Dim s_bk As Variant
    Dim range As Word.range
    Dim g_cmsword As New CWord
    Dim NewBar As CommandBar, NewMenu As CommandBarPopup, iB As Integer
    Dim u_Nom As String
    Dim strRole As String
    Dim AncUserName As String, AncUserInitial As String
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
'    If Not v_fmodif Then
'        nomdoc = p_CheminDossierTravailLocal + "\" + p_CodeUtil + ".doc"
'        If FICH_CopierFichier(v_nomdoc, nomdoc) = P_ERREUR Then
'            Word_AfficherDoc = P_ERREUR
'            Exit Function
'        End If
'    Else
        nomdoc = v_nomdoc
'    End If
    
    If Word_ReInit() = P_ERREUR Then
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Word_Fusionner(nomdoc, _
                  "", _
                  False, _
                  v_nomdata, _
                  True, _
                  "", _
                  True, _
                  v_passwd, _
                  False, _
                  WORD_CREATE, _
                  0, _
                  WORD_DEB_CROBJ, _
                  WORD_FIN_RIEN) = P_ERREUR Then
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
                        
    ' RV393
    If p_armer_suivi Then
        'MsgBox "TrackRevisions"
        AncUserName = Word_Obj.UserName
        AncUserInitial = Word_Obj.UserInitials
        
        Word_Doc.TrackRevisions = True
        Word_Obj.UserName = p_nom_suivi
        Word_Obj.UserInitials = p_initiale_suivi
    Else
        Word_Doc.TrackRevisions = False
    End If
    p_armer_suivi = False
    
    Set g_cmsword.App = Word_Obj

    If Not v_fmodif Then
        Word_Doc.Saved = False
        Word_Doc.Save
        Word_Doc.Close
        Set Word_Doc = Nothing
        If Word_OuvrirDoc(nomdoc, _
                            Not v_fmodif, _
                            v_passwd, _
                            Word_Doc) = P_ERREUR Then
            Word_AfficherDoc = P_ERREUR
            Exit Function
        End If
    End If
    
    Set g_cmsword.doc = Word_Doc
    
    If Not v_fmodif Then
        If Not v_fimprime Then
            nomdot = p_CheminModele_Loc + "\KaliDocNoFct.dot"
            imode = 3
        Else
            nomdot = p_CheminModele_Loc + "\KaliDocImp.dot"
            imode = 2
        End If
        g_cmsword.doc.Protect wdAllowOnlyComments
    Else
        If Word_CrMod_stypedoc = "MA" Then
            If FICH_FichierExiste(p_CheminModele_Loc + "\KaliMaquette.dot") Then
                nomdot = p_CheminModele_Loc + "\KaliMaquette.dot"
            Else
                nomdot = p_CheminModele_Loc + "\KaliDoc.dot"
            End If
            imode = 1
        Else
            nomdot = p_CheminModele_Loc + "\KaliDoc.dot"
            imode = 1
        End If
    End If
    
    If g_cmsword.InitConfig(imode, nomdot, False) = P_ERREUR Then
        Call Word_Doc.Close(savechanges:=False)
        Set Word_Doc = Nothing
        Word_Obj.NormalTemplate.Saved = True
        Word_Obj.Application.Quit
        Set Word_Obj = Nothing
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Word_CrMod_stypedoc = "MA" Then
        On Error GoTo Suite_Newbar
        Set NewBar = Word_Obj.CommandBars("KaliTech")
        For iB = 1 To NewBar.Controls.Count
            'MsgBox NewBar.Controls(i).Caption
            NewBar.Controls(iB).Delete
        Next iB
        NewBar.visible = True
    
        Set NewMenu = NewBar.Controls.Add(type:=msoControlPopup, Before:=1, Temporary:=True)
        NewMenu.visible = True
        NewMenu.Caption = "Maquettes : Liste des Outils"
        NewMenu.OnAction = "Ouvrir"
Suite_Newbar:
    End If
    
    On Error GoTo lab_fin_err
    g_cmsword.App.visible = True
    g_cmsword.App.ActiveWindow = True
    g_cmsword.doc.ActiveWindow.View.type = wdPageView
    If v_voir_signets Then
        g_cmsword.doc.ActiveWindow.View.ShowBookmarks = True
    End If
    If g_cmsword.App.WindowState <> wdWindowStateMaximize Then
        g_cmsword.App.WindowState = wdWindowStateMaximize
    End If
    g_cmsword.App.Activate
    
    g_cmsword.App.NormalTemplate.Saved = True
    g_cmsword.doc.AttachedTemplate.Saved = True

    On Error GoTo lab_fin
'    While Word_Obj.visible
    While Not g_cmsword.lafin
        SYS_Sleep (500)
        DoEvents
    Wend
    
lab_fin:
    On Error Resume Next
    ' RV393
    If AncUserName <> "" Or AncUserInitial <> "" Then
        If Word_ReInit() <> P_ERREUR Then
            Word_Obj.UserName = AncUserName
            Word_Obj.UserInitials = AncUserInitial
        End If
    End If
    
    Word_Doc.TrackRevisions = False
    
    Set Word_Doc = Nothing
    Set Word_Obj = Nothing
    Set g_cmsword = Nothing
    On Error GoTo 0
    Word_EstActif = False
    
    Word_AfficherDoc = P_OK
    Exit Function

lab_fin_err:
    ' RV393
    Word_Doc.TrackRevisions = False
    
    Set Word_Obj = Nothing
    Set Word_Doc = Nothing
    Set g_cmsword.App = Nothing
    MsgBox "Erreur WORD " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Word_AfficherDoc = P_ERREUR
    Exit Function

End Function

' Le fichier doit être en local dans tous les cas
Public Function Word_AfficherModele(ByVal v_nomdoc As String, _
                                    ByVal v_fmodif As Boolean) As Integer

    Dim s As String, nomdot As String
    Dim I As Integer, j As Integer, fd As Integer, n As Integer, pos As Integer
    Dim imode As Integer
    Dim g_cmsword As New CWord
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Word_AfficherModele = P_ERREUR
        Exit Function
    End If
    
    If Word_ReInit() = P_ERREUR Then
        Word_AfficherModele = P_ERREUR
        Exit Function
    End If
    
    If Word_OuvrirDoc(v_nomdoc, Not v_fmodif, "", Word_Doc) = P_ERREUR Then
        Word_AfficherModele = P_ERREUR
        Exit Function
    End If
                        
    Set g_cmsword.App = Word_Obj
    Set g_cmsword.doc = Word_Doc
    
    If Not v_fmodif Then
        nomdot = p_CheminModele_Loc + "\KaliDocImp.dot"
        imode = 2
        g_cmsword.doc.Protect wdAllowOnlyComments
    Else
        If Word_CrMod_stypedoc = "" Then
            nomdot = p_CheminModele_Loc + "\KaliDoc.dot"
        Else
            If Not FICH_FichierExiste(p_CheminModele_Loc + "\KaliModele.dot") Then
                nomdot = p_CheminModele_Loc + "\KaliDoc.dot"
                Word_CrMod_stypedoc = ""
            Else
                nomdot = p_CheminModele_Loc + "\KaliModele.dot"
            End If
        End If
        imode = 1
    End If
    
    If g_cmsword.InitConfig(imode, nomdot, False) = P_ERREUR Then
        Call Word_Doc.Close(savechanges:=False)
        Set Word_Doc = Nothing
        Word_Obj.NormalTemplate.Saved = True
        Word_Obj.Application.Quit
        Set Word_Obj = Nothing
        Word_AfficherModele = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo lab_fin_err
    g_cmsword.App.visible = True
    g_cmsword.App.ActiveWindow = True
    g_cmsword.doc.ActiveWindow.View.type = wdPageView
    g_cmsword.doc.ActiveWindow.View.ShowBookmarks = True
    If g_cmsword.App.WindowState <> wdWindowStateMaximize Then
        g_cmsword.App.WindowState = wdWindowStateMaximize
    End If
    g_cmsword.App.Activate
    
    If v_fmodif And Word_CrMod_stypedoc <> "" Then
        Call w_generer_btnfusion
    End If
    
    g_cmsword.App.NormalTemplate.Saved = True
    g_cmsword.doc.AttachedTemplate.Saved = True
    
    On Error GoTo lab_fin
'    While g_cmsword.App.Visible
    While Not g_cmsword.lafin
        SYS_Sleep (500)
        DoEvents
    Wend
    
lab_fin:
    On Error Resume Next
    Set Word_Doc = Nothing
    Set Word_Obj = Nothing
    Set g_cmsword = Nothing
    On Error GoTo 0
    Word_EstActif = False
    
    Word_AfficherModele = P_OK
    Exit Function

lab_fin_err:
    Set Word_Obj = Nothing
    Set Word_Doc = Nothing
    Set g_cmsword.App = Nothing
    MsgBox "Erreur WORD " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Word_AfficherModele = P_ERREUR
    Exit Function

End Function

Public Function Word_ChangerPasswd(ByVal v_nomdoc As String, _
                                   ByVal v_o_passwd As String, _
                                   ByVal v_n_passwd As String) As Integer
                             
    If Word_Init() = P_ERREUR Then
        Word_ChangerPasswd = P_OK
        Exit Function
    End If
    
    If Word_OuvrirDoc(v_nomdoc, False, v_o_passwd, Word_Doc) = P_ERREUR Then
        Word_ChangerPasswd = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_word
    ' Ruse : sinon le document n'est pas enregistré ...
    Word_Doc.Saved = False
    Word_Doc.Password = v_n_passwd
    Word_Doc.Save
    Word_Doc.Close
    Set Word_Doc = Nothing
    On Error GoTo 0
    
    Word_ChangerPasswd = P_OK
    Exit Function

err_word:
    MsgBox "Erreur WORD " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Word_ChangerPasswd = P_ERREUR
    Exit Function

End Function

Public Function Word_ConvHTML(ByVal v_nomdoc As String, _
                              ByVal v_nomhtml As String, _
                              ByVal v_conv As Integer) As Integer
                         
    If Word_Init() = P_ERREUR Then
        Word_ConvHTML = P_ERREUR
        Exit Function
    End If
    
    If Word_OuvrirDoc(v_nomdoc, True, "", Word_Doc) = P_ERREUR Then
        Call Word_Quitter(WORD_FIN_RAZOBJ)
        Word_ConvHTML = P_ERREUR
        Exit Function
    End If
    
    ' On supprime les marques de fautes pour la conv des Formulaires
    On Error Resume Next
    Word_Doc.SpellingChecked = False
    Word_Doc.ShowSpellingErrors = False
    Word_Doc.GrammarChecked = False
    Word_Doc.ShowGrammaticalErrors = False
    With Word_Doc.WebOptions
        .OrganizeInFolder = True
        .RelyOnVML = False
   End With
 
    On Error GoTo err_saveas
    Call Word_Doc.SaveAs(Filename:=v_nomhtml, FileFormat:=v_conv)
    On Error GoTo 0
    
    If Word_CrMod_stypedoc = "MA" Then
        ' permet de générer Maquette.outil
        Call MAJ_Maquette_Outil
    End If
    
    Call Word_Quitter(WORD_FIN_FERMDOC)
    
    Word_ConvHTML = P_OK
    Exit Function
    
err_saveas:
    MsgBox "Erreur WORD SaveAs HTML " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Call Word_Quitter(WORD_FIN_FERMDOC)
    Word_ConvHTML = P_ERREUR
    Exit Function

End Function
                         
Public Function MAJ_Maquette_Outil()
    'MsgBox "MAJ_Maquette_Outil"
    Dim numlig As Integer, encore As Boolean, laDim As Integer
    Dim fd As Integer, ligne As String, NiveauTPL As String
    Dim TypeObjet As String, NomObjet As String, VersionObjet As String
    Dim NomTPF As String, Parser As String, ChampsOutil As String
    Dim NumeroFichier As String, nb As Integer, I As Integer, NomOutil As String
    Dim CheminMaquetteOutil As String
    Dim leMot As String
    Dim Tablo()
    Dim LaDim1 As Integer, LaDim2 As Integer
    Dim sql As String
    Dim lnb As Long
Exit Function

    ' Charger OutilMaquette.txt
'**********************************************************************************************
'  Nom         |Outil ?|Version|Niveau|Numéro |Nom TPF               | Parser   |Champs de l'outil
'  objet       |       |Include| TPL  |Fichier|à utiliser            | de suite |
'**********************************************************************************************
'O|Contenudoc   |  N    |1      |0     |       |                      |O         |
'L|Contenudoc   |  Contenu du dossier                                 |{LISTEDOS}
''
'O|Listenewdoc  |  O    |1      |0     |       |listenewdoc_I1_V1.tpf |N         |Utilnewdoc,Listenewdoc
'L|Listenewdoc  |  Les nouveaux documents
'C|Listenewdoc  |  Utilnewdoc          |a des nouveaux documents      |{UTIL_NEWDOC_SUITE}
'C|Listenewdoc  |  Listenewdoc         |liste des nouveaux documents  |{LISTENEWDOC}
    
    ' contient tous les outils de la maquette choisie
    fd = FreeFile
    CheminMaquetteOutil = p_CheminKW & "\Maquettes\Base_Maquette\OutilMaquette.txt"
    Open CheminMaquetteOutil For Input As fd
    
    On Error GoTo err_open
    
    ' lire la première ligne
    numlig = 0
    encore = True
    While encore
        On Error GoTo err_fin_fichier1
        Line Input #fd, ligne
        GoTo suite_fin_fichier
err_fin_fichier1:
        encore = False
        Resume fin_fichier
suite_fin_fichier:
        If Mid(ligne, 1, 1) <> "'" And Mid(ligne, 1, 1) <> "*" And left(ligne, 6) <> "THEME=" Then
            TypeObjet = Trim(STR_GetChamp(ligne, "|", 0))
            If TypeObjet = "O" Then
            End If
        End If
    Wend
        
        ' O     |Util              |1      |0     |       |util_I1_V1.tpf      |N         |Utilident,Utilform,Utilkalimail,Utiltache,Utilliste
         
     
'Word_Obj.Selection.WholeStory
'MsgBox Word_Obj.Selection.range.start & " " & Word_Obj.Selection.range.End
'MsgBox Word_Obj.Selection.range.Text
'MsgBox Word_Obj.Selection.range.Words.Count

MsgBox Word_Obj.Selection.range.Words(20).Text
For I = 1 To Word_Obj.Selection.range.Words.Count
    
    MsgBox Word_Obj.Selection.range.Words(I).Text
    If Word_Obj.Selection.range.Words(I).Text = "{" Then
        If Word_Obj.Selection.range.Words(I + 2).Text = "}" Then
            leMot = Word_Obj.Selection.range.Words(I + 1).Text
            MsgBox leMot
        End If
    End If
Next I
    nb = STR_GetNbchamp(Word_Obj.Selection.range.Text, "{")
I = I
    'For i = 0 To nb
    'NomOutil = STR_GetChamp(ChampsOutil, ",", i)


Exit Function
' TabOutilMaquetteMaquette()
    
    On Error GoTo err_open
    fd = FreeFile
    ' contient tous les outils de la maquette choisie
    ' Vérifie s'il y a plusieurs maquettes accueil
    sql = "Select count(*) " _
        & "from documentation " _
        & "where do_intranet='t'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
         Exit Function
    End If
    If lnb = 1 Then
        CheminMaquetteOutil = p_CheminKW & "\Maquettes\Accueil\maquette.outil"
    Else
        CheminMaquetteOutil = p_CheminKW & "\Maquettes\Accueil" & p_NumDocs & "\maquette.outil"
    End If

    
    Open CheminMaquetteOutil For Input As fd
    
    
    On Error GoTo err_open
    
    ' lire la première ligne
    numlig = 0
    encore = True
    While encore
        On Error GoTo err_fin_fichier
        Line Input #fd, ligne
        GoTo suite_fin_fichier2
err_fin_fichier:
        encore = False
        Resume fin_fichier
suite_fin_fichier2:
        ' O     |Util              |1      |0     |       |util_I1_V1.tpf      |N         |Utilident,Utilform,Utilkalimail,Utiltache,Utilliste
        If Mid(ligne, 1, 1) <> "'" And Mid(ligne, 1, 1) <> "*" And left(ligne, 6) <> "THEME=" Then
            TypeObjet = Trim(STR_GetChamp(ligne, "|", 0))
            If TypeObjet = "O" Then
                NomObjet = Trim(STR_GetChamp(ligne, "|", 1))
                VersionObjet = Trim(STR_GetChamp(ligne, "|", 2))
                NiveauTPL = Trim(STR_GetChamp(ligne, "|", 3))
                NumeroFichier = Trim(STR_GetChamp(ligne, "|", 4))
                NomTPF = Trim(STR_GetChamp(ligne, "|", 5))
                'VerifierTPF (v_Chemin & "\" & NomTPF)
                Parser = Trim(STR_GetChamp(ligne, "|", 6))
                ChampsOutil = Trim(STR_GetChamp(ligne, "|", 7))
                '
                If ChampsOutil <> "" Then
                    nb = STR_GetNbchamp(ChampsOutil, ",")
                    If nb > 0 Then
                        ' une ligne par Outil
                        For I = 0 To nb
                            NomOutil = STR_GetChamp(ChampsOutil, ",", I)
                            If NomOutil <> "" Then
                                On Error Resume Next
                                laDim = 1
                                'laDim = UBound(TabOutilMaquette(), 2) + 1
                                'ReDim Preserve TabOutilMaquette(DimOutil, laDim)
                                'TabOutilMaquette(0, laDim) = numlig
                                'TabOutilMaquette(1, laDim) = NomObjet
                                'TabOutilMaquette(2, laDim) = NomOutil
                                'TabOutilMaquette(3, laDim) = ""
                            End If
                        Next I
                    End If
                Else
                    On Error Resume Next
                    laDim = 1
                    'laDim = UBound(TabOutilMaquette(), 2) + 1
                    'ReDim Preserve TabOutilMaquette(DimOutil, laDim)
                    'TabOutilMaquette(0, laDim) = numlig
                    'TabOutilMaquette(1, laDim) = NomObjet
                    'TabOutilMaquette(2, laDim) = ""
                    'TabOutilMaquette(3, laDim) = ""
                End If
                numlig = numlig + 1
            End If
        End If
    Wend

fin_fichier:
    Close fd
    Exit Function
err_open:
    MsgBox "Impossible d'ouvrir " & CheminMaquetteOutil
    On Error GoTo 0
    Exit Function
    
End Function

                         
Public Function Word_CopierModele(ByVal v_nomdoc As String, _
                                  ByVal v_nommodele As String, _
                                  ByVal v_bcopie_entete As Boolean, _
                                  ByVal v_bcopie_corps As Boolean, _
                                  ByVal v_garder_styles As Boolean, _
                                  ByVal v_passwd As String) As Integer

    Dim doc_modele As Word.Document
    Dim arange As Word.range

    If Word_Init() = P_ERREUR Then
        Word_CopierModele = P_ERREUR
        Exit Function
    End If
    
    ' Ouvre le document d'origine
    If Word_OuvrirDoc(v_nomdoc, False, v_passwd, Word_Doc) = P_ERREUR Then
        Call Word_Quitter(WORD_FIN_RAZOBJ)
        Word_CopierModele = P_ERREUR
        Exit Function
    End If
    
    ' Ouvre le modèle
    If Word_OuvrirDoc(v_nommodele, True, "", doc_modele) = P_ERREUR Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
        Call Word_Quitter(WORD_FIN_FERMDOC)
        Word_CopierModele = P_ERREUR
        Exit Function
    End If
    
'Word_Obj.Visible = True
    ' Recopie de l'entete -> on part du modèle
    If v_bcopie_entete Then
        ' Recopie du corps
        If v_bcopie_corps Then
            Call w_suppr_bk_doublon(Word_Doc, doc_modele)
        Else
            ' Efface le corps du modèle
            Set arange = doc_modele.range
            arange.Select
            arange.Text = ""
        End If
        ' Recopie du corps du document dans le modèle
        Call w_copier_corps(v_nomdoc, Word_Doc, doc_modele, v_garder_styles)
        Call FICH_EffacerFichier(v_nomdoc, False)
        ' Le modèle devient le nouveau document
        Call doc_modele.SaveAs(Filename:=v_nomdoc, Password:=v_passwd)
        Call doc_modele.Close(savechanges:=wdDoNotSaveChanges)
        Set doc_modele = Nothing
    ' Pas d'entete -> on part du document
    Else
        ' Recopie du corps du modèle
        If v_bcopie_corps Then
            ' Recopie du corps du modèle dans le document
            Call w_copier_corps(v_nommodele, doc_modele, Word_Doc, v_garder_styles)
        Else
            Call doc_modele.Close(savechanges:=wdDoNotSaveChanges)
            Set doc_modele = Nothing
        End If
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
        Set Word_Doc = Nothing
    End If
    
    Word_CopierModele = P_OK
    
End Function

Public Function Word_CreerModele(v_form As Form, _
                                 ByVal v_nomfich_chp As String, _
                                 ByVal v_nomdoc As String, _
                                 ByVal v_plusieurs_fois_meme_chp_autor As Boolean, _
                                 ByVal v_fmodif As Boolean, _
                                 Optional v_frm As Variant)
                        
    Dim tbl_name() As String, ssys As String
    Dim nomutil As String, sdat_av As String, sdat_ap As String
    Dim ya_un_tab As Boolean, est_danstab As Boolean, inheadfoot As Boolean
    Dim tbl_inhf() As Boolean, trouve As Boolean
    Dim pos As Integer, I As Integer, notab As Integer, n As Integer
    Dim cr As Integer, ntab As Integer, lig_tab As Integer, j As Integer
    Dim siz_tab As Long, tbl_start() As Long, tbl_end() As Long
    Dim arange As Word.range, trange As Word.range

    ' On vérifie l'existance de champ.txt
    If Not FICH_FichierExiste(v_nomfich_chp) Then
        Call MsgBox("Le fichier '" & v_nomfich_chp & "' étant inaccessible, vous ne pouvez pas accéder aux modèles.", vbInformation + vbOKOnly, "")
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
lab_debut:
    ' Le modèle n'est pas trouvé
    If Not FICH_FichierExiste(v_nomdoc) Then
        Call MsgBox("Le fichier '" & v_nomdoc & "' n'a pas été trouvé.", vbInformation + vbOKOnly, "")
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    ' Lecture seulement
    If Not v_fmodif Then
        Call Word_AfficherDoc(v_nomdoc, "", True, False, True, "")
        Call FICH_EffacerFichier(v_nomdoc, False)
        Word_CreerModele = P_NON
        Exit Function
    End If
    
    ' *** Ouverture en modif ***
    
    ' Stocke la date de modif du fichier
    sdat_av = FICH_FichierDateTime(v_nomdoc)
    
    ' L'utilisateur paramètre son document
    Call Word_AfficherModele(v_nomdoc, True)
    
    sdat_ap = FICH_FichierDateTime(v_nomdoc)
    ' Pas de modification
    If sdat_ap = sdat_av Then
        Word_CreerModele = P_NON
        Exit Function
    End If
    
    Call FRM_ResizeForm(v_form, v_form.width, v_form.Height)
    If Not IsMissing(v_frm) Then
        If Not v_frm Is Nothing Then
            v_frm.visible = True
        End If
    End If
    DoEvents
    v_form.Refresh
    
    If Word_Init() = P_ERREUR Then
        If Not IsMissing(v_frm) Then
            If Not v_frm Is Nothing Then
                v_frm.visible = False
            End If
        End If
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    If Word_OuvrirDoc(v_nomdoc, False, "", Word_Doc) = P_ERREUR Then
        If Not IsMissing(v_frm) Then
            If Not v_frm Is Nothing Then
                v_frm.visible = False
            End If
        End If
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_word
'Word_Obj.Visible = True
    Word_Doc.ActiveWindow.View.ShowFieldCodes = False
    Word_Doc.ActiveWindow.View.type = wdPageView
    
    Call w_init_tblsignet(True)
    
    ' En-tête
    For I = 1 To Word_Doc.Sections(1).Headers.Count
        Set trange = Word_Doc.Sections(1).Headers(I).range
        Call w_conv_champ_en_signet(trange)
    Next I
    ' Corps
    Word_Doc.ActiveWindow.View.type = wdPageView
    Word_Doc.ActiveWindow.View.SeekView = wdSeekMainDocument
    Set trange = Word_Doc.StoryRanges(wdMainTextStory)
    Call w_conv_champ_en_signet(trange)
' On laisse de coté pour l'instant
'    For i = 1 To Word_Doc.Shapes.Count
'        If Word_Doc.Shapes(i).TextFrame.HasText Then
'            Set trange = Word_Doc.Shapes(i).TextFrame.TextRange
'            Call w_conv_champ_en_signet(trange)
'        End If
'    Next i
    ' Pied
    For I = 1 To Word_Doc.Sections(1).Footers.Count
        Set trange = Word_Doc.Sections(1).Footers(I).range
        Call w_conv_champ_en_signet(trange)
    Next I

    ' On supprime les bookmark tableau glob
    I = 1
    While I <= Word_Doc.Bookmarks.Count
        If w_est_champ_tableau_global(Word_Doc.Bookmarks(I).Name) Then
            Word_Doc.Bookmarks(I).Delete
            I = I - 1
        End If
        I = I + 1
    Wend
    
    ' Insertion des bookmark tableau glob
    siz_tab = 0
    ya_un_tab = False
    For I = 1 To Word_Doc.Bookmarks.Count
        If w_est_champ_tableau(Word_Doc.Bookmarks(I).Name) = P_OUI Then
            est_danstab = w_ajouter_bkglob(Word_Doc.Bookmarks(I).Name, inheadfoot)
            If Not est_danstab Then GoTo lab_suivant
            notab = CInt(Mid$(Word_Doc.Bookmarks(I).Name, 2, 2))
            ya_un_tab = True
            If siz_tab < notab Then GoTo lab_cr_tab
lab_book_suiv:
            If tbl_start(notab) > Word_Doc.Bookmarks(I).start Then
                tbl_start(notab) = Word_Doc.Bookmarks(I).start
            End If
            If tbl_end(notab) < Word_Doc.Bookmarks(I).End Then
                tbl_end(notab) = Word_Doc.Bookmarks(I).End
            End If
            tbl_inhf(notab) = inheadfoot
        End If
lab_suivant:
    Next I
    If ya_un_tab Then
        For I = 1 To UBound(tbl_name)
            If tbl_name(I) <> "" Then
                If tbl_inhf(I) Then
                    trouve = False
                    For j = 1 To Word_Doc.Sections(1).Headers.Count
                        Word_Doc.Sections(1).Headers(j).range.Select
                        If w_ya_bktbl_dans_sel(tbl_name(I)) Then
                            Set arange = Word_Doc.Sections(1).Headers(j).range
                            arange.start = tbl_start(I)
                            arange.End = tbl_end(I) + 1
                            trouve = True
                            Exit For
                        End If
                    Next j
                    ' Pied
                    If Not trouve Then
                        For j = 1 To Word_Doc.Sections(1).Footers.Count
                            Word_Doc.Sections(1).Footers(j).range.Select
                            If w_ya_bktbl_dans_sel(tbl_name(I)) Then
                                Set arange = Word_Obj.Selection.range
                                arange.start = tbl_start(I)
                                arange.End = tbl_end(I) + 1
                                trouve = True
                                Exit For
                            End If
                        Next j
                    End If
                Else
                    If w_rangeb(tbl_start(I), tbl_end(I) + 1, arange) = P_OK Then
                        trouve = True
                    End If
                End If
                If trouve Then
                    If w_add_bookmark(Word_Doc, tbl_name(I), arange) = P_ERREUR Then GoTo lab_fin_err
                End If
            End If
        Next I
    End If
    
    Call w_init_tblsignet(v_plusieurs_fois_meme_chp_autor)
    
    GoTo lab_fin_ok
    
lab_cr_tab:
    siz_tab = notab
    ReDim Preserve tbl_name(notab) As String
    pos = InStr(Mid$(Word_Doc.Bookmarks(I).Name, 5), "_")
    tbl_name(notab) = left$(Word_Doc.Bookmarks(I).Name, pos + 3)
    ReDim Preserve tbl_start(notab) As Long
    tbl_start(notab) = 9999
    ReDim Preserve tbl_end(notab) As Long
    tbl_end(notab) = 0
    ReDim Preserve tbl_inhf(notab) As Boolean
    GoTo lab_book_suiv
    
err_word:
    MsgBox "Erreur WORD " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    GoTo lab_fin_err

lab_fin_err:
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    Set Word_Doc = Nothing
    If Not IsMissing(v_frm) Then
        If Not v_frm Is Nothing Then
            v_frm.visible = False
        End If
    End If
    Word_CreerModele = P_ERREUR
    Exit Function
    
lab_fin_ok:
    Call Word_Doc.Close(savechanges:=wdSaveChanges)
    Set Word_Doc = Nothing
    If Not IsMissing(v_frm) Then
        If Not v_frm Is Nothing Then
            v_frm.visible = False
        End If
    End If
    Word_CreerModele = P_OUI
    Exit Function
    
End Function

Public Function Word_Fusionner(ByVal v_nommod As String, _
                           ByVal v_nominit As String, _
                           ByVal v_garder_styles As Boolean, _
                           ByVal v_nomdata As String, _
                           ByVal v_garder_bookmark As Boolean, _
                           ByVal v_nomdest As String, _
                           ByVal v_ecraser As Boolean, _
                           ByVal v_passwd As String, _
                           ByVal v_word_visible As Boolean, _
                           ByVal v_word_mode As Integer, _
                           ByVal v_nbex As Integer, _
                           ByVal v_deb_mode As Integer, _
                           ByVal v_fin_mode As Integer) As Integer
    
    Dim s As String, mess_err As String
    Dim str_entete As String, nomchp As String, str_data As String
    Dim chptab As String, nomtab As String, chp As String
    Dim sval As String, nombk As String
    Dim encore As Boolean, again As Boolean
    Dim ya_book_glob As Boolean, a_redim As Boolean, b_fairefusion As Boolean
    Dim frempl As Boolean, encore_bk As Boolean
    Dim fd As Integer, poschp As Integer, pos As Integer, pos2 As Integer
    Dim I As Integer, j As Integer, ntab As Integer, nlig As Integer, n As Integer
    Dim lig_tab As Integer, col As Integer, ind As Integer
    Dim idebtab As Long, ifintab As Long
    Dim tbl_chp() As W_STBLCHP
    Dim arange As Word.range, arange2 As Word.range
    Dim doc2 As Word.Document
    Dim dochf As Variant
    
    If v_deb_mode = WORD_DEB_CROBJ Then
        If Word_Init() = P_ERREUR Then
            Word_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    If v_deb_mode <> WORD_DEB_RIEN Then
        If Word_OuvrirDoc(v_nommod, False, v_passwd, Word_Doc) = P_ERREUR Then
            Word_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    ' pas de traces de modifications pour les fusions de données
    Word_Doc.TrackRevisions = False
    
    g_garder_bookmark = v_garder_bookmark
    
    b_fairefusion = True
    If v_nomdata = "" Then
        b_fairefusion = False
        GoTo lab_fin_fusion
    End If
    
'v_word_visible = True
    If v_word_visible Then
        Word_Obj.visible = True
        Word_Obj.ActiveWindow = True
        a_redim = False
        On Error Resume Next
        If Word_Obj.WindowState <> wdWindowStateMaximize Then
            a_redim = True
        End If
        Word_Obj.Activate
        If a_redim Then
            Word_Obj.WindowState = wdWindowStateMaximize
        End If
        On Error GoTo 0
    Else
        Word_Obj.visible = False
    End If
    
    If v_nomdest <> "" Then
        On Error GoTo err_sav_dest
        If v_ecraser Then
            If v_passwd <> "" Then
                Word_Doc.Password = v_passwd
            End If
            Call FICH_EffacerFichier(v_nomdest, False)
            Word_Doc.SaveAs Filename:=v_nomdest
            On Error GoTo 0
        Else
            If v_deb_mode <> WORD_DEB_RIEN Then
                If Word_OuvrirDoc(v_nomdest, False, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                While doc2.Bookmarks.Count > 0
                    doc2.Bookmarks(1).Delete
                Wend
                Set arange = Word_Doc.range
                arange.Copy
                Set arange2 = doc2.Content
                arange2.Collapse wdCollapseEnd
                On Error GoTo lab_err_paste
                arange2.Paste
                On Error GoTo 0
                Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
                Set Word_Doc = Word_Obj.ActiveDocument
            Else
                While Word_Doc.Bookmarks.Count > 0
                    Word_Doc.Bookmarks(1).Delete
                Wend
                If Word_OuvrirDoc(v_nommod, False, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                Set arange = doc2.range
                arange.Copy
                Set arange2 = Word_Doc.Content
                arange2.Collapse wdCollapseEnd
                On Error GoTo lab_err_paste
                arange2.Paste
                On Error GoTo 0
                Call doc2.Close(savechanges:=wdDoNotSaveChanges)
                Set doc2 = Nothing
            End If
        End If
    End If
    
    On Error GoTo err_word1
    
    Word_Doc.MailMerge.MainDocumentType = wdNotAMergeDocument
    
    If Word_Doc.Bookmarks.Count < 1 Then
        GoTo lab_copie_fich
    End If

    On Error GoTo err_open_fus
    fd = FreeFile
    Open v_nomdata For Input As #fd
    On Error GoTo err_word2

    ' Ligne d'entete
    Line Input #fd, str_entete
    
    encore = True
    Do While encore
        If str_entete = "" Then
            GoTo lab_copie_fich
        End If
        poschp = InStr(str_entete, ";")
        If poschp = 0 Then
            ' fini
            nomchp = str_entete
            encore = False
        Else
            nomchp = left(str_entete, poschp - 1)
            str_entete = Right(str_entete, Len(str_entete) - poschp)
        End If
        pos = InStr(nomchp, "#")
        If pos > 0 Then
            ' c'est un tableau
            chptab = Mid$(nomchp, pos + 1)
            pos2 = InStr(pos + 1, chptab, "#")
            chptab = Mid$(chptab, pos2 + 1)
            nomtab = Mid$(nomchp, 2, pos2 - 1) & "_1"
            ya_book_glob = Word_Doc.Bookmarks.Exists(nomtab)
            ' Chargement du nom des champs
            again = True
            I = 0
            Do While again
                pos = InStr(chptab, "|")
                If pos > 0 Then
                    chp = left(chptab, pos - 1)
                    chptab = Right(chptab, Len(chptab) - pos)
                Else
                    chp = chptab
                    again = False
                End If
                I = I + 1
                ReDim Preserve tbl_chp(I) As W_STBLCHP
                tbl_chp(I).nombk = chp
                tbl_chp(I).exist = W_AEVALUER
            Loop
            ' Détermine s'il y a un tableau word
            If ya_book_glob Then
                lig_tab = -1
                Call w_bk_dans_tableau(nomtab, dochf, ntab, lig_tab)
                If lig_tab > 0 Then
                    ' Efface toutes les lignes du tableau sauf la 1e
                    For I = lig_tab + 1 To dochf.Tables(ntab).Rows.Count
                        dochf.Tables(ntab).Rows(lig_tab + 1).Delete
                    Next I
                    dochf.Tables(ntab).Rows(lig_tab).Select
                Else
                    ' Sauvegarde position bookmark tableau glob
                    idebtab = Word_Doc.Bookmarks(nomtab).start
                    Word_Doc.Bookmarks(nomtab).Select
                End If
                Word_Obj.Selection.Copy
            End If
            ' Traitement des lignes
            again = True
            frempl = True
            Do While again
                If w_lire_fich(fd, str_data) = P_NON Then Exit Do
                If Right$(str_data, 1) = ";" Then
                    again = False
                End If
                If Not frempl Then
                    GoTo lab_lig_suiv
                End If
                For col = 1 To UBound(tbl_chp())
                    pos = InStr(str_data, "|")
                    If pos > 0 Then
                        sval = left(str_data, pos - 1)
                        str_data = Right(str_data, Len(str_data) - pos)
                    Else
                        If Not again Then
                            sval = left$(str_data, Len(str_data) - 2)
                        Else
                            sval = left$(str_data, Len(str_data) - 1)
                        End If
                    End If
                    If sval = "" Then
                        GoTo lab_chp_suiv
                    End If
                    sval = STR_Remplacer(sval, "##", vbcr)
                    ' pour chaque donnée
                    If tbl_chp(col).exist = W_NON Then
                        GoTo lab_chp_suiv
                    End If
                    ind = 1
                    Do
                        encore_bk = False
                        nombk = left$(nomtab, Len(nomtab) - 2) & "_" & tbl_chp(col).nombk & "_" & ind
                        If tbl_chp(col).exist = W_AEVALUER Then
                            If Word_Doc.Bookmarks.Exists(nombk) = False Then
                                tbl_chp(col).exist = W_NON
                                GoTo lab_chp_suiv
                            Else
                                tbl_chp(col).exist = W_OUI
                            End If
                        End If
                        If tbl_chp(col).exist = W_OUI And Word_Doc.Bookmarks.Exists(nombk) Then
                            If w_put_txtbk(sval, nombk) = P_ERREUR Then
                                GoTo lab_fin_err
                            End If
                            If again And g_garder_bookmark And ya_book_glob Then
                                Word_Doc.Bookmarks(nombk).Delete
                            End If
                            encore_bk = True
                            ind = ind + 1
                        End If
                    Loop Until encore_bk = False
lab_chp_suiv:
                Next col
lab_lig_suiv:
                If Not ya_book_glob Then
                    frempl = False
                ElseIf again Then
                    If Word_Doc.Bookmarks.Exists(nomtab) Then
                        Word_Doc.Bookmarks(nomtab).Delete
                    End If
                    ' Recopie du bookmark tableau à la ligne précédente
                    If lig_tab <> -1 Then
                        dochf.Tables(ntab).Rows(lig_tab).Select
                        Word_Obj.Selection.Paste
                    Else
                        If w_range(idebtab, idebtab, arange) = P_ERREUR Then
                            GoTo lab_fin_err
                        End If
                        arange.InsertBefore vbcr
                        If w_range(idebtab, idebtab, arange) = P_ERREUR Then
                            GoTo lab_fin_err
                        End If
                        arange.Paste
                    End If
                End If
            Loop
        Else
            ' Récupère les données
            If w_lire_fich(fd, str_data) = P_NON Then
                mess_err = "Manque au moins un champs dans les datas"
                GoTo lab_fin_err
            End If
            If str_data <> "" Then
                ' Remplace le bookmark
                ind = 1
                Do
                    encore_bk = False
                    nombk = nomchp & "_" & ind
                    If Word_Doc.Bookmarks.Exists(nombk) = True Then
                        If ind = 1 Then
                            str_data = left$(str_data, Len(str_data) - 1)
                            str_data = Replace(str_data, "|", vbcr)
                            str_data = STR_Remplacer(str_data, "##", vbcr)
                        End If
                        If w_put_txtbk(str_data, nombk) = P_ERREUR Then
                            GoTo lab_fin_err
                        End If
                        encore_bk = True
                        ind = ind + 1
                    End If
                Loop Until encore_bk = False
            End If
        End If
    Loop

    ' Rapatriement du document à recopier
lab_copie_fich:
    If v_nominit <> "" Then
        If Word_OuvrirDoc(v_nominit, True, "", doc2) = P_ERREUR Then
            GoTo lab_fin_err2
        End If
        Call w_copier_corps(v_nominit, doc2, Word_Doc, v_garder_styles)
    End If
    
lab_fin_fusion:
'    If Not b_fairefusion Then
'        Word_Obj.ActivePrinter = Printer.DeviceName
'        Word_Obj.ActiveDocument.PrintOut Background:=False, Copies:=v_nbex
'        GoTo lab_fin
'    End If
    
    If v_word_mode = WORD_IMPRESSION Then
        ' Fax ?
        If w_lire_fich(fd, str_data) = P_OUI Then
            Close #fd
            On Error GoTo err_fax
            Word_Doc.SendFax left$(str_data, Len(str_data) - 1)
            On Error GoTo err_word2
        Else
            Close #fd
            'Word_Obj.ActivePrinter = Printer.DeviceName
            Word_Obj.WordBasic.FilePrintSetup Printer:=Printer.DeviceName, DoNotSetAsSysDefault:=1
            If Word_Doc.Bookmarks.Exists("ImpPaysage") = True Then
                If w_put_txtbk("ImpPaysage", "") = P_ERREUR Then
                    GoTo lab_fin_err
                End If
                Word_Doc.PageSetup.Orientation = wdOrientLandscape
            End If
            Word_Doc.PrintOut Background:=False, Copies:=v_nbex
        End If
    ElseIf v_word_mode = WORD_VISU Then
        Close #fd
        sval = Word_Doc.FullName
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
        Set Word_Doc = Nothing
        Word_Obj.Documents.Open Filename:=sval, readonly:=True, passworddocument:=v_passwd
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_MODIF Then
        Close #fd
        Word_Doc.Saved = True
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_CREATE Then
        Close #fd
        GoTo lab_fin_create
    End If
    
lab_fin:
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
        Set Word_Doc = Nothing
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
'
    End If
    Call FICH_EffacerFichier(v_nomdata, False)
    Word_Fusionner = P_OK
    Exit Function

lab_err_paste:
    MsgBox "Erreur Paste" & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Resume Next

lab_fin_create:
    Call FICH_EffacerFichier(v_nomdata, False)
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
        Set Word_Doc = Nothing
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
        '
    End If
    Word_Obj.visible = False
    Word_Fusionner = P_OK
    Exit Function

lab_fin_visible:
    Call FICH_EffacerFichier(v_nomdata, False)
    Word_Obj.visible = True
    Word_EstActif = False
    Word_Fusionner = P_OK
    Exit Function
    
lab_fin_err2:
    Close #fd
    Call FICH_EffacerFichier(v_nomdata, False)
lab_fin_err1:
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    Set Word_Doc = Nothing
    Word_Obj.visible = False
    Word_Fusionner = P_ERREUR
    Exit Function

err_sav_dest:
    MsgBox "Impossible de sauvegarder le fichier dans " & v_nomdest & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1

err_word1:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
'Resume Next
    GoTo lab_fin_err1
    
err_open_fus:
    MsgBox "Impossible d'ouvrir le fichier de données " & v_nomdata & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1
    
err_word2:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
Resume Next
    GoTo lab_fin_err2
    
lab_fin_err:
    Call MsgBox("Erreur détectée au cours de la fusion : " & mess_err, vbOKOnly, "Fusion")
    GoTo lab_fin_err2

err_fax:
    MsgBox "Impossible d'effectuer l'envoi par fax.", vbInformation + vbOKOnly, "Fusion"
    GoTo lab_fin_err2
    
End Function

Public Sub Word_Imprimer(ByVal v_nomdoc As String, _
                         ByVal v_passwd As String, _
                         ByVal v_nbex As Integer, _
                         ByVal v_deb_mode As Integer)
        
    If v_deb_mode = WORD_DEB_CROBJ Then
        If Word_Init() = P_ERREUR Then
            Exit Sub
        End If
    End If
    
    If v_deb_mode <> WORD_DEB_RIEN Then
        If Word_OuvrirDoc(v_nomdoc, False, v_passwd, Word_Doc) = P_ERREUR Then
            Exit Sub
        End If
    End If
    
    'Word_Obj.ActivePrinter = Printer.DeviceName
    Word_Obj.WordBasic.FilePrintSetup Printer:=Printer.DeviceName, DoNotSetAsSysDefault:=1
    Word_Obj.ActiveDocument.PrintOut Background:=False, Copies:=v_nbex
    
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    Set Word_Doc = Nothing
    
End Sub

Public Function Word_Init()

    If Word_EstActif Then
        On Error GoTo lab_plus_actif
        If Word_Obj.visible Then
            '''
        End If
        On Error GoTo 0
    End If
    
    If Not Word_EstActif Then
        Word_EstActif = True
        On Error GoTo err_create_obj
        Set Word_Obj = CreateObject("word.application")
        On Error GoTo 0
    End If
    
    Word_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet WORD." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Word_Init = P_ERREUR
    Exit Function

lab_plus_actif:
    Word_EstActif = False
    Resume Next
    
End Function

Public Function Word_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_readonly As Boolean, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Word.Document) As Integer

    On Error GoTo err_open_ficr
'    Set r_doc = Word_Obj.Documents.Open(FileName:=v_nomdoc, _
'                                        ReadOnly:=v_readonly, _
'                                        passworddocument:=v_passwd, _
'                                        addtorecentfiles:=False)
    Set r_doc = Word_Obj.Documents.Open(v_nomdoc, , v_readonly, False, v_passwd)
    On Error GoTo 0
    
    Word_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Word_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function

Public Sub Word_Quitter(ByVal v_mode As Integer)

    Word_EstActif = False
    
    On Error GoTo err_quit
    
    Clipboard.Clear
    Word_Obj.NormalTemplate.Saved = True 'inutile ici ?
    If v_mode = WORD_FIN_FERMDOC Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
        Set Word_Doc = Nothing
    End If
    Word_Obj.NormalTemplate.Saved = True
    Word_Obj.Application.Quit
    Set Word_Obj = Nothing
    
    On Error GoTo 0
    Exit Sub

err_quit:
    Exit Sub
    
End Sub

Public Function Word_ReInit()

    If Word_EstActif Then
        On Error GoTo lab_create
        If Word_Obj.visible Then
            '''
        End If
        Clipboard.Clear
        Word_Obj.NormalTemplate.Saved = True
        Word_Obj.Application.Quit
        Set Word_Obj = Nothing
        On Error GoTo 0
    End If
    
lab_create:
    On Error GoTo err_create_obj
    Set Word_Obj = CreateObject("word.application")
    On Error GoTo 0
    Word_EstActif = True
    
    Word_ReInit = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet WORD." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Word_ReInit = P_ERREUR
    Exit Function

End Function

Private Function w_add_bookmark(ByRef v_doc As Word.Document, _
                                ByVal v_nom As String, _
                                ByVal v_arange As Word.range) As Integer

    On Error GoTo err_add_book
    v_doc.Bookmarks.Add Name:=v_nom, range:=v_arange
    On Error GoTo 0
    w_add_bookmark = P_OK
    Exit Function
    
err_add_book:
    On Error GoTo 0
    Call MsgBox("Erreur w_add_bookmark " & v_nom & vbcr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    w_add_bookmark = P_ERREUR
    Exit Function
    
End Function

Private Function w_ajouter_bkglob(ByVal v_nombk As String, _
                                  ByRef r_inheadfoot As Boolean) As Boolean

    Dim I As Integer, j As Integer, ntab As Integer
    
    Word_Doc.Bookmarks(v_nombk).Select
    ' Le bookmark n'est pas dans un tableau
    If Not Word_Obj.Selection.Information(wdWithInTable) Then
        w_ajouter_bkglob = False
        Exit Function
    End If
    
    r_inheadfoot = Word_Obj.Selection.Information(wdInHeaderFooter)
    Word_Obj.Selection.Tables(1).Select
    For j = 1 To Word_Obj.Selection.Bookmarks.Count
        If w_est_champ_tableau(Word_Obj.Selection.Bookmarks(j).Name) Then
            If left$(Word_Obj.Selection.Bookmarks(j).Name, 3) <> left$(v_nombk, 3) Then
                w_ajouter_bkglob = False
                Exit Function
            End If
        End If
    Next j
    
    w_ajouter_bkglob = True
    Exit Function
                
End Function

Private Sub w_bk_dans_tableau(ByVal v_nombk As String, _
                              ByRef r_dochf As Variant, _
                              ByRef r_itab As Integer, _
                              ByRef r_lig As Integer)

    Dim I As Integer, j As Integer, itab As Integer, lig As Integer, k As Integer
    Dim atable As Word.Table
    
    r_lig = -1
    
    Word_Doc.Bookmarks(v_nombk).Select
    If Not Word_Obj.Selection.Information(wdWithInTable) Then
        Exit Sub
    End If

    ' On cherche dans le corps du document
    For itab = 1 To Word_Doc.Tables.Count
        Word_Doc.Tables(itab).Select
        For j = 1 To Word_Obj.Selection.Bookmarks.Count
            If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                Set r_dochf = Word_Doc
                r_itab = itab
                For lig = 1 To Word_Doc.Tables(itab).Rows.Count
                    Word_Doc.Tables(itab).Rows(lig).Select
                    For k = 1 To Word_Obj.Selection.Bookmarks.Count
                        If Word_Obj.Selection.Bookmarks(k).Name = v_nombk Then
                            r_lig = lig
                            Exit Sub
                        End If
                    Next k
                Next lig
' Ceci remplaçait la boucle for lig - mais ne fct pas en 2007
'                Word_Doc.Bookmarks(v_nombk).Select
'                r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
            End If
        Next j
    Next itab
    
    ' On cherche dans entete et pied
    Word_Doc.Bookmarks(v_nombk).Select
    If Not Word_Obj.Selection.Information(wdInHeaderFooter) Then
        Exit Sub
    End If
    
    ' Entete
    For I = 1 To Word_Doc.Sections(1).Headers.Count
        Word_Doc.Sections(1).Headers(I).range.Select
        For itab = 1 To Word_Obj.Selection.Tables.Count
            Set atable = Word_Obj.Selection.Tables(itab)
            atable.Select
            For j = 1 To Word_Obj.Selection.Bookmarks.Count
                If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                    r_dochf = Word_Doc.Sections(1).Headers(I)
                    r_itab = itab
                    For lig = 1 To atable.Rows.Count
                        atable.Rows(lig).Select
                        For k = 1 To Word_Obj.Selection.Bookmarks.Count
                            If Word_Obj.Selection.Bookmarks(k).Name = v_nombk Then
                                r_lig = lig
                                Exit Sub
                            End If
                        Next k
                    Next lig
' Ceci remplaçait la boucle for lig - mais ne fct pas en 2007
'                    Word_Obj.Selection.Bookmarks(j).Select
'                    r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
                    Exit Sub
                End If
            Next j
            Word_Doc.Sections(1).Headers(I).range.Select
        Next itab
    Next I
    
    ' Pied de page
    For I = 1 To Word_Doc.Sections(1).Footers.Count
        Word_Doc.Sections(1).Footers(I).range.Select
        For itab = 1 To Word_Obj.Selection.Tables.Count
            Set atable = Word_Obj.Selection.Tables(itab)
            atable.Select
            For j = 1 To Word_Obj.Selection.Bookmarks.Count
                If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                    r_dochf = Word_Doc.Sections(1).Footers(I)
                    r_itab = itab
                    For lig = 1 To atable.Rows.Count
                        atable.Rows(lig).Select
                        For k = 1 To Word_Obj.Selection.Bookmarks.Count
                            If Word_Obj.Selection.Bookmarks(k).Name = v_nombk Then
                                r_lig = lig
                                Exit Sub
                            End If
                        Next k
                    Next lig
' Ceci remplaçait la boucle for lig - mais ne fct pas en 2007
'                    Word_Obj.Selection.Bookmarks(j).Select
'                    r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
                    Exit Sub
                End If
            Next j
            Word_Doc.Sections(1).Footers(I).range.Select
        Next itab
    Next I
    
End Sub

Private Sub w_conv_champ_en_signet(ByVal v_range As Word.range)

    Dim slibchp As String, champ As String, nombk As String
    Dim trouve As Boolean
    Dim nbfields As Integer, ifield As Integer, ibk As Integer, isignet As Integer
    Dim I As Integer, j As Integer, isig As Integer, pos As Integer
    Dim arange As Word.range
    
    nbfields = v_range.Fields.Count
    ifield = 1
    For I = 1 To nbfields
        champ = v_range.Fields(ifield).code
        If InStr(champ, "CHAMPFUSION") > 0 Then
            slibchp = "CHAMPFUSION"
        ElseIf InStr(champ, "MERGEFIELD") > 0 Then
            slibchp = "MERGEFIELD"
        Else
            ifield = ifield + 1
            GoTo lab_chpe_suiv
        End If
        champ = Mid$(champ, Len(slibchp) + 3, Len(champ) - Len(slibchp) - 3)
        champ = Replace(champ, """", "")
        ' Champ tableau sous la forme Txx_Nomtableau_NomChamp
        If w_est_champ_tableau(champ) = P_OUI Then
            ' Supprime Txx_ du champ
            'champ = Right(champ, Len(champ) - 4)
        End If
        trouve = False
        For isig = 1 To Word_nbsignet
            If UCase(Word_tblsignet(isig).nom) = UCase(champ) Then
                ibk = Word_tblsignet(isig).indice
                trouve = True
            ElseIf trouve Then
                Exit For
            End If
        Next isig
        Word_nbsignet = Word_nbsignet + 1
        ReDim Preserve Word_tblsignet(1 To Word_nbsignet) As WORD_SSIGNET
        If Not trouve Then
            ibk = 1
            pos = Word_nbsignet
        Else
            ibk = ibk + 1
            pos = isig
            For j = Word_nbsignet To pos + 1 Step -1
                Word_tblsignet(j) = Word_tblsignet(j - 1)
            Next j
        End If
        Word_tblsignet(pos).nom = champ
        Word_tblsignet(pos).indice = ibk
        nombk = champ & "_" & ibk
        v_range.Fields(ifield).Select
        Set arange = Word_Obj.Selection.range
        If w_add_bookmark(Word_Doc, nombk, arange) = P_ERREUR Then
            Exit Sub
        End If
        Word_Obj.Selection.range.InsertBefore champ
        v_range.Fields(ifield).Delete
lab_chpe_suiv:
    Next I
    
End Sub

Private Sub w_copier_corps(ByVal v_nomdocsrc As String, _
                           ByRef v_docsrc As Word.Document, _
                           ByRef v_docdest As Word.Document, _
                           ByVal v_garder_styles As Boolean)

    Dim nomdoc_src As String, nomdot As String
    Dim garder_style As Boolean
    Dim nsect As Integer, nbsect As Integer, mep As Integer, mep_crt As Integer, incr As Integer
    Dim arange As Word.range
    Dim head_foot As Word.HeaderFooter
    Dim doc_src As Word.Document
    Dim section As Word.section
    
    ' !!
    ' 0 : Portrait / 1 : Paysage
    If v_docdest.Sections(1).PageSetup.Orientation = wdOrientPortrait Then
        mep_crt = 0
    Else
        mep_crt = 1
    End If
    incr = 1
    If v_garder_styles Then
        ' Enregistre le doc source en tq que modèle (.dot)
        nomdot = p_CheminDossierTravailLocal & "\" & p_CodeUtil & Format(Time, "hhmmss") & ".dot"
        On Error GoTo err_word
        Call v_docsrc.SaveAs(Filename:=nomdot, FileFormat:=wdFormatTemplate)
        Call v_docsrc.Close
        Set v_docsrc = Nothing
        On Error GoTo 0
        ' Réouvre le doc source pour travailler sur un .doc
        ' SI TRUE -> ERREUR
        If Word_OuvrirDoc(v_nomdocsrc, False, "", v_docsrc) = P_ERREUR Then
            Exit Sub
        End If
        ' Attache le doc dest avec le .dot en mettant à jour les styles
        On Error Resume Next
        v_docdest.UpdateStylesOnOpen = True
        On Error GoTo 0
        On Error Resume Next
        v_docdest.AttachedTemplate = nomdot
        On Error GoTo 0
    End If
    ' !!
    
    On Error GoTo err_word
    nomdoc_src = p_CheminDossierTravailLocal & "\" & p_CodeUtil & Format(Time, "hhmmss") & "A.doc"
    nbsect = v_docsrc.Sections.Count
    nsect = 1
    If nbsect > 1 Then
        Call v_docsrc.SaveAs(Filename:=nomdoc_src, Password:="")
    Else
        Set doc_src = v_docsrc
        Set v_docsrc = Nothing
    End If
lab_deb_sect:
    If nbsect > 1 Then
        Set doc_src = Word_Obj.Documents.Open(nomdoc_src, readonly:=True)
        Word_Obj.visible = False
    End If
    ' mettre mode page si ce n'est pas deja le cas
    If doc_src.ActiveWindow.ActivePane.View.type = wdNormalView Or _
       doc_src.ActiveWindow.ActivePane.View.type = wdOutlineView Or _
       doc_src.ActiveWindow.ActivePane.View.type = wdMasterView Then
            doc_src.ActiveWindow.ActivePane.View.type = wdPageView
    End If
        
    ' On ne garde que la section à traiter
    If nbsect > 1 Then
        ' Suppression des sections précédentes
        If nsect > 1 Then
            Set arange = doc_src.Sections(1).range
            If nsect > 2 Then
                arange.MoveEnd unit:=wdSection, Count:=nsect - 2
            End If
            arange.Select
            On Error Resume Next
            Word_Obj.Selection.Delete
            On Error GoTo 0
        End If
        ' Suppression des sections suivantes
        If nsect < nbsect Then
            Set arange = doc_src.Sections(1).range
            arange.Collapse wdCollapseEnd
            arange.MoveEnd unit:=wdCharacter, Count:=-1
            arange.Select
            Word_Obj.Selection.MoveEnd unit:=wdStory, Count:=1
            On Error Resume Next
            Word_Obj.Selection.Delete
            On Error GoTo 0
        End If
    End If
    
    On Error Resume Next
    For Each head_foot In doc_src.Sections(1).Headers
        If Not head_foot Is Nothing Then
            While head_foot.range.Bookmarks.Count > 0
                head_foot.range.Bookmarks(1).Delete
            Wend
            head_foot.range.Delete
        End If
    Next head_foot
    For Each head_foot In doc_src.Sections(1).Footers
        If Not head_foot Is Nothing Then
            While head_foot.range.Bookmarks.Count > 0
                head_foot.range.Bookmarks(1).Delete
            Wend
            head_foot.range.Delete
        End If
    Next head_foot
    doc_src.PageSetup.DifferentFirstPageHeaderFooter = False
'    doc_src.Sections(nsect).Headers(wdHeaderFooterFirstPage).LinkToPrevious = True
'    doc_src.Sections(nsect).Footers(wdHeaderFooterFirstPage).LinkToPrevious = True
    On Error GoTo 0
    
    ' Quand la section commence directement par un tableau cela plante ...
    mep = 0
    On Error GoTo err_orientation
    If doc_src.Sections(1).PageSetup.Orientation = wdOrientLandscape Then
        mep = 1
    Else
        mep = 0
    End If
err_orientation:
    On Error GoTo 0
    
    ' Se positionne à la fin du document dest
    Set arange = v_docdest.Content
    arange.Collapse wdCollapseEnd
    If nsect > 1 Then
        If mep <> mep_crt Then
            arange.Sections.Add
        Else
            arange.InsertAfter vbFormFeed
        End If
        arange.Collapse wdCollapseEnd
    End If
    ' Sélectionne tout le corps du document à recopier et "colle"
    On Error GoTo err_word
    Set arange = v_docdest.Content
    arange.Collapse wdCollapseEnd
    doc_src.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    doc_src.StoryRanges(wdMainTextStory).Copy
    Call doc_src.Close(savechanges:=wdDoNotSaveChanges)
    Set doc_src = Nothing
    arange.Paste
    ' Changement de mise en page
    If mep <> mep_crt Then
        v_docdest.Sections(v_docdest.Sections.Count).PageSetup.Orientation = IIf(mep > 0, wdOrientLandscape, wdOrientPortrait)
        ' Le document "dest" est paramétré avec une entete 1e page différente
        If nsect > 1 And v_docdest.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True Then
            ' Incrémente éventuellement les bk de l'entete 1e page -> arange
            Set arange = v_docdest.Sections(1).Headers(1).range
            Call w_incrementer_bk(v_docdest, arange, incr)
            ' Remplace l'entete 1 de la section par arange
            v_docdest.Sections(v_docdest.Sections.Count).PageSetup.DifferentFirstPageHeaderFooter = True
            v_docdest.Sections(v_docdest.Sections.Count).Headers(wdHeaderFooterFirstPage).LinkToPrevious = False
            v_docdest.Sections(v_docdest.Sections.Count).Headers(wdHeaderFooterPrimary).LinkToPrevious = True
            Set arange = v_docdest.Sections(v_docdest.Sections.Count).Headers(wdHeaderFooterFirstPage).range
            arange.Select
            arange.Paste
            Set arange = v_docdest.Sections(v_docdest.Sections.Count).Headers(wdHeaderFooterFirstPage).range
            arange.Collapse wdCollapseEnd
            arange.Select
            Word_Obj.Selection.TypeBackspace
            ' Incrémente éventuellement les bk du bas 1e page -> arange
            Set arange = v_docdest.Sections(1).Footers(1).range
            Call w_incrementer_bk(v_docdest, arange, incr)
            ' Remplace le bas 1 de la section par arange
            v_docdest.Sections(v_docdest.Sections.Count).Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
            v_docdest.Sections(v_docdest.Sections.Count).Footers(wdHeaderFooterPrimary).LinkToPrevious = True
            Set arange = v_docdest.Sections(v_docdest.Sections.Count).Footers(wdHeaderFooterFirstPage).range
            arange.Select
            arange.Paste
            Set arange = v_docdest.Sections(v_docdest.Sections.Count).Footers(wdHeaderFooterFirstPage).range
            arange.Collapse wdCollapseEnd
            arange.Select
            Word_Obj.Selection.TypeBackspace
            incr = incr + 1
        End If
        mep_crt = mep
    End If
    If nsect < nbsect Then
        nsect = nsect + 1
        GoTo lab_deb_sect
    End If
    
    If nbsect > 1 Then
        Call FICH_EffacerFichier(nomdoc_src, False)
    End If
    
lab_fin:
    ' !!
    If v_garder_styles Then
        ' Attache le document au template de départ
        v_docdest.UpdateStylesOnOpen = False
        v_docdest.AttachedTemplate = p_CheminModele_Loc & "\kalidoc.dot"
        Call FICH_EffacerFichier(nomdot, False)
    End If
    ' !!

    On Error GoTo 0
    Exit Sub
    
err_word:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "CopierModele"
    Exit Sub
    
End Sub

Private Function w_est_champ_tableau(ByVal v_champ As String) As Integer

    If Len(v_champ) > 5 Then
        If left(v_champ, 1) = "T" And Mid(v_champ, 4, 1) = "_" Then
            If InStr(left$(v_champ, 5), "_") > 0 Then
                w_est_champ_tableau = P_OUI
                Exit Function
            End If
        End If
    End If
    w_est_champ_tableau = P_NON

End Function

Private Function w_est_champ_tableau_global(ByVal v_champ As String) As Integer

    Dim s As String
    Dim pos As Integer
    
    If Len(v_champ) > 5 Then
        If left(v_champ, 1) = "T" And Mid(v_champ, 4, 1) = "_" Then
            s = Mid$(v_champ, 5)
            pos = InStr(s, "_")
            If pos = 0 Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
            s = Mid$(s, pos + 1)
            If s = "" Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
            If IsNumeric(s) Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
        End If
    End If
    w_est_champ_tableau_global = P_NON
    
End Function

Private Sub w_generer_btnfusion()

    Dim prem As Boolean
    Dim x As Integer, y As Integer
    Dim LeIndex As Integer
    Dim CheminFicTxt As String, encore As Boolean, fd As Integer
    Dim fd1 As Integer
    Dim LaLigneAide As String
    Dim tblChp(), laDim As Integer
    Dim font_size As Integer, I As Integer, n As Integer
    Dim ligne As String, codechp As String, déjàChp As Boolean, libCHP As String, etapeChp As String
    Dim TypeChp As String, listeChp As String
    Dim sql As String, rs As rdoResultset
    Dim nbCol As Integer
    Dim j As Integer
    Dim rschp As rdoResultset, rs_codechp As String
    Dim width As Integer, Height As Integer
    Dim cheminOut As String
    Dim rstmp As rdoResultset
    Dim BoolTblTmp As Boolean
    
    Dim MonControl As CommandBarButton
    Dim NewBar As CommandBar
    Dim NewMenu As CommandBarPopup
    Dim libval As String, opval As String
    Dim tblTmp()
    Dim laDimTmp As Integer
    
    On Error GoTo 0
    
    ' Créer le menu
    On Error GoTo Suite_Newbar
    Set NewBar = Word_Obj.CommandBars("KaliTech")
    For I = 1 To NewBar.Controls.Count
        NewBar.Controls(I).Delete
    Next I
    NewBar.visible = True
    
    Set NewMenu = NewBar.Controls.Add(type:=msoControlPopup, Before:=1, Temporary:=True)
    NewMenu.visible = True
    If Word_CrMod_stypedoc = "F" Then
        NewMenu.Caption = "Formulaires : Liste des champs de fusion"
        NewMenu.OnAction = "Ouvrir"
    ElseIf Word_CrMod_stypedoc = "M" Then
        NewMenu.Caption = "Modèles : Liste des champs de fusion"
        NewMenu.OnAction = "Ouvrir"
    ElseIf Word_CrMod_stypedoc = "D" Then
        NewMenu.Caption = "Documents : Liste des champs de fusion"
        NewMenu.OnAction = "Ouvrir"
    End If
Suite_Newbar:
    On Error GoTo 0
    
    If Word_CrMod_stypedoc = "D" Then
        ' on ouvre le fichier et on charge
        CheminFicTxt = p_CheminModele_Loc & "\" & Word_CrMod_chemin & "\champs.txt"
        If FICH_FichierExiste(CheminFicTxt) Then
            fd = FreeFile
            Open CheminFicTxt For Input As #fd
            While Not EOF(fd)
                ' Ligne d'entete
                Line Input #fd, ligne
            Wend
            Close #fd
        
            BoolTblTmp = False
            CheminFicTxt = p_CheminModele_Loc & "\" & Word_CrMod_chemin & "\Aidechamps.txt"
            If FICH_FichierExiste(CheminFicTxt) Then
                prem = True
                fd1 = FreeFile
                Open CheminFicTxt For Input As #fd1
                While Not EOF(fd1)
                    Line Input #fd1, LaLigneAide
                    If prem Then
                        laDimTmp = 1
                        prem = False
                    Else
                        laDimTmp = UBound(tblTmp(), 2) + 1
                    End If
                    ReDim Preserve tblTmp(2, laDimTmp)
                    tblTmp(1, laDimTmp) = STR_GetChamp(LaLigneAide, "=", 0)
                    tblTmp(2, laDimTmp) = STR_GetChamp(LaLigneAide, "=", 1)
                    BoolTblTmp = True
                Wend
                Close #fd1
            End If
                
            n = STR_GetNbchamp(ligne, vbTab)
            For I = 0 To n - 1
                On Error GoTo 0
                codechp = STR_GetChamp(ligne, vbTab, I)
                libCHP = ""
                If BoolTblTmp Then
                    For j = 1 To UBound(tblTmp(), 2)
                        'MsgBox tblTmp(1, j)
                        If tblTmp(1, j) = codechp Then
                            libCHP = tblTmp(2, j)
                            Exit For
                        End If
                    Next j
                End If
                laDim = I + 1
                nbCol = 3
                ReDim Preserve tblChp(nbCol, laDim)
                tblChp(1, laDim) = codechp
                tblChp(3, laDim) = libCHP
            Next I
        Else
            MsgBox CheminFicTxt & " n'existe pas"
        End If
        
        cheminOut = p_CheminModele_Loc & "\FichFormOut.txt"
        If FICH_EffacerFichier(cheminOut, False) <> P_ERREUR Then
            If FICH_OuvrirFichier(cheminOut, FICH_ECRITURE, fd) = P_ERREUR Then
                Exit Sub
            End If
        End If
        
        On Error GoTo 0
        ' Première ligne
        Print #fd, Word_CrMod_stypedoc
        ' Deuxième ligne
        Print #fd, CheminFicTxt
        For I = 1 To UBound(tblChp(), 2)
            ligne = tblChp(1, I) & "|" & tblChp(2, I) & "|" & tblChp(3, I)
            Print #fd, ligne
        Next I
        Close #fd
    End If

End Sub

Private Function w_get_txtp(ByVal v_deb As Long, _
                            ByVal v_fin As Long, _
                            ByRef v_buf As Variant) As Integer

    Dim arange As Word.range
    
    On Error GoTo err_get_txt
    Set arange = Word_Doc.range(v_deb, v_fin)
    v_buf = arange.Text
    On Error GoTo 0
    w_get_txtp = P_OK
    Exit Function
    
err_get_txt:
    On Error GoTo 0
    Call MsgBox("Erreur w_get_txtp " & v_deb & " " & v_fin & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_get_txtp = P_ERREUR
    
End Function

Private Sub w_incrementer_bk(ByRef v_doc As Word.Document, _
                             ByRef vr_arange As Word.range, _
                             ByVal v_incr As Integer)

    Dim nom As String
    Dim bk_exist As Boolean
    Dim pos As Integer, ind As Integer, I As Integer
    Dim arange As Word.range
    Dim doc As Word.Document
    
    On Error GoTo err_open_ficr
    Set doc = Word_Obj.Documents.Add()
    On Error GoTo 0
    
    vr_arange.Select
    Word_Obj.Selection.Copy
    Set arange = doc.Content
    arange.Collapse wdCollapseEnd
    arange.Paste

'    nom = ""
'    For i = 1 To v_doc.Bookmarks.Count
'        nom = nom & v_doc.Bookmarks(i).Name & vbCrLf
'    Next i
'    MsgBox nom
    
    For I = 1 To doc.Bookmarks.Count
        nom = doc.Bookmarks(I).Name
        pos = InStrRev(nom, "_")
        If pos > 0 Then
            If IsNumeric(Mid$(nom, pos + 1)) Then
                ind = Mid$(nom, pos + 1)
                Set arange = doc.Bookmarks(nom).range
                doc.Bookmarks(nom).Delete
                Do
                    nom = left$(nom, pos) & (ind + v_incr)
                    If v_doc.Bookmarks.Exists(nom) = False Then
                        Call w_add_bookmark(doc, nom, arange)
                        bk_exist = False
                    Else
'                        MsgBox nom & " existe"
                        bk_exist = True
                        ind = ind + 1
                    End If
                Loop Until Not bk_exist
            End If
        End If
    Next I
    
    On Error Resume Next
    Set vr_arange = doc.Content
    vr_arange.Collapse wdCollapseEnd
    vr_arange.Select
    Word_Obj.Selection.TypeBackspace
    Set vr_arange = doc.Content
    vr_arange.Select
    On Error GoTo 0
    vr_arange.Copy
    Call doc.Close(savechanges:=False)
    Set doc = Nothing
    
    
    Exit Sub
    
err_open_ficr:
    MsgBox "Impossible de créer le fichier " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Exit Sub

End Sub

Private Sub w_init_tblsignet(ByVal v_plusieurs_fois_meme_chp_autor As Boolean)

    Dim nom As String, s As String
    Dim encore As Boolean, trouve As Boolean
    Dim I As Integer, j As Integer, pos As Integer, ind  As Integer
    Dim un_signet As WORD_SSIGNET
Dim v As Variant

    Word_nbsignet = Word_Doc.Bookmarks.Count
    
    If Word_nbsignet = 0 Then Exit Sub
    
    ReDim Word_tblsignet(1 To Word_Doc.Bookmarks.Count) As WORD_SSIGNET
    For I = 1 To Word_Doc.Bookmarks.Count
        nom = Word_Doc.Bookmarks(I).Name
        pos = InStrRev(nom, "_")
        trouve = False
        If pos > 0 Then
            If IsNumeric(Mid$(nom, pos + 1)) Then
                trouve = True
                Word_tblsignet(I).nom = left$(nom, pos - 1)
                Word_tblsignet(I).indice = Mid$(nom, pos + 1)
            End If
        End If
        If Not trouve Then
            Word_tblsignet(I).nom = nom
            Word_tblsignet(I).indice = 0
        End If
    Next I
    ' Tri
    Do
        encore = False
        For I = 1 To UBound(Word_tblsignet) - 1
            For j = I + 1 To UBound(Word_tblsignet)
                If UCase(Word_tblsignet(I).nom) = UCase(Word_tblsignet(j).nom) Then
                    If j > I + 1 Then
                        If UCase(Word_tblsignet(I + 1).nom) <> UCase(Word_tblsignet(j).nom) Then
                            un_signet = Word_tblsignet(I + 1)
                            Word_tblsignet(j) = Word_tblsignet(I + 1)
                            Word_tblsignet(I + 1) = un_signet
                            encore = True
                        End If
                    ElseIf Word_tblsignet(I).indice > Word_tblsignet(j).indice Then
                        un_signet = Word_tblsignet(I)
                        Word_tblsignet(j) = Word_tblsignet(I)
                        Word_tblsignet(I) = un_signet
                        encore = True
                    End If
                End If
            Next j
        Next I
    Loop Until encore = False
'v = "Après tri" & vbCrLf
'For i = 1 To UBound(word_tblsignet)
'v = v & word_tblsignet(i).nom & " " & word_tblsignet(i).indice & vbCrLf
'Next i
'MsgBox v

    ' Renommage
    nom = ""
    I = 1
    While I <= UBound(Word_tblsignet)
        If UCase(Word_tblsignet(I).nom) <> UCase(nom) Then
            nom = Word_tblsignet(I).nom
            ind = 1
        End If
        j = I
        While j <= UBound(Word_tblsignet)
            If UCase(Word_tblsignet(j).nom) = UCase(nom) Then
                I = I + 1
                If Not v_plusieurs_fois_meme_chp_autor And ind > 1 Then
                    Call MsgBox("ATTENTION : Le champ '" & Word_tblsignet(j).nom & "' a été trouvé plusieurs fois !", vbInformation + vbOKOnly, "")
                End If
                If Word_tblsignet(j).indice <> ind Then
                    Call w_renommer_bk(Word_tblsignet(j).nom, Word_tblsignet(j).indice, ind)
                    Word_tblsignet(j).indice = ind
                End If
            End If
            ind = ind + 1
            j = j + 1
        Wend
    Wend
'v = "Après renomme" & vbCrLf
'For i = 1 To UBound(word_tblsignet)
'v = v & word_tblsignet(i).nom & " " & word_tblsignet(i).indice & vbCrLf
'Next i
'MsgBox v

End Sub

Private Function w_lire_fich(ByVal v_fd As Integer, _
                             ByRef a_ligne As Variant) As Integer

    On Error GoTo fin_fichier
    Line Input #v_fd, a_ligne
    On Error GoTo 0
    w_lire_fich = P_OUI
    Exit Function

fin_fichier:
    On Error GoTo 0
    w_lire_fich = P_NON

End Function

Private Function w_put_txtbk(ByVal v_str As String, _
                             ByVal v_nombk As String) As Integer

    Dim str As String, sparam As String, nomimg As String
    Dim n As Integer, n2 As Integer, I As Integer, j As Integer
    Dim arange As Word.range
'    Dim shp As Shape
'    Dim ctrl As Object, obj_shape As Object
    
    If left$(v_str, 1) = "ê" Then
        sparam = STR_GetChamp(Mid$(v_str, 2), "ê", 0)
        str = STR_GetChamp(Mid$(v_str, 2), "ê", 1)
    Else
        sparam = ""
        str = v_str
    End If
    On Error GoTo err_range
    Set arange = Word_Doc.Bookmarks(v_nombk).range
    If InStr(v_nombk, "HyperLien") > 0 Then
'        If arange.Hyperlinks.Count > 0 Then arange.Hyperlinks(1).Delete
        arange.Text = " "
        If Len(str) > 1 Then
            arange.Text = "Accès au document"
            On Error GoTo err_add_hyp
            Call Word_Doc.Hyperlinks.Add(Anchor:=arange, Address:=str, SubAddress:="")
            On Error GoTo 0
        End If
' Cas de gestion d'un label
'        arange.text = " "
'        Set obj_shape = g_doc.InlineShapes.AddOLEControl("KaliDocCtrl.KalidocCmd", _
'                                                         arange)
'        Set ctrl = obj_shape.OLEFormat.object
'        ctrl.hNotify = Documentation.txtWord.hWnd
    Else
        On Error GoTo err_put_txt
        arange.Text = str
        On Error GoTo 0
    End If
    If sparam <> "" Then
        n = STR_GetNbchamp(sparam, "|")
        For I = 0 To n - 1
            str = STR_GetChamp(sparam, "|", I)
            If left$(str, 4) = "lien" Then
                str = Mid$(str, 6)
                On Error GoTo err_add_hyp
                Call Word_Doc.Hyperlinks.Add(Anchor:=arange, Address:=str, SubAddress:="")
                On Error GoTo 0
            ElseIf left$(str, 3) = "img" Then
                str = Mid$(str, 5)
                n2 = STR_GetNbchamp(str, "$")
                For j = 0 To n2 - 1
                    On Error Resume Next
                    nomimg = STR_GetChamp(str, "$", j)
                    Call Word_Doc.InlineShapes.AddPicture(Filename:=nomimg, LinkToFile:=False, SaveWithDocument:=True, range:=arange)
                    On Error GoTo 0
                Next j
                If j = 0 Then
                    On Error GoTo err_put_txt
                    arange.Text = " "
                    On Error GoTo 0
                End If
            End If
        Next I
    End If
    
    ' On rajoute le bookmark si pas tableau
    If g_garder_bookmark Then
        w_put_txtbk = w_add_bookmark(Word_Doc, v_nombk, arange)
        Exit Function
    End If
    
    w_put_txtbk = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk : Erreur range " & v_nombk & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_range_hyp:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk : Erreur range hyperline" & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_put_txt:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk " & v_str & " " & v_nombk & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_add_hyp:
    On Error GoTo 0
    Call MsgBox("Erreur add hyperlink " & v_str & " " & v_nombk & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    
End Function

Private Function w_range(ByVal v_deb As Long, _
                         ByVal v_fin As Long, _
                         ByRef r_range As Word.range)

    On Error GoTo err_range
    Set r_range = Word_Doc.range(v_deb, v_fin)
    On Error GoTo 0
    w_range = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    Call MsgBox("Erreur w_range " & v_deb & " " & v_fin & vbcr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    w_range = P_ERREUR
    Exit Function
    
End Function

Private Function w_rangeb(ByVal v_deb As Long, _
                         ByVal v_fin As Long, _
                         ByRef r_range As Word.range)

    On Error GoTo err_range
    Set r_range = Word_Doc.range(v_deb, v_fin)
    On Error GoTo 0
    w_rangeb = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    w_rangeb = P_ERREUR
    Exit Function
    
End Function

Private Sub w_renommer_bk(ByVal v_nom As String, _
                          ByVal v_old_ind As Integer, _
                          ByVal v_new_ind As Integer)

    Dim nom As String
    Dim arange As Word.range
    
    nom = v_nom
    If v_old_ind > 0 Then
        nom = nom & "_" & v_old_ind
    End If
    Set arange = Word_Doc.Bookmarks(nom).range
    Word_Doc.Bookmarks(nom).Delete
    nom = v_nom & "_" & v_new_ind
    Call w_add_bookmark(Word_Doc, nom, arange)
    
End Sub

Private Sub w_suppr_bk_doublon(ByRef v_doc1 As Word.Document, _
                               ByRef v_doc2 As Word.Document)
                                 
    Dim I As Integer
    
    ' Suppression des bookmarks de doc_modele qui sont déjà dans doc
    For I = 1 To v_doc2.Bookmarks.Count
        If v_doc1.Bookmarks.Exists(v_doc2.Bookmarks(I).Name) Then
            v_doc1.Bookmarks(v_doc2.Bookmarks(I).Name).Delete
        End If
    Next I

End Sub

Private Function w_ya_bktbl_dans_sel(ByVal v_nombk_tbl As String) As Boolean

    Dim I As Integer, lenv As Integer
    
    lenv = Len(v_nombk_tbl)
    For I = 1 To Word_Obj.Selection.Bookmarks.Count
        If Len(Word_Obj.Selection.Bookmarks(I).Name) > lenv Then
            If left$(Word_Obj.Selection.Bookmarks(I).Name, lenv) = v_nombk_tbl Then
                w_ya_bktbl_dans_sel = True
                Exit Function
            End If
        End If
    Next I
    
    w_ya_bktbl_dans_sel = False
    
End Function


