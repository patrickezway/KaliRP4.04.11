Attribute VB_Name = "MOpenOffice"
Option Explicit
    
Private Declare Function GetTickCount Lib "kernel32" () As Long

' Mode de travail de OOff_Fusionner
Public Const OOFF_IMPRESSION = 0
Public Const OOFF_VISU = 1
Public Const OOFF_MODIF = 2
Public Const OOFF_CREATE = 3

' Ce qu'il y a à faire au début de l'appel à OOff_Fusionner
Public Const OOFF_DEB_CROBJ = 0
Public Const OOFF_DEB_OUVDOC = 1
Public Const OOFF_DEB_RIEN = 2

' Ce qu'il y a à faire à la fin de l'appel à OOff_Fusionner
Public Const OOFF_FIN_FERMDOC = 0
Public Const OOFF_FIN_RAZOBJ = 1
Public Const OOFF_FIN_RIEN = 2

Public Ooff_SM As Object
Public Ooff_Desk As Object
Public Ooff_Doc As Object
Public Ooff_EstActif As Boolean

' Pour Ooff_CreerModele
Public Type OOFF_SSIGNET
    nom As String
    indice As Integer
End Type
Public Ooff_tblsignet() As OOFF_SSIGNET
Public Ooff_nbsignet As Integer

' Pour OOff_Fusionner
Private g_garder_bookmark As Boolean
Private Type O_STBLCHP
    nombk As String
    exist As Integer
End Type
' Valeurs de exist
Private Const O_AEVALUER = 0
Private Const O_NON = 1
Private Const O_OUI = 2

Private Sub attendre(ByVal v_msec As Long)

    Dim t As Long
    
    t = GetTickCount() + v_msec
    Do Until GetTickCount() > t
        DoEvents
    Loop

End Sub

' Le fichier doit être en local dans tous les cas
Public Function OOff_AfficherDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByVal v_fimprime As Boolean, _
                               ByVal v_fmodif As Boolean, _
                               ByVal v_nomdata As String) As Integer

    Dim s As String, nombk As String, nomdot As String, nom As String
    Dim fexist As Boolean, a_redim As Boolean, visible As Boolean
    Dim I As Integer, j As Integer, fd As Integer, n As Integer, pos As Integer
    Dim imode As Integer, erreur As Integer
    Dim s_bk As Variant
    Dim range As Object, oDsp As Object
    Dim bidon(0) As Object
    Dim objEventlistener As New ooff_listener
    Dim doc_tmp As Object
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        OOff_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If OOff_Init() = P_ERREUR Then
        OOff_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If OOff_OuvrirDoc(v_nomdoc, _
                      v_passwd, _
                      Not v_fmodif, _
                      2, _
                      Ooff_Doc) = P_ERREUR Then
        OOff_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not v_fmodif Then
        On Error GoTo lab_fin_err
        erreur = 1
        Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
        Call oDsp.executeDispatch(Ooff_Doc.CurrentController.getFrame(), ".uno:MenuBarVisible", "", 0, Array())
        Call oDsp.executeDispatch(Ooff_Doc.CurrentController.getFrame(), ".uno:FunctionBarVisible", "", 0, Array())
        On Error GoTo 0
    End If
    
    While True
        On Error GoTo lab_fin
        Set doc_tmp = Ooff_Doc.CurrentController
        On Error GoTo 0
'        SYS_Sleep (200)
'        DoEvents
        attendre (200)
    Wend
    
lab_fin:
    On Error Resume Next
    Set doc_tmp = Nothing
    Set Ooff_Doc = Nothing
    Set Ooff_Desk = Nothing
    If Not v_fmodif Then
        Set oDsp = Nothing
    End If
    Set Ooff_SM = Nothing
    On Error GoTo 0
    Ooff_EstActif = False
    
    OOff_AfficherDoc = P_OK
    Exit Function

lab_fin_err:
    On Error Resume Next
    Set Ooff_Doc = Nothing
    Set Ooff_Desk = Nothing
    If Not v_fmodif And erreur <> 1 Then
        Set oDsp = Nothing
    End If
    Set Ooff_SM = Nothing
    Ooff_EstActif = False
    On Error GoTo 0
    MsgBox "Erreur OpenOffice (" & erreur & ")" & vbcr & vbLf & Err.Number & " : " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_AfficherDoc = P_ERREUR
    Exit Function

End Function

' Le fichier doit être en local dans tous les cas
Public Function OOff_AfficherDocFusion(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByVal v_fimprime As Boolean, _
                               ByVal v_fmodif As Boolean, _
                               ByVal v_nomdata As String) As Integer

    Dim s As String, nombk As String, nomdot As String, nom As String
    Dim fexist As Boolean, a_redim As Boolean, visible As Boolean
    Dim I As Integer, j As Integer, fd As Integer, n As Integer, pos As Integer
    Dim imode As Integer, erreur As Integer
    Dim s_bk As Variant
    Dim range As Object, oDsp As Object
    Dim bidon(0) As Object
    Dim objEventlistener As New ooff_listener
    Dim objReflexion As Object
    Dim doc_tmp As Object
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        OOff_AfficherDocFusion = P_ERREUR
        Exit Function
    End If
    
    ' Fusion pour mettre à jour le signet TypeImpression
    visible = False
    If OOff_Fusionner(v_nomdoc, _
                  "", _
                  v_nomdata, _
                  True, _
                  "", _
                  True, _
                  v_passwd, _
                  visible, _
                  OOFF_CREATE, _
                  0, _
                  OOFF_DEB_CROBJ, _
                  OOFF_FIN_RIEN) = P_ERREUR Then
        Set Ooff_Desk = Nothing
        Set Ooff_SM = Nothing
        Ooff_EstActif = False
        OOff_AfficherDocFusion = P_ERREUR
        Exit Function
    End If
                        
    If Not v_fmodif Then
        Call OOff_Close(Ooff_Doc, False)
        If OOff_OuvrirDoc(v_nomdoc, _
                          v_passwd, _
                          Not v_fmodif, _
                          2, _
                          Ooff_Doc) = P_ERREUR Then
            OOff_AfficherDocFusion = P_ERREUR
            Exit Function
        End If
        On Error GoTo lab_fin_err
        erreur = 1
        Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
        Call oDsp.executeDispatch(Ooff_Doc.CurrentController.getFrame(), ".uno:MenuBarVisible", "", 0, Array())
        Call oDsp.executeDispatch(Ooff_Doc.CurrentController.getFrame(), ".uno:FunctionBarVisible", "", 0, Array())
        On Error GoTo 0
    Else
        On Error GoTo lab_fin_err
        Call Ooff_Doc.setModified(False)
        On Error GoTo 0
        Call OOff_SetVisible(Ooff_Doc, True)
    End If
    
    While True
        On Error GoTo lab_fin
        Set doc_tmp = Ooff_Doc.CurrentController
        On Error GoTo 0
        SYS_Sleep (200)
        DoEvents
    Wend
    
lab_fin:
    On Error Resume Next
    Set doc_tmp = Nothing
    Set Ooff_Doc = Nothing
    If Not v_fmodif Then
        Set oDsp = Nothing
    End If
    Set Ooff_Desk = Nothing
    Set Ooff_SM = Nothing
    On Error GoTo 0
    Ooff_EstActif = False
    
    OOff_AfficherDocFusion = P_OK
    Exit Function

lab_fin_err:
    On Error Resume Next
    Set Ooff_Doc = Nothing
    If Not v_fmodif And erreur <> 1 Then
        Set oDsp = Nothing
    End If
    Set Ooff_Desk = Nothing
    Set Ooff_SM = Nothing
    Ooff_EstActif = False
    On Error GoTo 0
    MsgBox "Erreur OpenOffice " & vbcr & vbLf & Err.Number & " : " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_AfficherDocFusion = P_ERREUR
    Exit Function

End Function

    'register the listener with the document
'    On Error GoTo lab_fin_err
'    erreur = 2
'    objEventlistener.lafin = False
'    Ooff_Doc.addEventListener objEventlistener
'    On Error GoTo 0
    
'    On Error GoTo lab_fin
'    While Not objEventlistener.lafin
'        SYS_Sleep (200)
'        DoEvents
'    Wend

'lab_fin:
'    On Error Resume Next
'    Ooff_Doc.removeEventListener objEventlistener

Public Function OOff_ChangerPasswd(ByVal v_nomdoc As String, _
                                   ByVal v_o_passwd As String, _
                                   ByVal v_n_passwd As String) As Integer
                             
    If OOff_Init() = P_ERREUR Then
        OOff_ChangerPasswd = P_OK
        Exit Function
    End If
    
    If OOff_OuvrirDoc(v_nomdoc, v_o_passwd, False, 0, Ooff_Doc) = P_ERREUR Then
        OOff_ChangerPasswd = P_ERREUR
        Exit Function
    End If
    
    Call OOff_StoreAsUrl(Ooff_Doc, v_nomdoc, v_n_passwd, "")
    
'    Call Ooff_Doc.Close(False)
'    Set Ooff_Doc = Nothing
    Call OOff_Quitter(OOFF_FIN_FERMDOC)
    
    OOff_ChangerPasswd = P_OK
    Exit Function

err_ooff:
    MsgBox "Erreur OpenOffice (OOff_ChangerPasswd) " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    OOff_ChangerPasswd = P_ERREUR
    Exit Function

End Function

Private Sub OOff_Close(ByRef r_odoc As Object, _
                       ByVal v_fenreg As Boolean)

    Call attendre(2000)
    
    Call r_odoc.Close(False)
    Set r_odoc = Nothing
End Sub
' Conversion Calc - MSO
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvC_MSO(ByVal v_nomdoc As String, _
                                ByVal v_nomdest As String) As Integer
                         
    OOff_ConvC_MSO = o_creer_autre_format(v_nomdoc, v_nomdest, "sxc", "mso")

End Function
                         
' Conversion Impress - MSO
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvI_MSO(ByVal v_nomdoc As String, _
                                ByVal v_nomdest As String) As Integer
                         
    OOff_ConvI_MSO = o_creer_autre_format(v_nomdoc, v_nomdest, "sxi", "mso")

End Function
                         
' Conversion Writer - MSO
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvW_MSO(ByVal v_nomdoc As String, _
                                ByVal v_nomdest As String) As Integer
                         
    OOff_ConvW_MSO = o_creer_autre_format(v_nomdoc, v_nomdest, "sxw", "mso")

End Function
                         
' Conversion HTML Calc
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvHTML_C(ByVal v_nomdoc As String, _
                                ByVal v_nomhtml As String) As Integer
                         
    OOff_ConvHTML_C = o_creer_autre_format(v_nomdoc, v_nomhtml, "sxc", "html")

End Function
                         
' Conversion HTML Draw
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvHTML_D(ByVal v_nomdoc As String, _
                                ByVal v_nomhtml As String) As Integer
                         
    OOff_ConvHTML_D = o_creer_autre_format(v_nomdoc, v_nomhtml, "sxd", "html")

End Function
                         
' Conversion HTML Impress
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvHTML_I(ByVal v_nomdoc As String, _
                                ByVal v_nomhtml As String) As Integer
                         
    OOff_ConvHTML_I = o_creer_autre_format(v_nomdoc, v_nomhtml, "sxi", "html")

End Function
                         
' Conversion HTML Writer
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvHTML_W(ByVal v_nomdoc As String, _
                                ByVal v_nomhtml As String) As Integer
                         
    OOff_ConvHTML_W = o_creer_autre_format(v_nomdoc, v_nomhtml, "sxw", "html")

End Function
                         
' Conversion PDF Calc
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvPDF_C(ByVal v_nomdoc As String, _
                               ByVal v_nompdf As String) As Integer
                         
    OOff_ConvPDF_C = o_creer_autre_format(v_nomdoc, v_nompdf, "sxc", "pdf")

End Function
                         
' Conversion PDF Draw
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvPDF_D(ByVal v_nomdoc As String, _
                               ByVal v_nompdf As String) As Integer
                         
    OOff_ConvPDF_D = o_creer_autre_format(v_nomdoc, v_nompdf, "sxd", "pdf")

End Function
                         
' Conversion PDF Impres
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvPDF_I(ByVal v_nomdoc As String, _
                               ByVal v_nompdf As String) As Integer
                         
    OOff_ConvPDF_I = o_creer_autre_format(v_nomdoc, v_nompdf, "sxi", "pdf")

End Function
                         
' Conversion PDF Writer
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_ConvPDF_W(ByVal v_nomdoc As String, _
                               ByVal v_nompdf As String) As Integer
                         
    OOff_ConvPDF_W = o_creer_autre_format(v_nomdoc, v_nompdf, "sxw", "pdf")

End Function
                         
Public Function OOff_CopierModele(ByVal v_nomdoc As String, _
                             ByVal v_nommodele As String, _
                             ByVal v_bcopie_entete As Boolean, _
                             ByVal v_bcopie_corps As Boolean, _
                             ByVal v_passwd As String) As Integer

    Dim readonly As Boolean
    Dim visible As Integer
    Dim doc_modele As Object, oCurs As Object
    
    If OOff_Init() = P_ERREUR Then
        OOff_CopierModele = P_ERREUR
        Exit Function
    End If
    
    ' Recopie de l'entete -> on part du modèle
    If v_bcopie_entete Then
        ' Ouvre le modèle
        visible = 0
        readonly = False
        If OOff_OuvrirDoc(v_nommodele, "", readonly, visible, doc_modele) = P_ERREUR Then
'            Call OOff_Quitter(OOFF_FIN_FERMDOC)
            OOff_CopierModele = P_ERREUR
            Exit Function
        End If
        ' Recopie du corps
        If v_bcopie_corps Then
            ' Ouvre le document d'origine
            visible = 0
            readonly = False
            If OOff_OuvrirDoc(v_nomdoc, v_passwd, readonly, visible, Ooff_Doc) = P_ERREUR Then
                Call OOff_Close(doc_modele, False)
                Call OOff_Quitter(OOFF_FIN_RAZOBJ)
                OOff_CopierModele = P_ERREUR
                Exit Function
            End If
            Call o_suppr_bk_doublon(Ooff_Doc, doc_modele)
            Call OOff_Close(Ooff_Doc, False)
        Else
            ' Efface le corps du modèle
            Set oCurs = doc_modele.Text.createTextCursor()
            Call oCurs.gotoStart(False)
            Call oCurs.gotoend(True)
            Call oCurs.setString("")
        End If
        ' Recopie du corps du document dans le modèle
        Call o_copier_corps(v_nomdoc, doc_modele)
        Call FICH_EffacerFichier(v_nomdoc, False)
        ' Le modèle devient le nouveau document
        Call OOff_StoreAsUrl(doc_modele, v_nomdoc, v_passwd, "")
        Call OOff_Close(doc_modele, False)
    ' Pas d'entete -> on part du document
    Else
        ' Recopie du corps du modèle
        If v_bcopie_corps Then
            ' Ouvre le document d'origine
            visible = 0
            readonly = False
            If OOff_OuvrirDoc(v_nomdoc, v_passwd, readonly, visible, Ooff_Doc) = P_ERREUR Then
                Call OOff_Quitter(OOFF_FIN_RAZOBJ)
                OOff_CopierModele = P_ERREUR
                Exit Function
            End If
            ' Recopie du corps du modèle dans le document
            Call o_copier_corps(v_nommodele, Ooff_Doc)
        End If
    End If
    
    Call OOff_Quitter(OOFF_FIN_RAZOBJ)
     
    OOff_CopierModele = P_OK
    
End Function

Public Function OOff_CreerModele(v_form As Form, _
                                 ByVal v_nomfich_chp As String, _
                                 ByVal v_nomdoc As String)
                        
    Dim tbl_name() As String, ssys As String, nomdoc As String, nomdoc2 As String
    Dim nomutil As String, sdat_av As String, sdat_ap As String, nomlocal As String
    Dim sext As String
    Dim ya_un_tab As Boolean, est_danstab As Boolean, inheadfoot As Boolean
    Dim tbl_inhf() As Boolean, trouve As Boolean
    Dim pos As Integer, I As Integer, notab As Integer, reponse As Integer, n As Integer
    Dim cr As Integer, ntab As Integer, lig_tab As Integer, j As Integer
    Dim visible As Integer
    Dim siz_tab As Long, tbl_start() As Long, tbl_end() As Long, numutil As Long
    Dim arange As Word.range, trange As Word.range

    ' On vérifie l'existance de champ.txt
    If Not FICH_FichierExiste(v_nomfich_chp) Then
        Call MsgBox("Le fichier '" & v_nomfich_chp & "' étant inaccessible, vous ne pouvez pas accéder aux modèles.", vbInformation + vbOKOnly, "")
        OOff_CreerModele = P_ERREUR
        Exit Function
    End If
    
lab_debut:
    ' Le .doc n'existe pas
    If Not FICH_FichierExiste(v_nomdoc) Then
        nomdoc = left$(v_nomdoc, Len(v_nomdoc) - 3) + "mod*"
        ' Le .mod existe (doc en cours de modif)
        If FICH_FichierExiste(nomdoc) Then
            nomdoc2 = Dir$(nomdoc)
            pos = InStr(nomdoc2, ".mod")
            numutil = Mid$(nomdoc2, pos + 4)
            If numutil = p_NumUtil Then
                nomdoc2 = left$(nomdoc, Len(nomdoc) - 4) & "mod" & numutil
                Call FICH_RenommerFichier(nomdoc2, v_nomdoc)
                GoTo lab_debut
            End If
            ' Récupère le nom de la personne effectuant la modif
            If P_RecupUtilPpointNom(numutil, nomutil) = P_ERREUR Then
                OOff_CreerModele = P_ERREUR
                Exit Function
            End If
            Call MsgBox("Le modèle '" & Mid$(v_nomdoc, InStrRev(v_nomdoc, "\") + 1) & "' est en cours de modification par '" & nomutil & "'." & vbcr & vbLf & vbcr & vbLf _
                        & "Vous ne pouvez pas y accéder.", vbInformation + vbOKOnly, "")
            OOff_CreerModele = P_NON
            Exit Function
        Else
            Call MsgBox("Impossible d'accéder au modèle '" & v_nomdoc & "'.", vbInformation + vbOKOnly, "")
            OOff_CreerModele = P_ERREUR
            Exit Function
        End If
    Else
        nomdoc = v_nomdoc
    End If
    
    sext = Mid$(v_nomdoc, InStrRev(v_nomdoc, "."))
    nomlocal = p_CheminDossierTravailLocal & "\" & p_CodeUtil & Format(Time, "hhmmss") & sext
    
    ' Renomme .doc en .mod sur le serveur
    nomdoc = left$(v_nomdoc, Len(v_nomdoc) - 3) & "mod" & p_NumUtil
    Call FICH_RenommerFichier(v_nomdoc, nomdoc)
    Call FICH_CopierFichier(nomdoc, nomlocal)
    sdat_av = FICH_FichierDateTime(nomlocal)
    
    ' L'utilisateur paramètre son document
    Call OOff_AfficherDoc(nomlocal, "", True, True, "")
    
    sdat_ap = FICH_FichierDateTime(nomlocal)
    ' Pas de modification
    If sdat_ap = sdat_av Then
        ' Renomme .mod en .doc
        Call FICH_RenommerFichierNoMess(nomdoc, v_nomdoc)
        Call FICH_EffacerFichier(nomlocal, False)
        OOff_CreerModele = P_NON
        Exit Function
    End If
    
    Call FRM_ResizeForm(v_form, v_form.width, v_form.Height)
    DoEvents
    v_form.Refresh
    
    If OOff_Init() = P_ERREUR Then
        OOff_CreerModele = P_ERREUR
        Exit Function
    End If
    
    visible = 0
    If OOff_OuvrirDoc(nomlocal, "", False, visible, Ooff_Doc) = P_ERREUR Then
        OOff_CreerModele = P_ERREUR
        Exit Function
    End If
    
    Call o_init_tblsignet
    
    Call Ooff_Doc.Store
    Call OOff_Close(Ooff_Doc, False)
    
    OOff_CreerModele = P_OUI
    Exit Function
    
End Function

Public Function OOff_Fusionner(ByVal v_nommod As String, _
                           ByVal v_nominit As String, _
                           ByVal v_nomdata As String, _
                           ByVal v_garder_bookmark As Boolean, _
                           ByVal v_nomdest As String, _
                           ByVal v_ecraser As Boolean, _
                           ByVal v_passwd As String, _
                           ByVal v_ooff_visible As Boolean, _
                           ByVal v_ooff_mode As Integer, _
                           ByVal v_nbex As Integer, _
                           ByVal v_deb_mode As Integer, _
                           ByVal v_fin_mode As Integer) As Integer
    
    Dim s As String
    Dim str_entete As String, nomchp As String, str_data As String
    Dim chptab As String, nomtab As String, chp As String
    Dim sval As String, nombk As String
    Dim encore As Boolean, again As Boolean
    Dim ya_book_glob As Boolean, a_redim As Boolean, b_fairefusion As Boolean
    Dim frempl As Boolean, encore_bk As Boolean
    Dim fd As Integer, poschp As Integer, pos As Integer, pos2 As Integer
    Dim I As Integer, j As Integer, ntab As Integer, nlig As Integer, n As Integer
    Dim lig_tab As Integer, col As Integer, ind As Integer, visible As Integer
    Dim idebtab As Long, ifintab As Long
    Dim tbl_chp() As O_STBLCHP
    Dim oDsp As Object, oTable As Object, oCurs As Object, oDocFrame As Object
    Dim oCurs_bk As Object
    Dim doc2 As Object, dochf As Object
    
    If v_deb_mode = OOFF_DEB_CROBJ Then
        If OOff_Init() = P_ERREUR Then
            OOff_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
'v_ooff_visible = True
    If v_ooff_visible Then
        visible = 2
    Else
        visible = 0
    End If

    If v_deb_mode <> OOFF_DEB_RIEN Then
        If OOff_OuvrirDoc(v_nommod, v_passwd, False, visible, Ooff_Doc) = P_ERREUR Then
            OOff_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    g_garder_bookmark = v_garder_bookmark
    
    b_fairefusion = True
    If v_nomdata = "" Then
        b_fairefusion = False
        GoTo lab_fin_fusion
    End If
    
'v_word_visible = True
'*********
'    If v_ooff_visible Then
'        Word_Obj.Visible = True
'        Word_Obj.ActiveWindow = True
'        a_redim = False
'        On Error Resume Next
'        If Word_Obj.WindowState <> wdWindowStateMaximize Then
'            a_redim = True
'        End If
'        Word_Obj.Activate
'        If a_redim Then Word_Obj.WindowState = wdWindowStateMaximize
'        On Error GoTo 0
'    Else
'        Word_Obj.Visible = False
'    End If
    
    If v_nomdest <> "" Then
        If v_ecraser Then
            Call OOff_StoreAsUrl(Ooff_Doc, v_nomdest, v_passwd, "")
        Else
            If v_deb_mode <> OOFF_DEB_RIEN Then
                If OOff_OuvrirDoc(v_nomdest, "", False, visible, doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                While doc2.getbookmarks().Count > 0
                    Call o_suppr_bookmark(doc2, doc2.getbookmarks().getbyIndex(0).Name)
                Wend
'J                Call o_copier_doc(Ooff_Doc, "BODY", doc2, "BODY", "FIN")
                Call o_copier_corps(v_nommod, doc2)
                Call OOff_Close(Ooff_Doc, False)
                Set Ooff_Doc = doc2
            Else
                While Ooff_Doc.getbookmarks().Count > 0
                    Call o_suppr_bookmark(Ooff_Doc, Ooff_Doc.getbookmarks().getbyIndex(0).Name)
                Wend
                If OOff_OuvrirDoc(v_nommod, "", False, 0, doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                Call o_copier_doc(doc2, "BODY", Ooff_Doc, "BODY", "FIN")
                Call OOff_Close(doc2, False)
            End If
       End If
    End If
    
    On Error GoTo err_ooff1
'    Word_Doc.MailMerge.MainDocumentType = wdNotAMergeDocument
    
    If Ooff_Doc.getbookmarks().Count < 1 Then
        GoTo lab_copie_fich
    End If

    On Error GoTo err_open_fus
    fd = FreeFile
    Open v_nomdata For Input As #fd
    On Error GoTo err_ooff2

    ' Ligne d'entete
    Line Input #fd, str_entete
    
    encore = True
    Do While encore
        If str_entete = "" Then GoTo lab_copie_fich
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
            ya_book_glob = Ooff_Doc.getbookmarks().hasByName(nomtab)
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
                ReDim Preserve tbl_chp(I) As O_STBLCHP
                tbl_chp(I).nombk = chp
                tbl_chp(I).exist = O_AEVALUER
            Loop
            ' Détermine s'il y a un tableau word
            If ya_book_glob Then
                lig_tab = -1
                If o_bk_dans_tableau(nomtab, ntab, lig_tab) Then
                    ' Efface toutes les lignes du tableau sauf la 1e
                    Set oTable = Ooff_Doc.texttables.getbyIndex(ntab)
                    If oTable.GetRows.getCount - lig_tab - 1 > 0 Then
                        Call oTable.GetRows.removeByIndex(lig_tab + 1, oTable.GetRows.getCount - lig_tab - 1)
                    End If
                    Set oTable = Ooff_Doc.texttables.getbyIndex(ntab)
                    'getCellRangeByPosition(left,top,right,bottom)
                    Set oCurs = oTable.getCellRangeByPosition(0, lig_tab, oTable.getColumns().getCount() - 1, 0)
                    Call Ooff_Doc.CurrentController.Select(oCurs)
                    Set oDocFrame = Ooff_Doc.getCurrentController().getFrame()
                    Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
                    
'                    ReDim Preserve arg(0) As Object
'                    Set arg(0) = o_set_property("ToPoint", "A1:A1")
'                    Call oDsp.executeDispatch(oDocFrame, ".uno:GoToCell", "", 0, arg())
                    
                    Call oDsp.executeDispatch(oDocFrame, ".uno:Copy", "", 0, Array())
                
                Else
                    ' Sauvegarde position bookmark tableau glob
'************* A REVOIR
'                    lig_tab = -1
'                    Set oCurs_bk = Ooff_Doc.getbookmarks.getByName(nomtab).getAnchor
'                    Call Ooff_Doc.CurrentController.Select(oCurs_bk)
'                    Set oDocFrame = Ooff_Doc.getCurrentController().getFrame()
'                    Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
'                    Call oDsp.executeDispatch(oDocFrame, ".uno:Copy", "", 0, Array())
'************* FIN A REVOIR
                End If
            End If
            ' Traitement des lignes
            again = True
            frempl = True
            Do While again
                If o_lire_fich(fd, str_data) = P_NON Then Exit Do
                If Right$(str_data, 1) = ";" Then again = False
                If Not frempl Then GoTo lab_lig_suiv
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
                    If sval = "" Then GoTo lab_chp_suiv
                    sval = STR_Remplacer(sval, "##", vbLf)
                    ' pour chaque donnée
                    If tbl_chp(col).exist = O_NON Then GoTo lab_chp_suiv
                    ind = 1
                    Do
                        encore_bk = False
                        nombk = left$(nomtab, Len(nomtab) - 2) & "_" & tbl_chp(col).nombk & "_" & ind
                        If tbl_chp(col).exist = O_AEVALUER Then
                            If Ooff_Doc.getbookmarks().hasByName(nombk) = False Then
                                tbl_chp(col).exist = O_NON
                                GoTo lab_chp_suiv
                            Else
                                tbl_chp(col).exist = O_OUI
                            End If
                        End If
                        If tbl_chp(col).exist = O_OUI And Ooff_Doc.getbookmarks().hasByName(nombk) Then
                            If o_put_txtbk(sval, nombk) = P_ERREUR Then GoTo lab_fin_err
                            If again And g_garder_bookmark And ya_book_glob Then
                                Call o_suppr_bookmark(Ooff_Doc, nombk)
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
                    If Ooff_Doc.getbookmarks().hasByName(nomtab) Then
                        Set oCurs_bk = Ooff_Doc.getbookmarks().getByName(nomtab).getAnchor
                        Call o_suppr_bookmark(Ooff_Doc, nomtab)
                    End If
                    ' Recopie du bookmark tableau à la ligne précédente
                    If lig_tab <> -1 Then
                        Call oTable.GetRows.insertByIndex(lig_tab, 1)
                        Set oTable = Ooff_Doc.texttables.getbyIndex(ntab)
                        Set oCurs = oTable.getCellRangeByPosition(0, lig_tab, oTable.getColumns().getCount() - 1, 0)
                        Call Ooff_Doc.CurrentController.Select(oCurs)
                        Set oDocFrame = Ooff_Doc.getCurrentController().getFrame()
                        Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
                        Call oDsp.executeDispatch(oDocFrame, ".uno:Paste", "", 0, Array())
                    Else
'*** A REVOIR
'                        Call Ooff_Doc.CurrentController.Select(oCurs_bk)
'                        Call oCurs_bk.goTostart(False)
'                        Call Ooff_Doc.Text.insertString(oCurs_bk, vbLf, False)
'                        Set oDocFrame = Ooff_Doc.getCurrentController().getFrame()
'                        Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
'                        Call oDsp.executeDispatch(oDocFrame, ".uno:Paste", "", 0, Array())
'*** FIN A REVOIR
                    End If
                End If
            Loop
        Else
            ' Récupère les données
            If o_lire_fich(fd, str_data) = P_NON Then GoTo lab_fin_err
            If str_data <> "" Then
                ' Remplace le bookmark
                ind = 1
                Do
                    encore_bk = False
                    nombk = nomchp & "_" & ind
                    If Ooff_Doc.getbookmarks().hasByName(nombk) = True Then
                        If ind = 1 Then
                            str_data = left$(str_data, Len(str_data) - 1)
                            str_data = Replace(str_data, "|", vbLf)
                            str_data = STR_Remplacer(str_data, "##", vbLf)
                        End If
                        If o_put_txtbk(str_data, nombk) = P_ERREUR Then GoTo lab_fin_err
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
        visible = 0
        Call o_copier_corps(v_nominit, Ooff_Doc)
        Call attendre(2000)
    End If
    
lab_fin_fusion:
    If v_ooff_mode = OOFF_IMPRESSION Then
        Close #fd
        Call Ooff_Doc.setPrinter(Array(o_set_property("Name", Printer.DeviceName)))
        CallByName Ooff_Doc, "Print", VbMethod, Array(o_set_property("CopyCount", v_nbex))
    ElseIf v_ooff_mode = OOFF_VISU Then
'        Close #fd
'        sval = Ooff_Doc.FullName
'        Call Ooff_Doc.Close(savechanges:=wdSaveChanges)
'        Set Ooff_Doc = Nothing
'        Ooff_SM.Documents.Open FileName:=sval, ReadOnly:=True, passworddocument:=v_passwd
'        GoTo lab_fin_visible
    ElseIf v_ooff_mode = OOFF_MODIF Then
        Close #fd
'        Ooff_Doc.Saved = True
        GoTo lab_fin_visible
    ElseIf v_ooff_mode = OOFF_CREATE Then
        Close #fd
        GoTo lab_fin_create
    End If
    
lab_fin:
    If v_fin_mode <> OOFF_FIN_RIEN Then
        Call OOff_Close(Ooff_Doc, False)
    End If
    If v_fin_mode = OOFF_FIN_RAZOBJ Then
'
    End If
    Call FICH_EffacerFichier(v_nomdata, False)
    OOff_Fusionner = P_OK
    Exit Function

lab_err_paste:
    MsgBox "Erreur Paste" & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Resume Next

lab_fin_create:
'MsgBox ("Effacer nomdata après les tests !")
'    Call FICH_EffacerFichier(v_nomdata, False)
    If v_fin_mode <> OOFF_FIN_RIEN Then
        Call attendre(2000)
        Call Ooff_Doc.Store
        Call OOff_Close(Ooff_Doc, True)
    End If
    If v_fin_mode = OOFF_FIN_RAZOBJ Then
        '
    End If
    OOff_Fusionner = P_OK
    Exit Function

lab_fin_visible:
    Call FICH_EffacerFichier(v_nomdata, False)
    Ooff_EstActif = False
    OOff_Fusionner = P_OK
    Exit Function
    
lab_fin_err2:
    Close #fd
    Call FICH_EffacerFichier(v_nomdata, False)
lab_fin_err1:
    Call OOff_Close(Ooff_Doc, False)
    OOff_Fusionner = P_ERREUR
    Exit Function

err_sav_dest:
    MsgBox "Impossible de sauvegarder le fichier dans " & v_nomdest & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
Resume Next
    GoTo lab_fin_err1

err_ooff1:
    MsgBox "Erreur ooff " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
'Resume Next
    GoTo lab_fin_err1
    
err_open_fus:
    MsgBox "Impossible d'ouvrir le fichier de données " & v_nomdata & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1
    
err_ooff2:
    MsgBox "Erreur ooff " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
Resume Next
    GoTo lab_fin_err2
    
lab_fin_err:
    Call MsgBox("Erreur détectée au cours de la fusion", vbOKOnly, "Fusion")
    GoTo lab_fin_err2

End Function

' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Public Function OOff_Imprimer(ByVal v_nomdoc As String, _
                              ByVal v_devicepr As String, _
                              ByVal v_nbex As Integer) As Integer

    If v_nomdoc <> "" Then
        If OOff_Init() = P_ERREUR Then
            OOff_Imprimer = P_ERREUR
            Exit Function
        End If
        
        If OOff_OuvrirDoc(v_nomdoc, "", True, 0, Ooff_Doc) = P_ERREUR Then
            Call OOff_Quitter(OOFF_FIN_RAZOBJ)
            OOff_Imprimer = P_ERREUR
            Exit Function
        End If
    End If
    
    On Error GoTo err_setpr
    Call Ooff_Doc.setPrinter(Array(o_set_property("Name", v_devicepr)))
    On Error GoTo 0
    
    On Error GoTo err_print
'    Ooff_Doc.Print (Array(o_set_property("CopyCount", v_nbex)))
    CallByName Ooff_Doc, "Print", VbMethod, Array()
    On Error GoTo 0
    
    OOff_Imprimer = P_OK
    GoTo lab_fin
    
err_setpr:
    MsgBox "Erreur OpenOffice (Imprimer) : setPrinter " & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_Imprimer = P_ERREUR
    GoTo lab_fin

err_print:
    MsgBox "Erreur OpenOffice (Imprimer) : Print " & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_Imprimer = P_ERREUR
    GoTo lab_fin

lab_fin:
    If v_nomdoc <> "" Then
        Call OOff_Quitter(OOFF_FIN_FERMDOC)
    End If
    Exit Function
    
End Function

Public Function OOff_Init()

'    If Ooff_EstActif Then
        ' Tester qu'il est effectivement toujours actif
'        On Error GoTo lab_plus_actif
'    End If
        
'    If Not Ooff_EstActif Then
        On Error GoTo err_create_sm
        Set Ooff_SM = CreateObject("com.sun.star.ServiceManager")
        On Error GoTo err_create_desktop
        Set Ooff_Desk = Ooff_SM.createInstance("com.sun.star.frame.Desktop")
        On Error GoTo 0
        Ooff_EstActif = True
'    End If
    
    OOff_Init = P_OK
    Exit Function

err_create_sm:
    MsgBox "Impossible de créer com.sun.star.ServiceManager." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_Init = P_ERREUR
    Exit Function

err_create_desktop:
    MsgBox "Impossible de créer com.sun.star.frame.Desktop." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    OOff_Init = P_ERREUR
    Exit Function

lab_plus_actif:
    Ooff_EstActif = False
    Resume Next
    
End Function

' v_visible : 0 -> le document est hidden
'             1 -> le document n'est pas hidden mais on cache la fenêtre
'             2 -> tout est visible
Public Function OOff_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByVal v_readonly As Boolean, _
                               ByVal v_visible As Integer, _
                               ByRef r_doc As Object) As Integer

    Dim I As Integer
    Dim arg() As Object
    
    On Error GoTo err_open_ficr
    I = 0
    If v_readonly Then
        ReDim Preserve arg(I) As Object
        Set arg(I) = o_set_property("ReadOnly", v_readonly)
        I = I + 1
    End If
    If v_passwd <> "" Then
        ReDim Preserve arg(I) As Object
        Set arg(I) = o_set_property("Password", v_passwd)
        I = I + 1
    End If
' QD CA MARCHERA !
    If v_visible = 0 Then
        ReDim Preserve arg(I) As Object
        Set arg(I) = o_set_property("Hidden", True)
    End If
    
    If v_nomdoc = "" Then
        ' Création
        Set r_doc = Ooff_Desk.loadComponentFromURL("private:factory/swriter", "_blank", 0, _
                                                    arg())
    Else
        ' Ouverture d'un document existant
        Set r_doc = Ooff_Desk.loadComponentFromURL(o_conv_file_url(v_nomdoc), "_blank", 0, _
                                                   arg())
    End If
    On Error GoTo 0
    
' QD CA MARCHERA !
    If v_visible = 1 Then
        Call OOff_SetVisible(r_doc, False)
    End If
    
    OOff_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "OpenOffice : Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    OOff_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function

Public Sub OOff_Quitter(ByVal v_mode As Integer)

    Ooff_EstActif = False
    
    On Error GoTo err_quit
    
    If v_mode = OOFF_FIN_FERMDOC Then
        Call OOff_Close(Ooff_Doc, False)
    End If
    Set Ooff_Desk = Nothing
    Set Ooff_SM = Nothing
    
    On Error GoTo 0
    Exit Sub

err_quit:
    Exit Sub
    
End Sub

Public Function OOff_SetVisible(ByVal v_doc As Object, _
                                ByVal v_visible As Boolean) As Integer

    On Error GoTo err_visible
    v_doc.getCurrentController().getFrame().getContainerWindow.setVisible (v_visible)
    On Error GoTo 0
    
    OOff_SetVisible = P_OK
    Exit Function
    
err_visible:
    MsgBox "OpenOffice : (SetVisible) " & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    OOff_SetVisible = P_ERREUR
    Exit Function
    
End Function

Public Function OOff_StoreAsUrl(ByVal v_doc As Object, _
                                ByVal v_nomdoc As String, _
                                ByVal v_passwd As String, _
                                ByVal v_filtre As String) As Integer
    Dim s As String
    Dim I As Integer
    Dim arg() As Object
    
    I = 0
    If v_passwd <> "" Then
        ReDim Preserve arg(I) As Object
        Set arg(I) = o_set_property("Password", v_passwd)
        I = I + 1
    End If
    If v_filtre <> "" Then
        ReDim Preserve arg(I) As Object
        Set arg(I) = o_set_property("FilterName", v_filtre)
        I = I + 1
    Else
        s = Right$(v_nomdoc, 3)
        ReDim Preserve arg(I) As Object
        If s = "odt" Then
            Set arg(I) = o_set_property("FilterName", "writer8")
            I = I + 1
        ElseIf s = "ods" Then
            Set arg(I) = o_set_property("FilterName", "calc8")
            I = I + 1
        ElseIf s = "odp" Then
            Set arg(I) = o_set_property("FilterName", "impress8")
            I = I + 1
        ElseIf s = "odg" Then
            Set arg(I) = o_set_property("FilterName", "draw8")
            I = I + 1
        End If
    End If
    
    On Error GoTo err_save
    Call v_doc.storeAsURL(o_conv_file_url(v_nomdoc), arg())
    On Error GoTo 0
    
    OOff_StoreAsUrl = P_OK
    Exit Function
    
err_save:
    MsgBox "OpenOffice : Impossible de sauvegarder le fichier " & "" & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    OOff_StoreAsUrl = P_ERREUR
    Exit Function
    
End Function

' NE SERT PAS (NE fct pas)
Private Function o_add_bookmark(ByVal v_nom As String, _
                                ByVal v_ocurs As Object) As Integer

    Dim oBK As Object, oTxt As Object
    
    On Error GoTo err_add_book
    Set oBK = Ooff_Doc.createInstance("com.sun.star.text.Bookmark")
    oBK.Name = v_nom
    Call oTxt.insertTextContent(v_ocurs, oBK, True)
    On Error GoTo 0
    o_add_bookmark = P_OK
    Exit Function
    
err_add_book:
    On Error GoTo 0
    Call MsgBox("Erreur o_suppr_bookmark " & v_nom & vbcr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    o_add_bookmark = P_ERREUR
    Exit Function
    
End Function

Private Function o_bk_dans_tableau(ByVal v_nombk As String, _
                                   ByRef r_itab As Integer, _
                                   ByRef r_lig As Integer) As Boolean

    Dim I As Integer, j As Integer
    Dim nbcols As Integer, nbrows As Integer
    Dim y_est As Boolean
    Dim oCurs As Object, oCell As Object, oCell2 As Object, oTable As Object, oTable2 As Object
    Dim oRow As Object
   
    r_itab = -1
    r_lig = -1
    y_est = False
   
'    xray Ooff_Doc
    Set oCurs = Ooff_Doc.getbookmarks.getByName(v_nombk).getAnchor

    Set oTable = oCurs.Texttable
    Set oCell = oCurs.Cell
  
    ' chercher l'index du tableau contenant le bookmark
    For I = 0 To Ooff_Doc.texttables.Count - 1
        If Ooff_Doc.texttables.getbyIndex(I).Name = oCurs.Texttable.Name Then
            r_itab = I
        End If
    Next I
   
    nbcols = oTable.Columns.Count
    nbrows = oTable.Rows.Count
   
    For I = 0 To nbcols - 1
        For j = 0 To nbrows - 1
            If oTable.getCellByPosition(I, j).CellName = oCell.CellName Then
                r_lig = j
            End If
        Next j
    Next I
   
    On Error GoTo lab_fin
    y_est = (r_itab <> -1)
   
lab_fin:
    o_bk_dans_tableau = y_est
   
End Function

Private Function o_conv_file_url(ByVal v_nomfich As String) As String

    Dim s As String
    
    s = v_nomfich
    s = Replace(s, "\", "/")
    s = Replace(s, ":", "|")
    s = "file:///" + s
    
    o_conv_file_url = s
    
End Function

Private Sub o_copier_corps1(ByVal v_nomdoc_src As String, _
                           ByRef v_docdest As Object)

    Dim liberr As String
    Dim oCurs As Object
    
    ' Insère le corps du document source dans le document dest
    On Error GoTo lab_err
    Set oCurs = v_docdest.GetText().createTextCursor()
    Call oCurs.gotoend(False)
    Call v_docdest.Text.insertString(oCurs, vbLf, False)
    Call v_docdest.CurrentController.Select(oCurs)
    Call oCurs.insertDocumentFromUrl(o_conv_file_url(v_nomdoc_src), Array())
    On Error GoTo 0
    
'    Call oCurs.goRight(1, False)
'    Call oCurs.goLeft(1, False)
'    Call oCurs.goLeft(1, True)
'    Call Ooff_Doc.CurrentController.Select(oCurs)
'    Call oCurs.setString("")
    
    Set oCurs = Nothing
    Exit Sub
    
lab_err:
    MsgBox "Erreur o_copier_corps " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "CopierModele"
    Exit Sub
    
End Sub

Private Sub o_copier_corps(ByVal v_nomdoc_src As String, _
                           ByRef v_docdest As Object)

    Dim oDoc_src As Object, oDocCtrl_src As Object, oViewCursor_src As Object
    Dim oDocCtrl_dest As Object, oCurs As Object
   
    Call OOff_OuvrirDoc(v_nomdoc_src, "", True, False, oDoc_src)
    Set oDocCtrl_src = oDoc_src.getCurrentController()
    Set oViewCursor_src = oDocCtrl_src.ViewCursor
    oViewCursor_src.gotoStart (False)
    oViewCursor_src.gotoend (True)
' DEB MODIF
    Set oCurs = v_docdest.GetText().createTextCursor()
    Call oCurs.gotoend(False)
    Call v_docdest.Text.insertString(oCurs, vbLf, False)
    Call v_docdest.CurrentController.Select(oCurs)
' FIN MODIF
   
    Set oDocCtrl_dest = v_docdest.getCurrentController()
    oDocCtrl_dest.insertTransferable (oDocCtrl_src.Transferable)
    Call OOff_Close(oDoc_src, False)
   
End Sub

Private Function o_copier_doc(ByVal v_odoc_src As Object, _
                            ByVal v_sobj_src As String, _
                            ByVal v_odoc_dest As Object, _
                            ByVal v_sobj_dest As String, _
                            ByVal v_spos_dest As String) As Integer
                         
    Dim oDsp As Object, oDocCtrl As Object, oDocFrame As Object
    
    Set oDsp = Ooff_SM.createInstance("com.sun.star.frame.DispatchHelper")
    Select Case v_sobj_src
    Case "BODY"
        Set oDocCtrl = v_odoc_src.getCurrentController()
        Set oDocFrame = oDocCtrl.getFrame()
        oDsp.executeDispatch oDocFrame, ".uno:SelectAll", "", 0, Array()
        oDsp.executeDispatch oDocFrame, ".uno:Copy", "", 0, Array()
    End Select
    
    Call attendre(2000)

    Select Case v_sobj_dest
    Case "BODY"
        Set oDocCtrl = v_odoc_dest.getCurrentController()
        Set oDocFrame = oDocCtrl.getFrame()
        If v_spos_dest = "FIN" Then
            oDsp.executeDispatch oDocFrame, ".uno:GoToEndOfLine", "", 0, Array()
        End If
        oDsp.executeDispatch oDocFrame, ".uno:Paste", "", 0, Array()
    End Select

End Function

Private Sub o_init_tblsignet()

    Dim nom As String, s As String
    Dim encore As Boolean, trouve As Boolean
    Dim I As Integer, j As Integer, pos As Integer, ind  As Integer
    Dim un_signet As WORD_SSIGNET
Dim v As Variant

    Word_nbsignet = Ooff_Doc.getbookmarks().Count
    
    If Word_nbsignet = 0 Then
        Exit Sub
    End If
    
    ReDim Word_tblsignet(1 To Ooff_Doc.getbookmarks().Count) As WORD_SSIGNET
    For I = 1 To Ooff_Doc.getbookmarks().Count
        nom = Ooff_Doc.getbookmarks().getbyIndex(I - 1).Name
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
    ' Tri du tableau
    Do
        encore = False
        For I = 1 To UBound(Word_tblsignet) - 1
            For j = I + 1 To UBound(Word_tblsignet)
                If Word_tblsignet(I).nom = Word_tblsignet(j).nom Then
                    If j > I + 1 Then
                        If Word_tblsignet(I + 1).nom <> Word_tblsignet(j).nom Then
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
        If Word_tblsignet(I).nom <> nom Then
            nom = Word_tblsignet(I).nom
            ind = 1
        End If
        j = I
        While j <= UBound(Word_tblsignet)
            If Word_tblsignet(j).nom = nom Then
                I = I + 1
                If Word_tblsignet(j).indice <> ind Then
'                    Ooff_Doc.getBookmarks().getByIndex(Word_tblsignet(j).indice - 1).Name = Word_tblsignet(j).nom & "_" & Word_tblsignet(j).indice
'                    Call o_renommer_signet(Word_tblsignet(j).nom, Word_tblsignet(j).indice, ind)
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

' Conversion html ou pdf
' Si v_nomdoc est vide on s'appuie directement sur OOff_doc
Private Function o_creer_autre_format(ByVal v_nomdoc_src As String, _
                                      ByVal v_nomdoc_dest As String, _
                                      ByVal v_stype_src As String, _
                                      ByVal v_stype_dest As String) As Integer
                         
    Dim sfiltre As String
    Dim visible As Integer
    
    If v_stype_src = "sxw" Then
        If v_stype_dest = "html" Then
            sfiltre = "HTML (StarWriter)"
            visible = 1
        ElseIf v_stype_dest = "pdf" Then
            sfiltre = "writer_pdf_Export"
            visible = 0
        ElseIf v_stype_dest = "mso" Then
            sfiltre = "MS Word 97"
            visible = 1
        Else
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
    ElseIf v_stype_src = "sxc" Then
        If v_stype_dest = "html" Then
            sfiltre = "HTML (StarCalc)"
            visible = 1
        ElseIf v_stype_dest = "pdf" Then
            sfiltre = "calc_pdf_Export"
            visible = 0
        ElseIf v_stype_dest = "mso" Then
            sfiltre = "MS Excel 97"
            visible = 1
        Else
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
    ElseIf v_stype_src = "sxd" Then
        If v_stype_dest = "html" Then
            ' N'existe pas pour l'instant
'            sfiltre = "HTML (StarDraw)"
'            visible = 1
            o_creer_autre_format = P_ERREUR
            Exit Function
        ElseIf v_stype_dest = "pdf" Then
            sfiltre = "draw_pdf_Export"
            visible = 0
        Else
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
    ElseIf v_stype_src = "sxi" Then
        If v_stype_dest = "html" Then
            sfiltre = "HTML (StarImpress)"
            visible = 1
            o_creer_autre_format = P_ERREUR
            Exit Function
        ElseIf v_stype_dest = "pdf" Then
            sfiltre = "impress_pdf_Export"
            visible = 0
        ElseIf v_stype_dest = "mso" Then
            sfiltre = "MS Powerpoint 97"
            visible = 1
        Else
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
    Else
        o_creer_autre_format = P_ERREUR
        Exit Function
    End If
    
    If v_nomdoc_src <> "" Then
        If OOff_Init() = P_ERREUR Then
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
        
        If OOff_OuvrirDoc(v_nomdoc_src, "", True, visible, Ooff_Doc) = P_ERREUR Then
            Call OOff_Quitter(OOFF_FIN_RAZOBJ)
            o_creer_autre_format = P_ERREUR
            Exit Function
        End If
    End If
    
    On Error GoTo err_saveas
    Call Ooff_Doc.storeToURL(o_conv_file_url(v_nomdoc_dest), _
                             Array(o_set_property("FilterName", sfiltre)))
    On Error GoTo 0
    
    If v_nomdoc_src <> "" Then
        Call OOff_Quitter(OOFF_FIN_FERMDOC)
    End If
    
    o_creer_autre_format = P_OK
    Exit Function
    
err_saveas:
    MsgBox "Erreur OpenOffice (creer_autre_format) : storeToURL " & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    o_creer_autre_format = P_ERREUR
    Exit Function

End Function
                         
Private Function o_formate_nomfich(ByVal v_nomfich As String) As String

    Dim nom As String
    
    nom = Replace(v_nomfich, "\", "/")
    nom = Replace(nom, ":", "|")
    nom = Replace(nom, " ", "%20")
    nom = "file:///" + nom
    
    o_formate_nomfich = nom
    
End Function

Private Function o_lire_fich(ByVal v_fd As Integer, _
                             ByRef a_ligne As Variant) As Integer

    On Error GoTo fin_fichier
    Line Input #v_fd, a_ligne
    On Error GoTo 0
    o_lire_fich = P_OUI
    Exit Function

fin_fichier:
    On Error GoTo 0
    o_lire_fich = P_NON

End Function

Private Function o_put_txtbk(ByVal v_str As String, _
                             ByVal v_nombk As String) As Integer

    Dim str As String, sparam As String, nomimg As String
    Dim I As Integer, j As Integer, n As Integer, n2 As Integer
    Dim arange As Object, oText As Object, oCursor As Object, oImage As Object
    
    Call attendre(500)
    
    If left$(v_str, 1) = "ê" Then
        sparam = STR_GetChamp(Mid$(v_str, 2), "ê", 0)
        str = STR_GetChamp(Mid$(v_str, 2), "ê", 1)
    Else
        sparam = ""
        str = v_str
    End If
    
    On Error GoTo err_put_txt
    Set arange = Ooff_Doc.getbookmarks().getByName(v_nombk).getAnchor()
    arange.setString (v_str)
    On Error GoTo 0
    
    If sparam <> "" Then
        n = STR_GetNbchamp(sparam, "|")
        For I = 0 To n - 1
            str = STR_GetChamp(sparam, "|", I)
            If left$(str, 4) = "lien" Then
'                str = Mid$(str, 6)
'                On Error GoTo err_add_hyp
'                Call Word_Doc.Hyperlinks.Add(Anchor:=arange, Address:=str, SubAddress:="")
'                On Error GoTo 0
            ElseIf left$(str, 3) = "img" Then
                str = Mid$(str, 5)
                n2 = STR_GetNbchamp(str, "$")
                For j = 0 To n2 - 1
                    On Error Resume Next
                    nomimg = STR_GetChamp(str, "$", j)
                    Set arange = Ooff_Doc.getbookmarks().getByName(v_nombk).getAnchor()
                    Set oText = arange.Text
                    arange.setString ("")
                    Set oCursor = oText.createTextCursorByRange(arange.start)
                    Set oImage = Ooff_Doc.createInstance("com.sun.star.text.TextGraphicObject")
                    oImage.GraphicURL = o_conv_file_url(nomimg)
                    oImage.AnchorType = 1 '=com.sun.star.Text.TextContentAnchorType.AS_CHARACTER
                    Call oText.insertTextContent(oCursor, oImage, False)
                    
                    'forcer le chargement de l'image pour
                    'contourner bug openoffice 85105 version 2.3+
                    oImage.getPropertyValue ("IsPixelContour")
                    
                    Set oImage = Nothing
                    Set oCursor = Nothing
                    Set oText = Nothing
                    Set arange = Nothing
                Next j
                If j = 0 Then
                    On Error GoTo err_put_txt
                    arange.setString (v_str)
                    On Error GoTo 0
                End If
            End If
        Next I
    End If
    
    ' On supprime le bookmark (cas du tableau)
    If Not g_garder_bookmark Then
        o_put_txtbk = o_suppr_bookmark(Ooff_Doc, v_nombk)
        Exit Function
    End If
    
    o_put_txtbk = P_OK
    Exit Function
    
err_put_txt:
    On Error GoTo 0
    Call MsgBox("Erreur o_put_txtbk " & v_str & " " & v_nombk & vbcr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    o_put_txtbk = P_ERREUR
    Exit Function
    
End Function

Private Function o_range(ByVal v_deb As Long, _
                         ByVal v_fin As Long, _
                         ByRef r_range As Object)

    On Error GoTo err_range
    Set r_range = Ooff_Doc.range(v_deb, v_fin)
    On Error GoTo 0
    o_range = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    Call MsgBox("Erreur o_range " & v_deb & " " & v_fin & vbcr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    o_range = P_ERREUR
    Exit Function
    
End Function

Private Sub o_renommer_bk(ByVal v_nom As String, _
                              ByVal v_old_ind As Integer, _
                              ByVal v_new_ind As Integer)

    Dim nom As String
    Dim arange As Word.range
    
    nom = v_nom
    If v_old_ind > 0 Then
        nom = nom & "_" & v_old_ind
    End If
'    Set arange = Word_Doc.Bookmarks(nom).range
'    Word_Doc.Bookmarks(nom).Delete
'    nom = v_nom & "_" & v_new_ind
'    Call w_add_bookmark(nom, arange)
    
End Sub

Private Function o_set_property(ByVal v_name As Variant, _
                                ByVal v_val As Variant) As Object

    Dim struct As Object
    
    Set struct = Ooff_SM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    
    struct.Name = v_name
    struct.Value = v_val

    Set o_set_property = struct
    
End Function

Private Sub o_suppr_bk_doublon(ByRef v_doc1 As Object, _
                               ByRef v_doc2 As Object)
                                 
    Dim nombk As String
    Dim I As Integer
    
    ' Suppression des bookmarks de doc_modele qui sont déjà dans doc
    For I = 0 To v_doc2.getbookmarks().Count - 1
        nombk = v_doc2.getbookmarks().getbyIndex(I).Name
        If v_doc1.getbookmarks.hasByName(nombk) Then
            Call o_suppr_bookmark(v_doc1, nombk)
        End If
    Next I

End Sub

Private Function o_suppr_bookmark(ByVal v_odoc As Object, _
                                  ByVal v_nom As String) As Integer

    Dim oBK As Object, oTxt As Object
    
    On Error GoTo err_suppr_book
    Set oBK = v_odoc.getbookmarks().getByName(v_nom)
    Set oTxt = v_odoc.Text
    Call oTxt.removeTextContent(oBK)
    On Error GoTo 0
    
    o_suppr_bookmark = P_OK
    Exit Function
    
err_suppr_book:
    On Error GoTo 0
    Call MsgBox("Erreur o_suppr_bookmark " & v_nom & vbcr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    o_suppr_bookmark = P_ERREUR
    Exit Function
    
End Function




