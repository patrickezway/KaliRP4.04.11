Attribute VB_Name = "KD_Document"
Option Explicit

Public Const P_TYPDOC_PAPIER = 100

Public Const P_CYCLE_CHGTDATE = 100
Public Const P_CYCLE_RENOUV = 101
Public Const P_CYCLE_RELECTURE_REVISION = 102
Public Const P_CYCLE_EVAL = 103

' Cycles documentation
Public p_cycle_relecture As Long
Public p_cycle_verifliens As Long
Public p_cycle_consultable As Long
Public p_cycle_diffusion As Long
Private Type SCYCLEDOCS
    etape As String
    acteur As String
    action As String
    ordre_si_refus As Long
    informer_si_refus As Boolean
    modifiable As Boolean
End Type
Public p_scycledocs() As SCYCLEDOCS

' Champs dans les documents
'Public Const P_CHP_TYPIMP = 0
Public Const P_CHP_TITREDOC = 1
Public Const P_CHP_REFDOC = 2
Public Const P_CHP_DESCRDOC = 3
Public Const P_CHP_DESTDOC = 4
Public Const P_CHP_REFERENTIELDOC = 5
Public Const P_CHP_IMGREFERENTIELDOC = 6
Public Const P_CHP_LIENDOC = 7
Public Const P_CHP_NUMVERS = 8
Public Const P_CHP_ETATVERS = 9
Public Const P_CHP_DATEAPPLIVERS = 10
Public Const P_CHP_LIEUVERS = 11
Public Const P_CHP_LIENDOCVERS = 12
Public Const P_CHP_NOMENT = 13
Public Const P_CHP_CODENT = 14
Public Const P_CHP_TITREDOS = 15
Public Const P_CHP_DESCRDOS = 16
Public Const P_CHP_ARBORESCENCEDOC = 17
Public Const P_CHP_MOTIFVERS = 18
Public Const P_CHP_MOTIFLVERS = 19
Public Const P_CHP_ACTEURDOC = 20
Public Const P_CHP_POSTEACTEURDOC = 21
Public Const P_CHP_DATEACTION = 22
'Public Const P_CHP_LIBLASTMAJ = 23
Public Const P_CHP_REDACTEUR1 = 24
Public Const P_CHP_REVISION_PREV = 25
Public Const P_CHP_REVISION_SUIV = 26
Public Const P_CHP_NATUREDOC = 27
Public Const P_CHP_EMETTEUR = 28

Public p_nomdoc_ouv As String

' v_stype : A pour Action
'           D pour Diffusion
'           I pour Info
'           M pour demande de modif
'           F pour formulaires à traiter
Public Function P_MajUtilADIM(ByVal v_numutil As Long, _
                              ByVal v_stype As String, _
                              ByVal v_nb As Integer) As Integer
                               
                               
    Dim sql As String
    Dim newaction As Boolean, newdiff As Boolean, newinfo As Boolean, newdem As Boolean
    Dim nbessai As Integer, nbaction As Integer, nbdiff As Integer, nbinfo As Integer, nbdem As Integer
    Dim rs As rdoResultset
    
    ' 0 : efface
    ' > 0 : rajoute v_nb
    ' < 0 : retire v_nb
    nbessai = 0
lab_debut:
    sql = "select * from UtilADIM" _
        & " where UA_UNum=" & v_numutil _
        & " and UA_DONum=" & p_NumDocs
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    If rs.EOF Then
        If v_nb <= 0 Then
            rs.Close
            P_MajUtilADIM = P_OK
            Exit Function
        End If
        On Error GoTo err_addnew
        rs.AddNew
        On Error GoTo err_affecte
        rs("UA_UNum").Value = v_numutil
        rs("UA_DONum").Value = p_NumDocs
    Else
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
    End If
    Select Case v_stype
    Case "A"
        If v_nb <= 0 Then
            If Not IsNull(rs("UA_Nbaction").Value) Then
                rs("UA_NbAction").Value = rs("UA_NbAction").Value + v_nb
                If rs("UA_NbAction").Value < 0 Then
                    rs("UA_NbAction").Value = 0
                End If
            End If
            If v_numutil = p_NumUtil Then
                nbaction = rs("UA_NbAction").Value
            End If
        Else
            If IsNull(rs("UA_NbAction").Value) Then
                rs("UA_NbAction").Value = v_nb
            Else
                rs("UA_NbAction").Value = rs("UA_NbAction").Value + v_nb
            End If
            rs("UA_NewAction").Value = True
            If v_numutil = p_NumUtil Then
                nbaction = rs("UA_NbAction").Value
                newaction = True
            End If
        End If
    Case "D"
        If v_nb <= 0 Then
            If Not IsNull(rs("UA_NbDiff").Value) Then
                rs("UA_NbDiff").Value = rs("UA_NbDiff").Value + v_nb
                If rs("UA_NbDiff").Value < 0 Then
                    rs("UA_NbDiff").Value = 0
                End If
            End If
            If v_numutil = p_NumUtil Then
                nbdiff = rs("UA_NbDiff").Value
            End If
        Else
            If IsNull(rs("UA_NbDiff").Value) Then
                rs("UA_NbDiff").Value = v_nb
            Else
                rs("UA_NbDiff").Value = rs("UA_NbDiff").Value + v_nb
            End If
            rs("UA_NewDiff").Value = True
            If v_numutil = p_NumUtil Then
                nbdiff = rs("UA_NbDiff").Value
                newdiff = True
            End If
        End If
    Case "I"
        If v_nb <= 0 Then
            If Not IsNull(rs("UA_Nbinfo").Value) Then
                rs("UA_Nbinfo").Value = rs("UA_Nbinfo").Value + v_nb
                If rs("UA_Nbinfo").Value < 0 Then
                    rs("UA_Nbinfo").Value = 0
                End If
            End If
            If v_numutil = p_NumUtil Then
                nbinfo = rs("UA_Nbinfo").Value
            End If
        Else
            If IsNull(rs("UA_Nbinfo").Value) Then
                rs("UA_Nbinfo").Value = v_nb
            Else
                rs("UA_Nbinfo").Value = rs("UA_Nbinfo").Value + v_nb
            End If
            rs("UA_NewInfo").Value = True
            If v_numutil = p_NumUtil Then
                nbinfo = rs("UA_Nbinfo").Value
                newinfo = True
            End If
        End If
    Case "M"
        If v_nb <= 0 Then
            If Not IsNull(rs("UA_Nbmodif").Value) Then
                rs("UA_Nbmodif").Value = rs("UA_Nbmodif").Value + v_nb
                If rs("UA_Nbmodif").Value < 0 Then
                    rs("UA_Nbmodif").Value = 0
                End If
            End If
            If v_numutil = p_NumUtil Then
                nbdem = rs("UA_Nbmodif").Value
            End If
        Else
            If IsNull(rs("UA_Nbmodif").Value) Then
                rs("UA_Nbmodif").Value = v_nb
            Else
                rs("UA_Nbmodif").Value = rs("UA_Nbmodif").Value + v_nb
            End If
            rs("UA_NewModif").Value = True
            If v_numutil = p_NumUtil Then
                nbdem = rs("UA_Nbmodif").Value
                newdem = True
            End If
        End If
    Case "F"
        If v_nb <= 0 Then
            If Not IsNull(rs("UA_Nbform").Value) Then
                rs("UA_Nbform").Value = rs("UA_Nbform").Value + v_nb
                If rs("UA_Nbform").Value < 0 Then
                    rs("UA_Nbform").Value = 0
                End If
            End If
            If v_numutil = p_NumUtil Then
'                p_nbform = rs("UA_Nbform").Value
            End If
        Else
            If IsNull(rs("UA_Nbform").Value) Then
                rs("UA_Nbform").Value = v_nb
            Else
                rs("UA_Nbform").Value = rs("UA_Nbform").Value + v_nb
            End If
            rs("UA_Newform").Value = True
            If v_numutil = p_NumUtil Then
'                p_nbdem = rs("UA_Nbform").Value
'                p_newform = True
            End If
        End If
    End Select
    On Error GoTo err_update
    rs.Update
    On Error GoTo 0
    rs.Close
    
    P_MajUtilADIM = P_OK
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    P_MajUtilADIM = P_ERREUR
    Exit Function

err_no_resultset:
    MsgBox "Pas de ligne pour " + sql, vbOKOnly + vbCritical, ""
    rs.Close
    P_MajUtilADIM = P_ERREUR
    Exit Function

err_edit:
    MsgBox "Erreur Edit pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    rs.Close
    P_MajUtilADIM = P_ERREUR
    Exit Function
    
err_addnew:
    MsgBox "Erreur AddNew pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    rs.Close
    P_MajUtilADIM = P_ERREUR
    Exit Function
    
err_affecte:
    MsgBox "Erreur affectation pour " + sql, vbOKOnly + vbCritical, ""
    rs.Close
    P_MajUtilADIM = P_ERREUR
    Exit Function
    
err_update:
    rs.Close
    If nbessai > 3 Then
        MsgBox "Erreur Update (4 tentatives) pour " & sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
        P_MajUtilADIM = P_ERREUR
        Exit Function
    End If
    nbessai = nbessai + 1
    GoTo lab_debut

End Function

Public Function P_AjouterMajDoc(ByVal v_numdoc As Long, _
                                ByVal v_numvers As Long, _
                                ByVal v_type As Integer) As Integer
                                
    Dim sql As String
    Dim num As Long, lbid As Long
    Dim rs As rdoResultset
    
    sql = "select md_num, md_type from majdoc where md_dnum=" & v_numdoc _
        & " and md_numvers=" & v_numvers
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_AjouterMajDoc = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        Call Odbc_AddNew("MajDoc", _
                         "MD_Num", _
                         "md_seq", _
                         False, _
                         lbid, _
                         "MD_DNum", v_numdoc, _
                         "MD_NumVers", v_numvers, _
                         "MD_Type", v_type)
    Else
        If rs("md_type").Value >= v_type Then
            rs.Close
        Else
            num = rs("md_num").Value
            rs.Close
            Call Odbc_Update("majdoc", "md_num", "where md_num=" & num, _
                             "md_type", v_type)
        End If
    End If
    
End Function

' Charge les cycles associés à la documentation en cours dans p_scycledocs
Public Function P_ChargerCycles() As Integer

    Dim sql As String
    Dim I As Integer
    Dim rs As rdoResultset
    
    Erase p_scycledocs()
    p_cycle_relecture = -1
    p_cycle_verifliens = -1
    p_cycle_diffusion = -1
    p_cycle_consultable = -1
    
    ' Récupère les paramètres CYCLE de la documentation
    sql = "select * from Cycle" _
        & " order by CY_Ordre"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        P_ChargerCycles = P_ERREUR
        Exit Function
    End If
    I = 0
    While Not rs.EOF
        ReDim Preserve p_scycledocs(I) As SCYCLEDOCS
        p_scycledocs(I).etape = rs("CY_Etape").Value
        p_scycledocs(I).acteur = rs("CY_Acteur").Value
        p_scycledocs(I).action = rs("CY_Action").Value
        p_scycledocs(I).ordre_si_refus = rs("CY_OrdreSiRefus").Value
        p_scycledocs(I).informer_si_refus = rs("CY_InformerSiRefus").Value
        p_scycledocs(I).modifiable = rs("CY_Modifiable").Value
        If rs("CY_Relecture").Value Then p_cycle_relecture = I
        If rs("CY_VerifLiens").Value Then p_cycle_verifliens = I
        If rs("CY_Diffusion").Value Then p_cycle_diffusion = I
        If rs("CY_Consultable").Value Then p_cycle_consultable = I
        rs.MoveNext
        I = I + 1
    Wend
    rs.Close
    
    P_ChargerCycles = P_OK

End Function

Public Function P_DecoderMention(ByVal v_mention As String) As String

    Dim pos As Integer
    
    If v_mention = "" Then
        P_DecoderMention = " "
    Else
        pos = InStr(UCase(v_mention), "<DATE>")
        If pos > 0 Then
            P_DecoderMention = left$(v_mention, pos - 1) + Format(Date, "dd/mm/yyyy") + Mid$(v_mention, pos + 6)
        Else
            P_DecoderMention = v_mention
        End If
    End If

End Function

Public Function P_DecodePasswd(ByVal v_numdoc As Long, _
                               ByVal v_passwd As String) As String
                               
    If v_passwd = "AUTO" Then
        P_DecodePasswd = v_numdoc
    Else
        P_DecodePasswd = v_passwd
    End If
    
End Function

Public Function P_DocUneVersion(ByVal v_numdoc As Long) As Boolean

    Dim sql As String
    Dim rs As rdoResultset
    
    sql = "select DON_UneSeuleVersion from Document, DocsNature" _
        & " where D_Num=" & v_numdoc _
        & " and DON_DONum=D_DONum" _
        & " and DON_NDNum=D_NDNum"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        P_DocUneVersion = False
    Else
        P_DocUneVersion = rs("DON_UneSeuleVersion").Value
        rs.Close
    End If
    
    
End Function

Public Function P_EstActeur(ByVal v_numdoc As Long, _
                            ByVal v_numutil As Long) As Integer

    Dim sql As String
    Dim nb As Long
    Dim rs As rdoResultset
    
    sql = "select count(*)" _
        & " from DocUtil" _
        & " where DU_DNum=" & v_numdoc _
        & " and DU_UNum=" & v_numutil _
        & " and DU_CYOrdre>0"
    If Odbc_Count(sql, nb) = P_ERREUR Then
        P_EstActeur = P_NON
        Exit Function
    End If
    If nb > 0 Then
        P_EstActeur = P_OUI
        Exit Function
    End If

    ' Cas du chgt d'acteur temporaire
    sql = "select count(*)" _
        & " from DocAction" _
        & " where DAC_DNum=" & v_numdoc _
        & " and DAC_UNum=" & v_numutil
    If Odbc_Count(sql, nb) = P_ERREUR Then
        P_EstActeur = P_NON
        Exit Function
    End If
    If nb > 0 Then
        P_EstActeur = P_OUI
        Exit Function
    End If
        
    P_EstActeur = P_NON
    Exit Function
    
' On laisse tomber ce cas pour l'instant
'    If v_numvers = 0 Then
'        sql = "select D_NumVers from Document where D_Num=" & v_numdoc
'        If Odbc_Select(sql, rs) = P_ERREUR Then
'            P_EstActeur = P_NON
'            Exit Function
'        End If
'        numvers = rs("D_NumVers").Value
'        rs.Close
'    Else
'        numvers = v_numvers
'    End If
'    sql = "select count(*)" _
'        & " from DocEtapeVersion" _
'        & " where DEV_DNum=" & v_numdoc _
'        & " and DEV_UNum=" & v_numutil _
'        & " and DEV_CYOrdre>0"
'    If Odbc_Count(sql, nb) = P_ERREUR Then
'        P_EstActeur = P_NON
'        Exit Function
'    End If
'   If nb > 0 Then
'        P_EstActeur = P_OUI
'        Exit Function
'    End If
    
End Function

Public Function P_MajUtilAboRecu(ByVal v_numdoc As Long) As Integer

    Dim sql As String
    Dim est_public As Boolean, ajouter As Boolean
    Dim ind_dos As Integer, ind_util As Integer, I As Integer
    Dim lnb As Long, numDos As Long, lbid As Long, tbldos() As Long, tblutil() As Long
    Dim rs As rdoResultset
    
    ' Charge l'arborescence de dossiers du document
    sql = "select D_DSNum, D_Public from Document where D_Num=" & v_numdoc
    If Odbc_RecupVal(sql, numDos, est_public) = P_ERREUR Then
        P_MajUtilAboRecu = P_ERREUR
        Exit Function
    End If
    ind_dos = -1
    While numDos > 0
        ind_dos = ind_dos + 1
        ReDim Preserve tbldos(ind_dos) As Long
        tbldos(ind_dos) = numDos
        sql = "select DS_NumPere from Dossier where DS_Num=" & numDos
        If Odbc_RecupVal(sql, numDos) = P_ERREUR Then
            P_MajUtilAboRecu = P_ERREUR
            Exit Function
        End If
    Wend
    
    If Not est_public Then
        ind_util = -1
        sql = "select distinct DU_UNum from DocUtil where DU_DNum=" & v_numdoc _
                & " union select DPD_UNum from DocPrmDiffusion where DPD_DNum=" & v_numdoc _
                & " and DPD_UNum in (select UABO_UNum from UtilAbon where UABO_DNum=" & v_numdoc & ")"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            P_MajUtilAboRecu = P_ERREUR
            Exit Function
        End If
        While Not rs.EOF
            ind_util = ind_util + 1
            ReDim Preserve tblutil(ind_util) As Long
            tblutil(ind_util) = rs("DU_UNum").Value
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    sql = "select UABO_UNum from UtilAbon where UABO_DNum=" & v_numdoc
    For I = 0 To UBound(tbldos)
        sql = sql & " or UABO_DSNum=" & tbldos(I)
    Next I
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_MajUtilAboRecu = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        If est_public Then
            ajouter = True
        Else
            ajouter = False
            For I = 0 To ind_util
                If tblutil(I) = rs("UABO_UNum").Value Then
                    ajouter = True
                    Exit For
                End If
            Next I
        End If
        If ajouter Then
            Call Odbc_Delete("UtilAbonRecu", _
                             "UABOR_Num", _
                            "where UABOR_UNum=" & rs("UABO_UNum").Value & " and UABOR_DNum=" & v_numdoc, _
                            lnb)
            Call Odbc_AddNew("UtilAbonRecu", _
                             "UABOR_Num", _
                             "uabor_seq", _
                             False, _
                             lbid, _
                             "UABOR_DNum", v_numdoc, _
                             "UABOR_UNum", rs("UABO_UNum").Value, _
                             "UABOR_Datenv", Date, _
                             "UABOR_Ack", False)
            ' Envoyer un mail ?
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    P_MajUtilAboRecu = P_OK

End Function

Public Function P_RecupDocIdentFich(ByVal v_numdoc As Long, _
                                    ByRef r_ident As String) As Integer
            
    Dim sql As String
    
    sql = "select D_Ident from Document" _
        & " where D_Num=" & v_numdoc
    If Odbc_RecupVal(sql, r_ident) = P_ERREUR Then
        P_RecupDocIdentFich = P_ERREUR
        Exit Function
    End If

    r_ident = Replace(r_ident, " ", "-")
    r_ident = Replace(r_ident, "/", "-")
    r_ident = Replace(r_ident, "\", "-")
    
    P_RecupDocIdentFich = P_OK
    
End Function

Public Function P_SaisirDateFinNews(ByVal v_sdatefin As String) As String
                                         
    Dim sdate As String
    Dim ichp As Integer
    
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Marquer le document comme étant nouveau", "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    If v_sdatefin <> "" Then
        Call SAIS_AddChamp("Date de fin de nouveauté actuelle : " & v_sdatefin, 10, 10, SAIS_TYP_DATE, True, v_sdatefin)
        ichp = 1
    Else
        ichp = 0
    End If
    Call SAIS_AddChamp("Jusqu'au", 10, 10, SAIS_TYP_DATE, True)
    If v_sdatefin <> "" Then
        Call SAIS_AddChamp("(Ne saisissez rien si vous ne voulez pas changer la date de fin de nouveauté)", 10, 10, SAIS_TYP_DATE, True)
        Call SAIS_AddChamp("(Pour SUPPRIMER la marque de nouveauté, cliquez sur 'Annuler la marque de nouveauté')", 0, 0, SAIS_TYP_DATE, True)
        Call SAIS_AddBouton("&Annuler la marque de nouveauté", "", 0, 0, 2000)
    Else
        Call SAIS_AddChamp("(Ne saisissez rien si vous ne voulez pas que le document soit dans les nouveautés de KaliWeb)", 0, 0, SAIS_TYP_DATE, True)
    End If
    
lab_saisie:
    Saisie.Show 1
        
    If SAIS_Saisie.retour = 1 Then
        P_SaisirDateFinNews = ""
        Exit Function
    End If
    
    sdate = SAIS_Saisie.champs(ichp).sval
    If v_sdatefin <> "" Then
        If sdate = "" Then
            sdate = SAIS_Saisie.champs(0).sval
        End If
    End If
    
    If sdate <> "" Then
        If CDate(sdate) < Date Then
            Call MsgBox("La date de fin de marque de la nouveauté ne peut être antérieure à la date du jour.", vbExclamation + vbOKOnly, "")
            GoTo lab_saisie
        End If
    End If
    
    P_SaisirDateFinNews = sdate
    
End Function

Public Function P_UtilEstRespPrincOuRemplDoc(ByVal v_numutil As Long, _
                                             ByVal v_numdoc As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    If v_numdoc = 0 Then
        sql = "select count(*) from Document" _
            & " where (D_UNumResp=" & v_numutil _
            & " or D_LstResp like '%U" & v_numutil & ";1;%')" _
            & " and D_DONum=" & p_NumDocs
    Else
        sql = "select count(*) from Document" _
            & " where D_Num=" & v_numdoc _
            & " and (D_UNumResp=" & v_numutil _
            & " or D_LstResp like '%U" & v_numutil & ";1;%')"
    End If
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        P_UtilEstRespPrincOuRemplDoc = False
        Exit Function
    End If
        
    If lnb > 0 Then
        P_UtilEstRespPrincOuRemplDoc = True
    Else
        P_UtilEstRespPrincOuRemplDoc = False
    End If
    
End Function

Public Function P_UtilEstSuperviseur(ByVal v_numutil As Long, _
                                     ByVal v_numdocs As Long, _
                                     ByRef r_idroit As Integer) As Boolean

    Dim sql As String, slsts As String, s As String
    Dim I As Integer, n As Integer
    
    sql = "select DO_LstSuperv from Documentation" _
        & " where DO_Num=" & v_numdocs
    If Odbc_RecupVal(sql, slsts) = P_ERREUR Then
        P_UtilEstSuperviseur = False
        Exit Function
    End If
    
    n = STR_GetNbchamp(slsts, "|")
    For I = 0 To n - 1
        s = STR_GetChamp(slsts, "|", I)
        ' L'utilisateur est superviseur
        If CLng(Mid$(STR_GetChamp(s, ";", 0), 2)) = v_numutil Then
            r_idroit = STR_GetChamp(s, ";", 2)
            P_UtilEstSuperviseur = True
            Exit Function
        End If
    Next I
    
    P_UtilEstSuperviseur = False

End Function

Public Function P_UtilEstUnRespDoc(ByVal v_numutil As Long, _
                                   ByVal v_numdoc As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    If v_numdoc = 0 Then
        sql = "select count(*) from Document" _
            & " where (D_UNumResp=" & v_numutil & " or D_LstResp like '%U" & v_numutil & ";%')" _
            & " and D_DONum=" & p_NumDocs
    Else
        sql = "select count(*) from Document" _
            & " where D_Num=" & v_numdoc _
            & " and (D_UNumResp=" & v_numutil & " or D_LstResp like '%U" & v_numutil & ";%')"
    End If
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        P_UtilEstUnRespDoc = False
        Exit Function
    End If
        
    If lnb > 0 Then
        P_UtilEstUnRespDoc = True
    Else
        P_UtilEstUnRespDoc = False
    End If
    
End Function

Public Function P_UtilPeutModifierDoc(ByVal v_numdoc As Long, _
                                  ByVal v_numvers As Long, _
                                  ByVal v_numutil As Long) As Boolean

    Dim sql As String
    Dim est_resp As Boolean, bmodresp_autor As Boolean, est_superv As Boolean
    Dim I As Integer, idroit As Integer
    Dim cyordre As Long, cyordre_action As Long, numnat As Long
    Dim rs As rdoResultset
    
    ' Recherche la version en cours
    If v_numvers = 0 Then
        If Odbc_RecupVal("select D_NumVers, D_NDNum from Document where D_Num=" & v_numdoc, _
                          v_numvers, _
                          numnat) = P_ERREUR Then
            P_UtilPeutModifierDoc = False
            Exit Function
        End If
    Else
        If Odbc_RecupVal("select D_NDNum from Document where D_Num=" & v_numdoc, _
                          numnat) = P_ERREUR Then
            P_UtilPeutModifierDoc = False
            Exit Function
        End If
    End If
    
    ' Récupère l'état du document
    If Odbc_Select("select DV_CYOrdre, DV_UNumLock" _
                        & " from DocVersion" _
                        & " where DV_DNum=" & v_numdoc _
                        & " and DV_NumVers=" & v_numvers, _
                     rs) = P_ERREUR Then
        P_UtilPeutModifierDoc = False
        Exit Function
    End If
    
    ' Document locké par un autre utilisateur
    If Not IsNull(rs("DV_UNumLock").Value) And rs("DV_UNumLock").Value <> 0 And rs("DV_UNumLock").Value <> v_numutil Then
        rs.Close
        P_UtilPeutModifierDoc = False
        Exit Function
    End If
        
    est_resp = P_UtilEstRespPrincOuRemplDoc(v_numutil, v_numdoc)
    If Not est_resp Then
        est_superv = P_UtilEstSuperviseur(v_numutil, p_NumDocs, idroit)
        If est_superv And idroit = 2 Then
            est_resp = True
        End If
    End If
    cyordre = rs("DV_CYOrdre").Value
    rs.Close
    ' Document périmé, personne ne peut modifier
    If cyordre > p_cycle_consultable Then
        P_UtilPeutModifierDoc = False
        Exit Function
    End If
    
    ' Document applicable
    If cyordre = p_cycle_consultable Then
        ' Le responsable a le droit de modifier un document applicable ?
        If est_resp Then
            If Odbc_RecupVal("select DON_ModifRespAutor from DocsNature" _
                                & " where DON_DONum=" & p_NumDocs _
                                & " and DON_NDNum=" & numnat, _
                             bmodresp_autor) = P_ERREUR Then
                P_UtilPeutModifierDoc = False
                Exit Function
            End If
            P_UtilPeutModifierDoc = bmodresp_autor
        ' Les autres : jamais
        Else
            P_UtilPeutModifierDoc = False
        End If
        Exit Function
    End If
    
    ' Récupère l'action
    sql = "select DAC_CYOrdre, DAC_UNum, DAC_UNumModif" _
        & " from Docaction" _
        & " where DAC_DNum=" & v_numdoc
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_UtilPeutModifierDoc = False
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        P_UtilPeutModifierDoc = False
        Exit Function
    End If
    While Not rs.EOF
        cyordre_action = rs("DAC_CYOrdre").Value
        ' Cycle modifiable
        If p_scycledocs(cyordre_action).modifiable Then
            ' Le resp peut modifier
            If est_resp Then
                rs.Close
                P_UtilPeutModifierDoc = True
                Exit Function
            End If
            ' L'acteur peut modifier
            If rs("DAC_UNum").Value = v_numutil Or rs("DAC_UNumModif").Value = v_numutil Then
                rs.Close
                P_UtilPeutModifierDoc = True
                Exit Function
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    P_UtilPeutModifierDoc = False
    
End Function



