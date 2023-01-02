Attribute VB_Name = "KD_Fichier"
Option Explicit

Public Function KF_EffacerRepertoire(ByVal v_nomrep As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_effacer_repertoire(v_nomrep, liberr)
    If cr = HTTP_OK Then
        KF_EffacerRepertoire = P_OK
    Else
        KF_EffacerRepertoire = P_ERREUR
    End If
    
End Function

Public Function KF_EstRepertoire(ByVal v_nomrep As String, _
                                 ByVal v_bmess As Boolean) As Boolean

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_est_repertoire(v_nomrep, v_bmess, liberr)
    If cr = HTTP_OK Then
        KF_EstRepertoire = True
    Else
        If v_bmess Then
            Call MsgBox("Impossible d'accéder au répertoire " & v_nomrep & vbCrLf & liberr, vbInformation + vbOKOnly, "")
        End If
        KF_EstRepertoire = False
    End If
    
End Function

Public Function KF_FichierExiste(ByVal v_nomfich As String) As Boolean

    Dim bex As Boolean
    Dim sql As String, liberr As String, nomfich As String
    Dim iret As Integer
    
    iret = HTTP_Appel_fichier_existe(v_nomfich, False, liberr)
    If iret = HTTP_OK Then
        KF_FichierExiste = True
    Else
        KF_FichierExiste = False
    End If
    
End Function

Public Function KF_CopierFichier(ByVal v_nomfich_src As String, _
                                 ByVal v_nomfich_dest As String) As Integer


    Dim sql As String, nomfichd As String, nomfichs As String, liberr As String
    Dim cr As Integer
    Dim lcr As Long
    
    cr = HTTP_Appel_copyfile(v_nomfich_src, v_nomfich_dest, liberr)
    If cr < 0 Then
        Call MsgBox("Impossible de copier " & v_nomfich_src & " dans " & v_nomfich_dest & vbCrLf & vbCrLf & liberr, vbInformation + vbOKOnly, "KF_CopierFichier (HTTP)")
        KF_CopierFichier = P_ERREUR
        Exit Function
    End If
    
    KF_CopierFichier = P_OK
    
End Function

' Copie le fichier v_nomfich_src dans v_nomfich_dest en marquant des poses ...
Public Sub KF_CopierFichierLentement(ByVal v_nomfich_src As String, _
                                    ByVal v_nomfich_dest As String)

    Dim ssys As String
    
    ssys = p_chemin_appli + "\CopieLente.exe " + v_nomfich_src + ";" + v_nomfich_dest
    Call SYS_ExecShell(ssys, False, False)
    
End Sub

Public Function KF_CopierRepertoire(ByVal v_nomrep_src As String, _
                                    ByVal v_nomrep_dest As String) As Integer

    Dim sql As String, nomfichd As String, nomfichs As String, liberr As String
    Dim nomdos As String, cheminsrc As String, chemindest As String
    Dim op As String
    Dim I As Integer
    Dim cr As Integer
    Dim lcr As Long
    
    nomdos = STR_GetChamp(v_nomrep_dest, "/", STR_GetNbchamp(v_nomrep_dest, "/") - 1)
    cheminsrc = Mid(v_nomrep_src, 1, Len(v_nomrep_src) - Len(nomdos) - 1)
    chemindest = Mid(v_nomrep_dest, 1, Len(v_nomrep_dest) - Len(nomdos) - 1)
    cr = HTTP_Appel_GetDos(nomdos, chemindest, cheminsrc, liberr)
    If cr < 0 Then
        Call MsgBox("Impossible de copier " & v_nomrep_src & " dans " & v_nomrep_dest & vbCrLf & vbCrLf & liberr, vbInformation + vbOKOnly, "KF_CopierRepertoire (HTTP)")
        KF_CopierRepertoire = P_ERREUR
        Exit Function
    End If
    
    KF_CopierRepertoire = P_OK
    
End Function

Public Function KF_CreerRepertoire(ByVal v_nomrep As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_creer_repertoire(v_nomrep, liberr)
    If cr = HTTP_OK Then
        KF_CreerRepertoire = P_OK
    Else
        KF_CreerRepertoire = P_ERREUR
    End If
    
End Function

Public Function KF_DelockerDoc(ByVal v_nomdoc As String) As Integer

    Dim nomdoc_lock As String, sql As String, ext As String, liberr As String
    Dim lcr As Long
    
    lcr = HTTP_Appel_LockDelock_file(False, v_nomdoc, False, liberr)

    KF_DelockerDoc = P_OK
    
End Function

Public Function KF_EffacerMod(ByVal v_nomdoc As String) As Integer

    Dim nomdoc As String, sql As String, ext As String, liberr As String
    Dim lcr As Long
    
    ext = Mid$(v_nomdoc, InStrRev(v_nomdoc, ".") + 1)
    nomdoc = left$(v_nomdoc, InStrRev(v_nomdoc, ".")) & "mod_" & ext & "_" & p_NumUtil
    If HTTP_Appel_EffacerFichier(nomdoc, False, liberr) <> HTTP_OK Then
        KF_EffacerMod = P_ERREUR
        Exit Function
    End If

    KF_EffacerMod = P_OK
    
End Function

Public Function KF_EffacerFichier(ByVal v_nomdoc As String, _
                                 ByVal v_bmesserr As Boolean) As Integer

    Dim sql As String
    Dim lcr As Long
    Dim liberr As String
    
    If HTTP_Appel_EffacerFichier(v_nomdoc, v_bmesserr, liberr) <> HTTP_OK Then
        KF_EffacerFichier = P_ERREUR
        Exit Function
    End If

    KF_EffacerFichier = P_OK
    
End Function

Public Function KF_GetFichier(ByVal v_nomfich_srv As String, _
                              ByVal v_nomfich_loc As String) As Integer

    Dim liberr As String, nomfich As String
    Dim pos As Integer
    
    Call FICH_EffacerFichier(v_nomfich_loc, False)
    
    If HTTP_Appel_GetFile(v_nomfich_srv, v_nomfich_loc, False, False, liberr) <> HTTP_OK Then
        Call MsgBox("Impossible de rapatrier " & v_nomfich_srv & " dans " & v_nomfich_loc & vbCrLf & "Erreur : " & liberr, vbInformation + vbOKOnly, "KF_GetFichier (HTTP serv vers loc)")
        KF_GetFichier = P_ERREUR
        Exit Function
    End If

    KF_GetFichier = P_OK

End Function

Public Function KF_PutFichier(ByVal v_nomfich_srv As String, _
                              ByVal v_nomfich_loc As String) As Integer
    
    Dim liberr As String, nomfich As String
    Dim pos As Integer
    
    If HTTP_Appel_PutFile(v_nomfich_srv, v_nomfich_loc, False, False, liberr) <> HTTP_OK Then
        Call MsgBox("Impossible de transférer " & v_nomfich_loc & " dans " & v_nomfich_srv & vbCrLf & "Erreur : " & liberr, vbInformation + vbOKOnly, "KF_PutFichier (HTTP loc vers serv)")
        KF_PutFichier = P_ERREUR
        Exit Function
    End If

    KF_PutFichier = P_OK

End Function

Public Function KF_RenommerFichier(ByVal v_nomsrc As String, _
                                   ByVal v_nomdest As String) As Integer

    Dim sql As String, liberr As String
    Dim cr As Integer
    Dim lcr As Long
    
    cr = HTTP_Appel_renamefile(v_nomsrc, v_nomdest, liberr)
    If cr < 0 Then
        Call MsgBox("Impossible de renommer " & v_nomsrc & " en " & v_nomdest & vbCrLf & vbCrLf & liberr, vbInformation + vbOKOnly, "KF_RenommerFichier (HTTP)")
        KF_RenommerFichier = P_ERREUR
        Exit Function
    End If
    
    KF_RenommerFichier = P_OK

End Function

