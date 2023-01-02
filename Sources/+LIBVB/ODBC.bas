Attribute VB_Name = "Modbc"
Option Explicit

' Type de base
Public Odbc_type_base As Integer
Public Const ODBC_BDD_MDB = 1
Public Const ODBC_BDD_PG = 2

Public Odbc_nberr As Long

Public Odbc_ev As rdoEnvironment
Public Odbc_Cnx As rdoConnection

Public odbc_trans_encours As Boolean

Public Function Odbc_AddNew(ByVal v_nomtbl As String, _
                            ByVal v_nomcol0 As String, _
                            ByVal v_nomseq As String, _
                            ByVal v_brecupcle As Boolean, _
                            ByRef r_scle As Variant, _
                            ParamArray v_tval() As Variant) As Integer

    Dim sql As String, scol As String, scol_update As String
    Dim n As Integer, pos As Integer
    Dim lnum As Long
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim sTooLong As String
    Dim sqlTooLong As String
    Dim I As Integer
    Dim opTooLong As String
    Dim retODBC As Integer
    Dim retTooLong As Boolean
    Dim s_chp As String, s_val As String
Lab_Again:
    If Odbc_type_base = ODBC_BDD_MDB Then
        scol_update = ""
        sql = "select * from " & v_nomtbl _
            & " where " & v_nomcol0 & "=0"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo err_addnew
        rs.AddNew
        On Error GoTo 0
        n = 0
        For Each val In v_tval
            If n Mod 2 = 0 Then
                scol = val
            Else
                If IsNull(val) Then GoTo lab_affecte
                On Error GoTo pas_une_string
                pos = InStr(val, "%%NEXTVAL")
                If pos > 0 Then
                    scol_update = scol
                    val = Mid$(val, pos + 10)
                End If
lab_affecte:
                On Error GoTo err_affecte
                rs(scol).Value = val
            End If
            n = n + 1
        Next val
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        If v_brecupcle Then
            rs.MoveLast
            r_scle = rs(0).Value
            If scol_update <> "" Then
                On Error GoTo err_edit
                rs.Edit
                On Error GoTo err_affecte
                rs(scol_update).Value = r_scle
                On Error GoTo err_update
                rs.Update
            End If
        End If
        On Error GoTo 0
        rs.Close
    Else
lab_debut:
        sql = "select nextval('" & v_nomseq & "')"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
        If rs.EOF Then GoTo err_no_resultset
        lnum = rs(0).Value
        rs.Close
        sql = "select * from " & v_nomtbl _
            & " where " & v_nomcol0 & "=0"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo err_addnew
        rs.AddNew
        On Error GoTo err_affecte
        rs(0).Value = lnum
        sqlTooLong = "select * from " & v_nomtbl & " where " & v_nomcol0 & "=" & lnum
        n = 0
        For Each val In v_tval
            If n Mod 2 = 0 Then
                scol = val
            Else
                If IsNull(val) Then
                    pos = 0
                    GoTo lab_affecte2
                End If
                On Error GoTo pas_une_string2
                pos = InStr(val, "%%NEXTVAL")
lab_affecte2:
                On Error GoTo err_affecte
                If pos > 0 Then
                    rs(scol).Value = lnum
                Else
                    rs(scol).Value = val
                End If
            End If
            n = n + 1
        Next val
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
        r_scle = lnum
    End If
    
    Odbc_AddNew = P_OK
    Exit Function
        
pas_une_string:
    GoTo lab_affecte
    
pas_une_string2:
    pos = 0
    GoTo lab_affecte2
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultset " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    Odbc_AddNew = P_ERREUR
    Exit Function

err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " & sql, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    rs.Close
    Odbc_AddNew = P_ERREUR
    Exit Function

err_addnew:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur AddNew " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    rs.Close
    Odbc_AddNew = P_ERREUR
    Exit Function

err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    rs.Close
    Odbc_AddNew = P_ERREUR
    Exit Function

err_affecte:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Affectation pour " & scol & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    rs.Close
    Odbc_AddNew = P_ERREUR
    Exit Function

err_update:
    
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Update " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    rs.Close
    Odbc_AddNew = P_ERREUR
    Exit Function

End Function


Public Function Odbc_BeginTrans() As Integer

    If odbc_trans_encours Then
        MsgBox "Une transaction est déjà en cours", vbOKOnly + vbCritical, "MOdbc (Odbc_BeginTrans)"
        Odbc_BeginTrans = P_ERREUR
        Exit Function
    End If
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_begintrans
    Odbc_Cnx.BeginTrans
    On Error GoTo 0
    
    odbc_trans_encours = True
    
    Odbc_BeginTrans = P_OK
    Exit Function
    
err_begintrans:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur BeginTrans" & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_BeginTrans)"
    Odbc_BeginTrans = P_ERREUR
    Exit Function

End Function

Public Function Odbc_Bool(ByVal v_bool As Boolean) As String

'    If Odbc_type_base = ODBC_BDD_PG Then
        Odbc_Bool = IIf(v_bool, "true", "false")
'    Else
'    End If
    
End Function

Public Sub Odbc_Close()

    Odbc_Cnx.Close
    
End Sub

Public Function Odbc_CommitTrans() As Integer

    If Not odbc_trans_encours Then
        MsgBox "Pas de transaction en cours", vbOKOnly + vbCritical, "MOdbc (Odbc_CommitTrans)"
        Odbc_CommitTrans = P_ERREUR
        Exit Function
    End If
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_committrans
    Odbc_Cnx.CommitTrans
    On Error GoTo 0
    
    odbc_trans_encours = False
    
    Odbc_CommitTrans = P_OK
    Exit Function
    
err_committrans:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur CommitTrans", vbOKOnly + vbCritical, "MOdbc (Odbc_CommitTrans)"
    Odbc_CommitTrans = P_ERREUR
    Exit Function

End Function

Public Function Odbc_Count(ByVal v_sql As String, _
                            ByRef r_count As Long, _
                            Optional v_indrs As Variant) As Integer

    Dim ind As Integer
    Dim rs As rdoResultset
    
    If IsMissing(v_indrs) Then
        ind = 0
    Else
        ind = v_indrs
    End If
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenStatic)
    On Error GoTo 0
    If rs.EOF Then
        r_count = 0
    ElseIf IsNull(rs(ind).Value) Then
        r_count = 0
    Else
        r_count = rs(ind).Value
    End If
    rs.Close
    
    Odbc_Count = P_OK
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + v_sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Count)"
    Odbc_Count = P_ERREUR
    Exit Function

End Function

Public Function Odbc_CreateTable(ByVal v_nomtbl As String, _
                                 ParamArray v_tval() As Variant) As Integer

    Dim sql As String
    Dim n As Integer, I As Integer, lg As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    sql = "create table " & v_nomtbl & " ("
    n = 0
    For Each val In v_tval
        If n Mod 2 = 0 Then
            If n > 0 Then
                sql = sql + ", "
            End If
            sql = sql + val + " "
        Else
            Select Case val
            Case "int4"
                If Odbc_type_base = ODBC_BDD_MDB Then
                    val = "long"
                End If
            Case "int2"
                If Odbc_type_base = ODBC_BDD_MDB Then
                    val = "short"
                End If
            Case "bool"
                If Odbc_type_base = ODBC_BDD_MDB Then
                    val = "bit"
                End If
            Case Else
                If left$(val, 3) = "str" Then
                    lg = Mid$(val, 4)
                    If Odbc_type_base = ODBC_BDD_PG Then
                        val = "varchar(" & lg & ")"
                    Else
                        val = "text(" & lg & ") not null"
                    End If
                End If
            End Select
            sql = sql + val
        End If
        n = n + 1
    Next val
    sql = sql + " )"
    
    On Error GoTo err_execute
    Odbc_Cnx.Execute (sql)
    On Error GoTo 0
    
    Odbc_CreateTable = P_OK
    Exit Function
        
err_execute:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Create Table " & sql & vbCrLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_CreateTable)"
    Odbc_CreateTable = P_ERREUR
    Exit Function

End Function

Public Function Odbc_CreateTableOnly(ByVal v_nomtbl As String, _
                                     ByVal v_nomcol0 As String) As Integer

    Dim sql As String
    Dim n As Integer, I As Integer, lg As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    sql = "create table " & v_nomtbl & " (" & v_nomcol0 & ")"
    
    On Error GoTo err_execute
    Odbc_Cnx.Execute (sql)
    On Error GoTo 0
    
    Odbc_CreateTableOnly = P_OK
    Exit Function
        
err_execute:
    MsgBox "Erreur Create Table Only" & sql & vbCrLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_CreateTableOnly)"
    Odbc_CreateTableOnly = P_ERREUR
    Exit Function

End Function

Public Function Odbc_AddColumn(ByVal v_nomtbl As String, _
                                ByVal v_nomcol As String) As Integer

    Dim sql As String
    
    sql = "alter table " & v_nomtbl & " add column " & v_nomcol
    
    On Error GoTo err_execute
    Odbc_Cnx.Execute (sql)
    On Error GoTo 0
    
    Odbc_AddColumn = P_OK
    Exit Function
        
err_execute:
    MsgBox "Erreur Add Column" & sql & vbCrLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddColumn)"
    Odbc_AddColumn = P_ERREUR
    Exit Function

End Function

'Convertit la date française sdate en chaine date ODBC
'qui doit être sous la forme {d 'aaaa-mm-dd'}
Public Function Odbc_Date(ByVal v_date As Date) As String

    Odbc_Date = "{d '" & Format(v_date, "yyyy-mm-dd") & "'}"
    
End Function

Public Function Odbc_Delete(ByVal v_nomtbl As String, _
                            ByVal v_nomcol0 As String, _
                            ByVal v_sclause As String, _
                            ByRef r_lnb As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    sql = "select " & v_nomcol0 & " from " + v_nomtbl + " " + v_sclause
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    On Error GoTo 0
    r_lnb = 0
    While Not rs.EOF
        r_lnb = r_lnb + 1
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_delete
        rs.Delete
        On Error GoTo 0
        rs.MoveNext
    Wend
    rs.Close
    
    If r_lnb = 0 Then
        Odbc_Delete = P_NON
    Else
        Odbc_Delete = P_OUI
    End If
    Exit Function
        
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultset " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delete)"
    Odbc_Delete = P_ERREUR
    Exit Function

err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delete)"
    rs.Close
    Odbc_Delete = P_ERREUR
    Exit Function

err_delete:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Delete " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delete)"
    rs.Close
    Odbc_Delete = P_ERREUR
    Exit Function

End Function

Public Sub Odbc_Delock(ByVal v_nomtbl As String, _
                        ByVal v_scols As String, _
                        ByVal v_scond As String)
                        
    Dim sql As String
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    sql = "select " & v_scols & " from " & v_nomtbl & " " & v_scond
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    On Error GoTo err_edit
    rs.Edit
    On Error GoTo err_affecte
    rs(0).Value = 0
    On Error GoTo err_update
    rs.Update
    On Error GoTo 0
    rs.Close
    
    Exit Sub
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delock)"
    Exit Sub
    
err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delock)"
    rs.Close
    Exit Sub
    
err_affecte:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Affectation pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delock)"
    rs.Close
    Exit Sub
    
err_update:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Update pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Delock)"
    rs.Close
    Exit Sub
    
End Sub

Public Function Odbc_Erreur_Value_Too_Long(ByVal v_sql As String, _
                                            ByVal v_chp As String, ByVal v_val As Variant) As Boolean

    Dim rs As rdoResultset
    Dim n As Integer, I As Integer
    Dim val As Variant

    On Error GoTo err_essai_trouver_champ
    Set rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenKeyset, rdConcurRowVer)
    rs.Edit
    rs(v_chp).Value = v_val
    rs.Update
    rs.Close
    Odbc_Erreur_Value_Too_Long = True
    Exit Function
    
err_essai_trouver_champ:
    MsgBox "Erreur de taille sur le champ " & v_chp & Chr(13) & Chr(10) & " SQL = " & v_sql & Chr(13) & Chr(10) & " Taille de la valeur = " & Len(v_val) & Chr(13) & Chr(10) & " Valeur = " & v_val
    Odbc_Erreur_Value_Too_Long = False
End Function

Function Odbc_EstDoublon(ByVal v_nomtbl As String, _
                         ByVal v_nomcol As String, _
                         ByVal v_svalcol As String, _
                         ByVal v_nomcol0 As String, _
                         ByVal v_valcol0 As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    
    sql = "select " & v_nomcol0 & ", " & v_nomcol & " from " & v_nomtbl _
        & " where " & v_nomcol & "=" & Odbc_String(v_svalcol)
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
    If rs.EOF Then
        Odbc_EstDoublon = P_NON
    Else
        If rs(v_nomcol0).Value <> v_valcol0 Then
            Odbc_EstDoublon = P_OUI
        Else
            Odbc_EstDoublon = P_NON
        End If
    End If
    rs.Close
       
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_EstDoublon)"
    Odbc_EstDoublon = P_ERREUR
    Exit Function
    
End Function

Public Function Odbc_Execute_Insert(ByVal v_nomtbl As String, _
                                    ParamArray v_tval() As Variant) As Integer

    Dim sql As String, scol As String, sval As String
    Dim n As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    scol = ""
    sval = ""
    n = 0
    For Each val In v_tval
        If n Mod 2 = 0 Then
            If scol = "" Then
                scol = scol + "("
            Else
                scol = scol + ","
            End If
            scol = scol & val
        Else
            If sval = "" Then
                sval = sval + "("
            Else
                sval = sval + ","
            End If
            sval = sval & val
        End If
        n = n + 1
    Next val
    scol = scol + ")"
    sval = sval + ")"
    sql = "insert into " & v_nomtbl & " " & scol & " values " & sval
    On Error GoTo err_execute
    Call Odbc_Cnx.Execute(sql)
    On Error GoTo 0
    
    Odbc_Execute_Insert = P_OK
    Exit Function
        
err_execute:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Execute " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    Odbc_Execute_Insert = P_ERREUR

End Function

Public Function Odbc_Execute_Update(ByVal v_nomtbl As String, _
                                    ByVal v_scond As String, _
                                    ParamArray v_tval() As Variant) As Integer

    Dim sql As String, sval As String
    Dim n As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    sval = ""
    n = 0
    For Each val In v_tval
        If n Mod 2 = 0 Then
            If sval <> "" Then
                sval = sval + ","
            End If
            sval = sval & val & "="
        Else
            sval = sval & val
        End If
        n = n + 1
    Next val
    sql = "update " & v_nomtbl & " set " & sval & " " & v_scond
    On Error GoTo err_execute
    Call Odbc_Cnx.Execute(sql)
    On Error GoTo 0
    
    Odbc_Execute_Update = P_OK
    Exit Function
        
err_execute:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Execute " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_AddNew)"
    Odbc_Execute_Update = P_ERREUR

End Function

Public Function Odbc_Init(ByVal v_stypebdd As String, _
                          ByVal v_nombdd As String, _
                          ByVal v_mode_verbeux As Boolean) As Integer

    Dim nom_source As String, nom_prm_connex As String
    Dim code As String
    
    Odbc_nberr = 0
    
    If v_stypebdd = "MDB" Then
        Odbc_type_base = ODBC_BDD_MDB
    Else
        Odbc_type_base = ODBC_BDD_PG
    End If

    ' Connexion à la base
    On Error GoTo err_env
    Set Odbc_ev = rdoEngine.rdoEnvironments(0)
    On Error GoTo 0
    If Odbc_type_base = ODBC_BDD_PG Then
        Odbc_ev.CursorDriver = rdUseOdbc
        nom_source = v_nombdd
        nom_prm_connex = ""
    Else
        Odbc_ev.CursorDriver = rdUseServer
        nom_source = ""
        nom_prm_connex = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & v_nombdd
    End If
    On Error GoTo err_connection
    Set Odbc_Cnx = Odbc_ev.OpenConnection(nom_source, Connect:=nom_prm_connex)
    On Error GoTo 0

    odbc_trans_encours = False
    
    Odbc_Init = P_OK
    Exit Function
    
err_env:
    If v_mode_verbeux Then MsgBox "Erreur Environnement " & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Init)"
    Odbc_Init = P_ERREUR
    Exit Function
    
err_connection:
    If v_mode_verbeux Then
        code = Trim(STR_GetChamp(Err.Description, ":", 0))
        If Err.Number = 40002 And code = "08001" Then
            MsgBox "Connexion à la base <" & v_nombdd & "> impossible", vbOKOnly + vbCritical, "MOdbc (Odbc_Init) " & Err.Number
        Else
            MsgBox "Erreur Connexion à " & v_nombdd & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Init)"
        End If
    End If
    Odbc_Init = P_ERREUR
    Exit Function
    
End Function

Public Function Odbc_Lock(ByVal v_nomtbl As String, _
                          ByVal v_scols As String, _
                          ByVal v_scond As String, _
                          ByVal v_numutil As Long, _
                          ByRef r_numutil_lock As Long) As Integer
                        
    Dim sql As String
    Dim num As Long
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    sql = "select " & v_scols & " from " & v_nomtbl & " " & v_scond
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    On Error GoTo 0
    If rs.EOF Then GoTo err_no_resultset
    If rs(0).Value > 0 And rs(0).Value <> v_numutil Then
        r_numutil_lock = rs(0).Value
        rs.Close
        Odbc_Lock = P_NON
        Exit Function
    End If
    On Error GoTo err_edit
    rs.Edit
    On Error GoTo err_affecte
    rs(0).Value = v_numutil
    On Error GoTo err_update
    rs.Update
    On Error GoTo 0
    rs.Close
    
    ' On revérifie
lab_verif:
    sql = "select " & v_scols & " from " & v_nomtbl & " " & v_scond
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
    On Error GoTo 0
    If rs.EOF Then GoTo err_no_resultset
    If rs(0).Value <> v_numutil Then
        r_numutil_lock = rs(0).Value
        rs.Close
        Odbc_Lock = P_NON
        Exit Function
    End If
    rs.Close
    
    Odbc_Lock = P_OUI
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_lock)"
    Odbc_Lock = P_ERREUR
    Exit Function
    
err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " + sql, vbOKOnly + vbCritical, "MOdbc (Odbc_lock)"
    rs.Close
    Odbc_Lock = P_ERREUR
    Exit Function
    
err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_lock)"
    rs.Close
    Odbc_Lock = P_ERREUR
    Exit Function
    
err_affecte:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Affectation pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_lock)"
    rs.Close
    Odbc_Lock = P_ERREUR
    Exit Function
    
err_update:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Update pour " + sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_lock)"
    rs.Close
    GoTo lab_verif
    
End Function

Public Function Odbc_MinMax(ByVal v_sql As String, _
                            ByRef r_lnum As Long) As Integer

    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenStatic)
    On Error GoTo 0
    If rs.EOF Then
        r_lnum = 0
    ElseIf IsNull(rs(0).Value) Then
        r_lnum = 0
    Else
        r_lnum = rs(0).Value
    End If
    rs.Close
    
    Odbc_MinMax = P_OK
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + v_sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_MinMax)"
    Odbc_MinMax = P_ERREUR
    Exit Function

End Function

Public Function Odbc_RecupVal(ByVal v_sql As String, _
                              ParamArray r_tval() As Variant) As Integer

    Dim sql As String
    Dim I As Integer
    Dim val As Variant, val2 As Variant
    Dim rs As rdoResultset
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenStatic)
    On Error GoTo 0
    If rs.EOF Then GoTo err_no_resultset
    
    I = 0
    For Each val In r_tval
        On Error GoTo err_no_val
        val2 = rs(I).Value
        If IsNull(val2) Then
            r_tval(I) = ""
        Else
            r_tval(I) = val2
        End If
        On Error GoTo 0
        I = I + 1
    Next val
    rs.Close
    
    Odbc_RecupVal = P_OK
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + v_sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, ""
    Odbc_RecupVal = P_ERREUR
    Exit Function

err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " + v_sql, vbOKOnly + vbCritical, ""
    rs.Close
    Odbc_RecupVal = P_ERREUR
    Exit Function

err_no_val:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de valeur en position " & I & " pour " & v_sql & vbcr & vbLf & Err.Description, vbOKOnly + vbCritical, ""
    rs.Close
    Odbc_RecupVal = P_ERREUR
    Exit Function

End Function

Public Function Odbc_RollbackTrans() As Integer

    If Not odbc_trans_encours Then
        MsgBox "Pas de transaction en cours", vbOKOnly + vbCritical, "MOdbc (odbc_RollbackTrans)"
        Odbc_RollbackTrans = P_ERREUR
        Exit Function
    End If
    
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_rollbacktrans
    Odbc_Cnx.RollbackTrans
    On Error GoTo 0
    
    odbc_trans_encours = False
    
    Odbc_RollbackTrans = P_OK
    Exit Function
    
err_rollbacktrans:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur RollbackTrans", vbOKOnly + vbCritical, "MOdbc (Odbc_RollbackTrans)"
    Odbc_RollbackTrans = P_ERREUR
    Exit Function

End Function

Public Function Odbc_Select(ByVal v_sql As String, _
                            ByRef r_rs As rdoResultset) As Integer

    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_open_resultset
    Set r_rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenStatic)
    On Error GoTo 0
    If r_rs.EOF Then GoTo err_no_resultset

    Odbc_Select = P_OK
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + v_sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Select)"
    Odbc_Select = P_ERREUR
    Exit Function

err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " + v_sql, vbOKOnly + vbCritical, "MOdbc (Odbc_Select)"
    r_rs.Close
    Odbc_Select = P_ERREUR
    Exit Function

End Function

Public Function Odbc_SelectV(ByVal v_sql As String, _
                             ByRef r_rs As rdoResultset) As Integer
    Dim retODBC As Integer
Lab_Again:
    On Error GoTo err_open_resultset
    Set r_rs = Odbc_Cnx.OpenResultset(v_sql, rdOpenStatic)
    On Error GoTo 0

    Odbc_SelectV = P_OK
    Exit Function
    
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultSet pour " + v_sql & vbLf & vbcr & "Erreur=" & Err.Number & vbLf & vbcr & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_SelectV)"
    Odbc_SelectV = P_ERREUR
    Exit Function

End Function


Private Function Odbc_Voir_Cnx(ByVal v_errNumber As Long, ByVal v_errDescription As String) As Boolean
    Dim mess As String
    Dim smess As String
    Dim code As String
    Dim bVoir As Boolean
    
    Odbc_Voir_Cnx = False
    If v_errNumber = 40011 Or v_errNumber = 40002 Then
        ' Essayer un fois de reconnecter
        code = Trim(STR_GetChamp(v_errDescription, ":", 0))
        If v_errNumber = 40002 And code = "08001" Then
            bVoir = True
        ElseIf v_errNumber = 40002 And code = "08S01" Then
            bVoir = True
        End If
        
        If bVoir Then
            mess = "La connexion à la base de données " & p_Nom_BDD & " a été interrompue" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            mess = mess & "             (Procédure de sauvegarde ?)" & Chr(13) & Chr(10)
            If odbc_trans_encours Then
                MsgBox mess & Chr(13) & Chr(10) & "une transaction étant en cours, la reconnexion automatique est impossible (" & v_errNumber & ")", vbCritical
                End
            Else
                smess = mess & Chr(13) & Chr(10) & "la reconnexion automatique est impossible (" & v_errNumber & ")"
lab_init:
                If Odbc_Init(p_Type_BDD, p_Nom_BDD, False) = P_ERREUR Then
                    'MsgBox mess & Chr(13) & Chr(10) & "la reconnexion automatique est impossible (" & v_errNumber & ")", vbCritical
                    If MsgBox(smess & Chr(13) & Chr(10) & "Re-tenter la connexion à " & p_Nom_BDD & " ?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                        GoTo lab_init
                    Else
                        End
                    End If
                Else
                    MsgBox mess & Chr(13) & Chr(10) & "la reconnexion automatique a été effectuée", vbInformation
                    Odbc_Voir_Cnx = True
                End If
            End If
        Else
            MsgBox "Odbc_Voir_Cnx : " & v_errNumber & " " & v_errDescription
        End If
    End If
End Function

'Rajoute les ' en début et fin de chaine
'Transforme les ' en ''
'Transforme les * en %
Public Function Odbc_String(ByVal v_str As String) As String

    Dim s As String
    
    s = v_str
    s = Replace(s, "*", "%")
    s = Replace(s, "'", "''")
    Odbc_String = "'" & s & "'"
    
End Function

Public Function Odbc_StringJoker(ByVal v_str As String) As String

    Dim s As String
    
    s = v_str
    If Odbc_type_base <> ODBC_BDD_MDB Then
        ' _ = Joker un caractère
        s = Replace(s, "_", "\\_")
        ' % = Joker plusieurs caractères
        s = Replace(s, "%", "\\%")
    End If
    s = Replace(s, "*", "%")
    s = Replace(s, "'", "''")
    Odbc_StringJoker = "'" & s & "'"
    
End Function

Public Function Odbc_TableExiste(ByVal v_nomtbl As String) As Boolean

    Dim sql As String
    Dim lnb As Long
    Dim tbl As rdoTable
    
    If Odbc_type_base = ODBC_BDD_MDB Then
        For Each tbl In Odbc_Cnx.rdoTables
            If tbl.Name = v_nomtbl Then
                Odbc_TableExiste = True
                Exit Function
            End If
        Next tbl
        Odbc_TableExiste = False
    Else
        sql = "select count(*) from pg_tables where tablename='" & v_nomtbl & "'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            lnb = 0
        End If
        If lnb = 0 Then
            Odbc_TableExiste = False
        Else
            Odbc_TableExiste = True
        End If
    End If

End Function

Public Function Odbc_Update(ByVal v_nomtbl As String, _
                            ByVal v_nomcol0 As String, _
                            ByVal v_scond As String, _
                            ParamArray v_tval() As Variant) As Integer

    Dim sql As String
    Dim n As Integer, I As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim sTooLong As String
    Dim opTooLong As String
    Dim retODBC As Integer
    Dim retTooLong As Boolean
    Dim s_chp As String, s_val As String
Lab_Again:
    sql = "select " & v_nomcol0
    n = 0
    For Each val In v_tval
        If n Mod 2 = 0 Then sql = sql + ", " + val
        n = n + 1
    Next val
    sql = sql + " from " + v_nomtbl + " " + v_scond
    
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    If rs.EOF Then GoTo err_no_resultset
    On Error GoTo err_edit
    rs.Edit
    On Error GoTo err_affecte
    I = 1
    n = 0
    For Each val In v_tval
        If n Mod 2 = 1 Then
            rs(I).Value = val
            I = I + 1
        End If
        n = n + 1
    Next val
    On Error GoTo err_update
    rs.Update
    On Error GoTo 0
    rs.Close
    
    Odbc_Update = P_OK
    Exit Function
        
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultset " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Update)"
    Odbc_Update = P_ERREUR
    Exit Function

err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " & sql, vbOKOnly + vbCritical, "MOdbc (Odbc_Update)"
    rs.Close
    Odbc_Update = P_ERREUR
    Exit Function

err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Update)"
    rs.Close
    Odbc_Update = P_ERREUR
    Exit Function

err_affecte:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Affectation colonne " & n & " pour " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Update)"
    rs.Close
    Odbc_Update = P_ERREUR
    Exit Function

err_update:
    If Err.Number = 40002 And InStr(Err.Description, "S1000") > 0 Then
        ' Essayer de savoir quel est le champ qui déconne
        I = 1
        n = 0
        For Each val In v_tval
            If n Mod 2 = 1 Then
                s_chp = rs(I).Name
                retTooLong = Odbc_Erreur_Value_Too_Long(sql, s_chp, val)
                If Not retTooLong Then
                    Exit For
                End If
                rs(I).Value = val
                I = I + 1
            End If
            n = n + 1
        Next val
        Odbc_Update = P_ERREUR
        Exit Function
    End If
    
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Update " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_Update)"
    rs.Close
    Odbc_Update = P_ERREUR
    Exit Function

End Function

Public Function Odbc_UpdateP(ByVal v_nomtbl As String, _
                            ByVal v_nomcol0 As String, _
                            ByVal v_scond As String, _
                            ByRef r_lnbu As Long, _
                            ParamArray v_tval() As Variant) As Integer

    Dim sql As String
    Dim n As Integer, I As Integer
    Dim val As Variant
    Dim rs As rdoResultset
    
    Dim sTooLong As String
    Dim opTooLong As String
    Dim retODBC As Integer
    Dim retTooLong As Boolean
    Dim s_chp As String, s_val As String
Lab_Again:
    sql = "select " & v_nomcol0
    n = 0
    For Each val In v_tval
        If n Mod 2 = 0 Then sql = sql + ", " + val
        n = n + 1
    Next val
    sql = sql + " from " + v_nomtbl + " " + v_scond
    
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
    r_lnbu = 0
    While Not rs.EOF
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
        I = 1
        n = 0
        For Each val In v_tval
            If n Mod 2 = 1 Then
                rs(I).Value = val
                I = I + 1
            End If
            n = n + 1
        Next val
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        r_lnbu = r_lnbu + 1
        rs.MoveNext
    Wend
    rs.Close
    
    Odbc_UpdateP = P_OK
    Exit Function
        
err_open_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur OpenResultset " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_UpdateP)"
    Odbc_UpdateP = P_ERREUR
    Exit Function

err_no_resultset:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Pas de ligne pour " & sql, vbOKOnly + vbCritical, "MOdbc (Odbc_UpdateP)"
    rs.Close
    Odbc_UpdateP = P_ERREUR
    Exit Function

err_edit:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Edit " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_UpdateP)"
    rs.Close
    Odbc_UpdateP = P_ERREUR
    Exit Function

err_affecte:
    If Err.Number = 40002 And InStr(Err.Description, "S1000") > 0 Then
        ' Essayer de savoir quel est le champ qui déconne
        I = 1
        n = 0
        For Each val In v_tval
            If n Mod 2 = 1 Then
                s_chp = rs(I).Name
                retTooLong = Odbc_Erreur_Value_Too_Long(sql, s_chp, val)
                If Not retTooLong Then
                    Exit For
                End If
                rs(I).Value = val
                I = I + 1
            End If
            n = n + 1
        Next val
        Odbc_UpdateP = P_ERREUR
        Exit Function
    End If
    
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Affectation colonne " & n & " pour " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_UpdateP)"
    rs.Close
    Odbc_UpdateP = P_ERREUR
    Exit Function

err_update:
    retODBC = Odbc_Voir_Cnx(Err.Number, Err.Description)
    If retODBC Then Resume Lab_Again
    MsgBox "Erreur Update " & sql & vbcr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "MOdbc (Odbc_UpdateP)"
    rs.Close
    Odbc_UpdateP = P_ERREUR
    Exit Function

End Function

Public Function Odbc_upper() As String

    If Odbc_type_base = ODBC_BDD_MDB Then
        Odbc_upper = "ucase"
    Else
        Odbc_upper = "upper"
    End If
    
End Function
