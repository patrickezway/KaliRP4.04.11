Attribute VB_Name = "MLog"
Option Explicit

' Description du log
' Le fichier log se trouve sur le dossier "p_chemin_appli"\logs\NOMAPPLI.log
' ou sur APPDATA\Kalidoc\log\NOMAPPLI.log dans le cas où APPDATA n'est pas dans C:\Users\... (certains TSE)

' A chaque Message est attribué un niveau numérique
' correspondant à une des constantes mnémotechiques suivantes:
' VALEUR                        ' INTERPRETATION PROPOSEE
Public Const LOG_NONE = 0       ' Base (valeur plus petite que la plus petite constante)
Public Const LOG_ERROR = 1      ' Erreur fatale = arrêt du programme
Public Const LOG_WARNING = 2    ' Erreur récuperable = traitement non effectué mais le programme peut continuer
Public Const LOG_INFO = 3       ' Condition particulière pouvant interresser l'utilisateur
Public Const LOG_DEBUG = 4      ' Condition particulière pouvant interresser le developpeur
Public Const LOG_DEBUG5 = 5     ' Condition particulière pouvant interresser le développeur dans certaines circonstances

' pour trace operation
Public p_FichierOrigine As String
Public p_FichierImport As String
Public p_IdentOrigine As String
Public p_ActionOrigine As String
Public p_ActionResume As String
Public p_IP_Locale As String
Public p_HostName As String

Public P_Nom_Log As String

' Chaque message est exprimé sous forme de MsgBox et/ou de ligne dans le fichier log et/ou dans la table log.
' Pour que le message sorte dans un canal donné son niveau doit être >= au niveau Maximum de ce canal.
' Ce niveau Maximum est indiqué par les variables "Public" déclarées ci-dessous et qui doivent être affectées
' dans le programme principal. Voir aussi dans la fonction log_config les cas d'affectation automatique de
' ces variables à partir de l'environement (LOG_FICHIER par exemple) ou de la base de données ou du .INI.
Public LOG_MsgBoxes As Long     ' Valeur Maximum pour ouvrir MsgBox
Public LOG_Fichier As Long      ' Valeur Maximum pour ecrire dans le fichier log
Public LOG_Table As Long        ' Valeur Maximum pour ecrire dans la table log
Public LOG_MaxSize As Long
Public LOG_BDD_ON As Boolean

Public logPath As String
Private logMEM As String
Private config_lue As Boolean

' Affecte la variable globale logPath si elle est vide
' retourne false si le fichier n'existe pas et ne peut pas être créé, true sinon
Public Function log_getPath() As Boolean

Dim fso As Object
Dim errMsg As String

    If Len(logPath) > 0 Then
        log_getPath = True
        Exit Function
    End If
    On Error GoTo err_handler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If App.EXEName = "KaliScrute" Or App.EXEName = "KaliScrute_Lance" Then  ' on le force à rester dans chemin_appli, sinon => Appdata\roaming ...
        logPath = p_chemin_appli
    ElseIf left(Environ$("APPDATA"), 8) = "C:\Users" Then
        logPath = p_chemin_appli  ' <> App.path
    Else
        logPath = Environ$("APPDATA") & "\Kalidoc"
        If Not fso.FolderExists(logPath) Then
            fso.CreateFolder (logPath)
        End If
    End If
    logPath = logPath & "\logs\"
   
    If Not fso.FolderExists(logPath) Then
        fso.CreateFolder (logPath)
    End If
    If fso.FolderExists(logPath) Then
        log_getPath = True
        Exit Function
    End If
    Set fso = Nothing
    log_getPath = False
    logPath = vbNullString
    Exit Function
    
err_handler:
    Set fso = Nothing
    If LOG_MsgBoxes > 0 Then
        errMsg = "logPath: " & logPath & ", erreur " & Err.Number & vbCrLf & Err.Description
        MsgBox errMsg, "Erreur dans log_getPath:"
    End If
    logPath = vbNullString
    log_getPath = False
    Exit Function
        
End Function

' configuration des canaux de log en fonction de l'environement de la BDD ou du .INI
Private Sub LOG_config()

    Dim logPoste As String, logAppli As String
    Dim sql As String
    Dim vals As Scripting.Dictionary
    
    logPoste = SYS_GetComputerName()
    logAppli = App.EXEName
    
    'MsgBox "logPoste=" & logPoste & ", logAppli=" & logAppli
    If Environ$("LOG_FICHIER") <> "" Then
        LOG_Fichier = CInt(Environ$("LOG_FICHIER"))
        LOG_Table = 0
    ElseIf Odbc_TableExiste("logconf") Then
        sql = "select max( lc_logfichier ) log_fichier, max( lc_logtable ) log_table " & _
                "from logconf where true " & _
                "and (length(lc_poste)=0) or (? ~* ('^' || lc_poste || '$')) " & _
                "and (length(lc_user)=0) or (? ~* ('^' || lc_user || '$')) " & _
                "and (length(lc_appli)=0) or (? ~* ('^' || lc_appli || '$')) "
            
        Set vals = DB_getDICT(sql, logPoste, 0, logAppli)
        LOG_Fichier = vals("log_fichier")
        LOG_Table = vals("log_table")
    ElseIf IsNumeric(SYS_GetIni("LOG", "LOG_FICHIER", p_nomini)) Then
        LOG_Fichier = CInt(SYS_GetIni("LOG", "LOG_FICHIER", p_nomini))
        If IsNumeric(SYS_GetIni("LOG", "LOG_TABLE", p_nomini)) Then
            LOG_Table = CInt(SYS_GetIni("LOG", "LOG_TABLE", p_nomini))
        Else
            LOG_Table = 0
        End If
    Else
        LOG_Fichier = LOG_WARNING '2
        LOG_Table = 0
    End If
    config_lue = True
    'MsgBox "LOG_fichier=" & LOG_Fichier
    
End Sub

' Envoi du message dans la Base de données
Private Sub LOG_BDD(logName As String, logMsg As String, logType As Long)

    Dim logPoste As String, logUser As String, logAppli As String
    Dim sql As String
    
    If LOG_Table < logType Then Exit Sub
    LOG_BDD_ON = False
    logPoste = SYS_GetComputerName()
    logUser = SYS_GetUserName() & "-" & p_NumUtil
    logAppli = App.EXEName
    
    sql = "insert into log (log_poste, log_user, log_appli, log_type, log_name, log_message) values (?,?,?,?,?,?)"
    Call DB_execute(sql, logPoste, logUser, logAppli, logType, logName, logMsg)
    LOG_BDD_ON = True
    
End Sub

' Extraction du message et du numéro d'erreur à partir d'un ErrObject
Public Function LOG_ERRMSG(e As Variant) As String

    If TypeName(e) = "ErrObject" Then
        LOG_ERRMSG = "Erreur " & e.Number & " " & e.Description
    Else
        LOG_ERRMSG = "Erreur: TypeName(e) <> ErrObject"
    End If

End Function

' Envoi du message vers les differents canaux: Fichier / MsgBox / BDD
' logName: nom de la fonction concernée
' logMsg:  message
' logType: ErrObject (Err) ou LOG_ERROR ou LOG_INFO ou rien (=LOG_ERROR)
' Retourne le message formatté
Public Function LOG(logName As String, logMsg As String, Optional logType As Variant) As String
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim logFile As String, logFileOld As String, lType As Long
    Dim fso As Object, F As Object
    Dim sErr As String
    
    sErr = ""
    If IsMissing(logType) Then logType = LOG_ERROR
    If TypeName(logType) = "ErrObject" Then
        sErr = LOG_ERRMSG(logType)
        logType = LOG_ERROR
    End If
    lType = logType
    
On Error GoTo ErrHandler
    If LOG_MsgBoxes >= logType Then
        MsgBox logMsg & vbCrLf & sErr, vbOKOnly + vbCritical, logName
    End If
    logMsg = logMsg & IIf(Len(sErr) > 0, " ", "") & sErr

    If Not log_getPath() Then
        'MsgBox "Not log_getPath"
        LOG = logName & ": " & logMsg
        Exit Function
    End If
    
    logFile = logPath & App.EXEName & ".log"
    logFileOld = logPath & App.EXEName & "-old.log"
    If App.EXEName = "KaliScrute" Then
        logFile = logPath & App.EXEName & "_" & p_numdocinit & ".log"
        logFileOld = logPath & App.EXEName & "_" & p_numdocinit & "-old.log"
    End If
   
    If Not config_lue Or LOG_Fichier >= logType Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FileExists(logFile)) Then
            If LOG_MaxSize = 0 Then LOG_MaxSize = &H400000 ' 4Mo par defaut
            If fso.GetFile(logFile).size > LOG_MaxSize Then
                Call fso.CopyFile(logFile, logFileOld, OverWriteFiles:=True)
                Call fso.DeleteFile(logFile)
            End If
        End If
        Set F = fso.OpenTextFile(logFile, ForAppending, True)
        F.WriteLine Now() & " L" & logType & vbTab & logName & vbTab & logMsg
        F.Close
        Set F = Nothing
        Set fso = Nothing
        If P_Nom_Log = "" Then
            P_Nom_Log = logFile
        End If
    End If
    
    Debug.Print Now() & " L" & logType & vbTab & logName & vbTab & logMsg
    LOG = logName & ": " & logMsg
    If Not LOG_BDD_ON Then Exit Function
    If Not config_lue Then LOG_config
    Call LOG_BDD(logName, logMsg, lType)
    Exit Function
    
ErrHandler:
    If LOG_MsgBoxes > 0 Then
        MsgBox "Erreur dans LOG:" & Err.Number & vbCrLf & Err.Description
        MsgBox logMsg & vbCrLf & sErr, vbOKOnly + vbCritical, logName
    End If
    
End Function

' ?
Public Function LOG_Conversion(logFctName As String, logMsg As String) As String
    Dim liberr As String, iret As Integer
    Dim nomlog As String
    
    nomlog = "Conversion"
    iret = HTTP_Ecrire_Log(nomlog, logFctName, logMsg, liberr)
    If iret <> HTTP_OK Then
        Call LOG_WARN("LOG_Conversion", liberr)
    End If
End Function

' Appel la fonction LOG avec LOG_ERROR en dernier paramètre
Public Function LOG_ERR(logName As String, logMsg As String) As String
    LOG_ERR = LOG(logName, logMsg, LOG_ERROR)
End Function

' Appel la fonction LOG avec LOG_WARNING en dernier paramètre
Public Function LOG_WARN(logName As String, logMsg As String) As String
    LOG_WARN = LOG(logName, logMsg, LOG_WARNING)
End Function

' Appel la fonction LOG avec LOG_INFO en dernier paramètre
Public Function LOG_INF(logName As String, logMsg As String) As String
    LOG_INF = LOG(logName, logMsg, LOG_INFO)
End Function

' Appel la fonction LOG avec LOG_DEBUG en dernier paramètre
Public Function LOG_DBG(logName As String, logMsg As String) As String
    LOG_DBG = LOG(logName, logMsg, LOG_DEBUG)
End Function

' Appel la fonction LOG avec LOG_DEBUG5 en dernier paramètre
Public Function LOG_DBG5(logName As String, logMsg As String) As String
    LOG_DBG5 = LOG(logName, logMsg, LOG_DEBUG5)
End Function

' Retourne la taille du fichier en bytes.
' Ou -1 si le fichier n'existe pas
Private Function GetLogSize(FileSpec As String) As Long

   Dim fso, F
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (fso.FileExists(FileSpec)) Then
        Set F = fso.GetFile(FileSpec)
        GetLogSize = F.size
   Else
        GetLogSize = -1
   End If
End Function

' pour ordonner les modifs de la variable globale logMEM
Public Sub LOG_RAZMEM()

    logMEM = vbNullString
    
End Sub

' pour ordonner les modifs de la variable globale logMEM
Public Sub LOG_ADDMEM(v_msg As String)

    logMEM = logMEM & v_msg
    
End Sub

' pour ordonner les modifs de la variable globale logMEM
Public Function LOG_GETMEM() As String

    LOG_GETMEM = logMEM
    
End Function

Public Function AddASlash(InString As String) As String
    If Mid(InString, Len(InString), 1) <> "\" Then
        AddASlash = InString & "\"
    Else
        AddASlash = InString
    End If
End Function

Public Function write_trace_operation(ByVal v_type As String, _
                                        ByVal v_operation As String, _
                                        ByVal v_sdate As String, _
                                        ByVal v_stime As String, _
                                        ByVal v_util As Long, _
                                        ByVal v_objet As String, _
                                        ByVal v_commentaire As String, _
                                        ByVal v_detail As String, _
                                        ByVal v_description As String, _
                                        ByVal v_fichier_trace As String, _
                                        ByVal v_fichier As String)
    Dim Destination As String, nb As Long
    Dim sql As String, To_Num As Long, rs As rdoResultset
    Dim ret As Integer
    
    ' Ajouter les colonnes qui manquent (to_IP)
    sql = "SELECT count(*) FROM pg_class,pg_attribute WHERE relname='trace_operation' and attname='to_ip'"
    Call Odbc_Count(sql, nb)
    If nb = 0 Then
        sql = "alter table trace_operation add column to_ip varchar(400)"
        On Error GoTo err_execute
        Odbc_Cnx.Execute (sql)
    End If
err_execute:
 
    If Odbc_AddNew("Trace_Operation", _
                    "TO_Num", _
                    "to_Seq", _
                    True, _
                    To_Num, _
                    "to_type", v_type, "to_commentaire", v_commentaire, "to_operation", v_operation, "to_date", v_sdate, _
                    "to_heure", v_stime, "to_unum", v_util, "to_succes", True, "to_objet", v_objet, _
                    "to_detail", v_detail, "to_ip", p_IP_Locale & " " & p_HostName, "to_description", v_description) = P_ERREUR Then
    End If
    ' copier le fichier
    If v_fichier <> "" Then
        Destination = Replace(p_CheminDoc, "documentation", "")
        Destination = Destination & "kalitmp/doc_import"
        If Not KF_EstRepertoire(Destination, False) Then
            KF_CreerRepertoire (Destination)
        End If
        ret = KF_PutFichier(AddASlash(Destination) & v_fichier_trace, v_fichier)
        If ret <> P_OK Then
            sql = "update trace_operation set to_succes=false, to_detail = '" & v_detail & " non sauvegardé' where to_num=" & To_Num
            Call Odbc_Select(sql, rs)
        End If
    End If
    p_FichierOrigine = ""
    p_ActionOrigine = ""
    p_ActionResume = ""
    p_IdentOrigine = ""
    p_FichierImport = ""
    
    write_trace_operation = To_Num
End Function
Public Function Write_Log(Msg, Optional nomlog As String = "App.EXEName")
    
    Dim fs
    Dim s As String
    Dim Logging As Integer
    Dim s_NomLog As String
    Dim nb As Integer, num As Integer
    Dim sep As String
    
    ' si nomlog As String = "App.EXEName"       =>  le log est en direct
    ' sinon on prend p_Nom_Log et on remplace la fin par nomlog
    
    If LOG_MsgBoxes Then
        MsgBox Msg
    End If
    If nomlog = "App.EXEName" Then
        ' on prend tel quel
        s_NomLog = P_Nom_Log
    Else
        sep = AddASlash("TEST")
        sep = Replace(sep, "TEST", "")
        nb = STR_GetNbchamp(P_Nom_Log, sep)
        s = STR_GetChamp(P_Nom_Log, sep, nb - 1)
        s_NomLog = P_Nom_Log
        s_NomLog = Replace(P_Nom_Log, s, "")
        s_NomLog = s_NomLog & nomlog & ".log"
    End If
    Logging = FreeFile
    Open s_NomLog For Append As #Logging
    Print #Logging, format(Date, "Long Date") & " " & format(Time, "Long Time") & " " & Msg
    Close #Logging
    Exit Function
    If P_Nom_Log = "" Then
        s = AddASlash(App.Path) & "logs"
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Not fs.FolderExists(s) Then
            MkDir (s)
        End If
        If nomlog = "App.EXEName" Then
            s = AddASlash(s) & App.EXEName & ".log"
            P_Nom_Log = s
        Else
            s = AddASlash(s) & nomlog & ".log"
        End If
        'If Not fs.FileExists(s) Then
            Logging = FreeFile
            Open s For Append As #Logging
            Print #Logging, format(Date, "Long Date") & " " & format(Time, "Long Time") & " " & Msg
            Close #Logging
        'End If
    Else
        Logging = FreeFile
        Open P_Nom_Log For Append As #Logging
        Print #Logging, format(Date, "Long Date") & " " & format(Time, "Long Time") & " " & Msg
        Close #Logging
    End If
End Function


