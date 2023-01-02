Attribute VB_Name = "HttpIO"
Option Explicit

Private g_HTTP_VoirProgression As Boolean
Private g_HTTP_FormMaj As Form

Public p_AdrServeur As String
Public p_HTTP_CheminDepot As String
Public p_HTTP_PrgConvPDF As String
Public p_HTTP_TailleFichier As Long
Public p_HTTP_DebutTransaction As Boolean
Public p_HTTP_TimeDébut As Date

' Tableau pour mémoriser les chargements HTTP
' Servira à rePoser les fichiers concernés après utilisation (si non modale)
Public Type HTTP_Fichiers_Chargés
    HTTP_Fullname_Serveur As String
    HTTP_Fullname_Local As String
    HTTP_Type_DosDoc As String
    HTTP_Chargé As Boolean
    HTTP_Locké As Boolean
End Type
Public p_tbl_HTTP_Fichiers_Chargés() As HTTP_Fichiers_Chargés
Public p_bool_HTTP_Fichiers_Chargés As Boolean

Public Type HTTP_Fichiers_Multiples
    HTTP_Numero As Integer
    HTTP_FileName As String
    HTTP_Chargé As Boolean
End Type
Public p_tbl_HTTP_Fichiers_Multiples() As HTTP_Fichiers_Multiples
Public p_bool_HTTP_Fichiers_Multiples As Boolean

' Variables de code d'erreur des différentes fonctions
Public Const HTTP_OK = 0
' Charger un dossier Serveur -> Client (récursif)
Public Const HTTP_GETDOS_ERREUR = -1
Public Const HTTP_GETDOS_LOCKE = -2
Public Const HTTP_GETDOS_VIDE = -3

' Creer un dossier sur le Serveur
Public Const HTTP_CREERDOS_EXISTE_DEJA = 2

' Charger un dossier Client -> Serveur (récursif)
Public Const HTTP_PUTDOS_ERREUR = -4
Public Const HTTP_PUTDOS_DEJA = -5

' Calcul de la taille d'un fichier sur le Serveur
Public Const HTTP_TAILLE_ERREUR = -6

' Charger un fichier à partir du Serveur
Public Const HTTP_GET_ERREUR = -7
Public Const HTTP_GET_LOCKE = -8
Public Const HTTP_GET_OK_VIDE = -9
Public Const HTTP_GET_FIC_INTROUVABLE = -10
Public Const HTTP_GET_DEJA_EN_LOCAL = -11
Public Const HTTP_GET_PAS_COMPLET = -12

' Charger un fichier du Client vers le Serveur
Public Const HTTP_PUT_ERREUR = -13
Public Const HTTP_PUT_PAS_COMPLET = -14

' Supprimer un fichier sur le Serveur
Public Const HTTP_DEL_ERREUR = -15

' Locker un fichier sur le Serveur -> ATTENTION si chgt changer aussi PHP
Public Const HTTP_LOCK_ERREUR = -16
Public Const HTTP_LOCK_AUTRE_USER = -17
Public Const HTTP_LOCK_PASFAIT = -18

' Convertir un fichier en PDF
Public Const HTTP_CONVERT_ERREUR = -19

' Vérifier l'existance d'un fichier sur le Serveur
Public Const HTTP_EXIST_ERREUR = -20

' Copier un fichier du serveur vers le Serveur
Public Const HTTP_COPY_ERREUR = -21

' Reconstituer un fichier (pour PutFile) losrqu'il est de grande taille (> à p_HTTP_MaxParFichier)
Public p_HTTP_MaxParFichier As Double
Public p_HTTP_MaxParPaquet As Double
Public Const HTTP_RECONST_ERREUR = -22

' Liste_Fichiers
Public Const HTTP_LISTEFICH_DOSINEX = -23
Public Const HTTP_LISTEFICH_DOSINACC = -24
Public Const HTTP_LISTEFICH_ERREUR = -25
            
' Renommer un fichier sur le serveur
Public Const HTTP_RENAME_ERREUR = -26

' Vérifier que c'est un répertoire
Public Const HTTP_ESTREP_ERREUR = -27

' Créer un répertoire
Public Const HTTP_CREERREP_ERREUR = -28

' Effacer un répertoire
Public Const HTTP_EFFREP_ERREUR = -29

Public Enum INTERNET_DEF
    INTERNET_DEFAULT_HTTP_PORT2 = 80
    INTERNET_DEFAULT_HTTPS_PORT = 443
End Enum

Private Type S_HTTP_REQUEST
    lInternetSession As Long
    lInternetConnect As Long
    lHttpRequest As Long
End Type

Public Function http_copyfile(ByVal v_sURL As String, _
                              ByVal v_CheminHTTP As String, _
                              ByVal v_CheminFichSrc As String, _
                              ByVal v_CheminFichDest As String, _
                              ByVal v_Session As String, _
                              ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean, bret As Boolean
    Dim stLoad As String, stPost As String, sret As String, sBuf As String
    Dim buf As String, ligne As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    ' Initialise Connect
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP _
            & "&v_CheminFichierSrc=" & v_CheminFichSrc _
            & "&v_CheminFichierDest=" & v_CheminFichDest _
            & "&v_Session=" & v_Session _
            & "&v_NumUtil=" & p_NumUtil
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        http_copyfile = HTTP_COPY_ERREUR
        Exit Function
    End If
    
    ' Send request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "CopyFile : HttpSendRequest=0 : Apache arrêté ?"
        http_copyfile = HTTP_COPY_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "CopyFile : Erreur HttpQueryInfo "
        http_copyfile = HTTP_COPY_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "CopyFile : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        http_copyfile = HTTP_COPY_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        Call http_CloseConnect(http_req)
        http_copyfile = HTTP_COPY_ERREUR
        Exit Function
    End If
    
    Call http_CloseConnect(http_req)
    
    http_copyfile = HTTP_OK

End Function

' Supprimer un fichier sur le serveur
Public Function http_deletefile(ByVal v_sURL As String, _
                                ByVal v_CheminHTTP As String, _
                                ByVal v_CheminFichier As String, _
                                ByVal v_Session As String, _
                                ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean, bret As Boolean
    Dim stLoad As String, stPost As String, sret As String, sBuf As String
    Dim buf As String, ligne As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    ' Initilise Connect
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP & "&v_CheminFichier=" & v_CheminFichier & "&v_Session=" & v_Session & "&v_NumUtil=" & p_NumUtil
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        http_deletefile = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFile : HttpSendRequest=0 : Apache arrêté ?"
        http_deletefile = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFile : Erreur HttpQueryInfo "
        http_deletefile = HTTP_DEL_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFile : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        http_deletefile = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        Call http_CloseConnect(http_req)
        http_deletefile = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    Call http_CloseConnect(http_req)
    
    http_deletefile = HTTP_OK

End Function

' Suppression du serveur
Public Function http_deletefile_simple(ByVal v_sURL As String, _
                                       ByVal v_CheminHTTP As String, _
                                       ByVal v_CheminFichier As String, _
                                       ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean
    Dim stLoad As String, stPost As String, sret As String
    Dim buf As String, ligne As String, sBuf As String
    Dim ret As Integer, fpIn As Integer
    Dim maxn As Long, hFileLocal As Long, lgTotal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long
    Dim RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    ' Initialise Connect
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP & "&v_CheminFichier=" & v_CheminFichier & "&v_NumUtil=" & p_NumUtil
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        http_deletefile_simple = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFileSimple : HttpSendRequest=0 : Apache arrêté ?"
        http_deletefile_simple = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFileSimple : Erreur HttpQueryInfo"
        http_deletefile_simple = HTTP_DEL_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "DeleteFileSimple : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        http_deletefile_simple = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        If InStr(sBuf, "mod_") > 0 Then
            r_liberr = Mid(STR_GetChamp(sBuf, "|", 2), InStr(STR_GetChamp(sBuf, "|", 2), "mod_"))
            r_liberr = Replace(r_liberr, "mod_", "")
        Else
            r_liberr = STR_GetChamp(sBuf, "|", 2)
        End If
        Call http_CloseConnect(http_req)
        http_deletefile_simple = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    Call http_CloseConnect(http_req)
    
    http_deletefile_simple = HTTP_OK

End Function

Public Function HTTP_Appel_copyfile(ByVal v_FichServSrc As String, _
                                    ByVal v_FichServDest As String, _
                                    ByRef r_liberr As String) As Integer
    
    Dim strHTTP As String, Session As String
    Dim iret As Integer
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/copy_file.php"
    
    iret = http_copyfile(strHTTP, p_HTTP_CheminDepot, v_FichServSrc, v_FichServDest, Session, r_liberr)
        
    HTTP_Appel_copyfile = iret
    
End Function

Public Function HTTP_Appel_creer_repertoire(ByVal v_nomrep_srv As String, _
                                            ByRef r_liberr As String) As Integer
                                          
    Dim bresult As Boolean
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim stStatusCode As String, stStatusText As String, sBuf As String
    Dim strHTTP As String, ligne As String, stLoad As String
    Dim stPost As String, sret As String, buf As String
    Dim nomfich_Serveur As String
    Dim ret As Integer, iret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Initialise Connect
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/creer_repertoire.php"
    ret = http_InitConnect(strHTTP, http_req)
    If CBool(ret) = False Then
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_NomRep=" & v_nomrep_srv & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Creer_Repertoire : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "Creer_Repertoire : Erreur HttpQueryInfo "
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Creer_Repertoire : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) = "OK" Then
        r_liberr = STR_GetChamp(sBuf, "|", 1)
        r_liberr = STR_GetChamp(r_liberr, " ", 0)
        HTTP_Appel_creer_repertoire = HTTP_OK
    ElseIf left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
    Else
        r_liberr = sBuf
        HTTP_Appel_creer_repertoire = HTTP_CREERREP_ERREUR
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_Appel_deletefile(ByVal v_FichServeur As String, _
                                      ByVal v_bMessage As Boolean, _
                                      ByRef r_liberr As String) As Integer
    
    Dim strHTTP As String, Session As String
    Dim iret As Integer
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/delete_file.php"
    Session = HTTP_RandomAlphaNumString(5)
    
    iret = http_deletefile(strHTTP, p_HTTP_CheminDepot, v_FichServeur, Session, r_liberr)
        
    HTTP_Appel_deletefile = iret
    
End Function

Public Function HTTP_Appel_deletefile_simple(ByVal v_FichServeur As String, _
                                             ByVal v_bMessage As Boolean, _
                                             ByRef r_liberr As String) As Integer
                                             
    Dim strHTTP As String
    Dim iret As Integer
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/delete_file_simple.php"
    
    iret = http_deletefile_simple(strHTTP, p_HTTP_CheminDepot, v_FichServeur, r_liberr)
        
    HTTP_Appel_deletefile_simple = iret
    
End Function

Public Function HTTP_Appel_ListeFichiers(ByVal v_FichServeur As String, _
                                         ByRef r_liberr As String) As Integer
    
    Dim strHTTP As String, strChemin As String, FichServeur_Extension As String
    Dim FichServeur_fichier As String, FichServeur_chemin As String
    Dim iret As Integer
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String

    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_ListeFichiers Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_ListeFichiers = HTTP_LISTEFICH_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_ListeFichiers = HTTP_LISTEFICH_ERREUR
        Exit Function
    End If
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/liste_fichiers.php"
    
    iret = http_listefichiers(strHTTP, p_HTTP_CheminDepot, FichServeur_chemin, FichServeur_fichier, FichServeur_Extension, r_liberr)
        
    HTTP_Appel_ListeFichiers = iret
    
End Function

Public Function HTTP_Appel_LockDelock_file(ByVal v_locker As Boolean, _
                                            ByVal v_FichServeur As String, _
                                            ByVal v_bMessage As Boolean, _
                                            ByRef r_liberr As String) As Integer
    
    ' Locker ou Délocker un fichier sur le serveur
    
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim strHTTP As String
    Dim iret As Integer
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_LockDelock_file Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_LockDelock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "HTTP_Appel_LockDelock_file Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_LockDelock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If
    
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/locker_file.php"
    Session = HTTP_RandomAlphaNumString(5)
    
    iret = HTTP_lock_delock_file(v_locker, strHTTP, p_HTTP_CheminDepot, FichServeur_chemin, FichServeur_fichier, FichServeur_Extension, True, Session, r_liberr)
    
    HTTP_Appel_LockDelock_file = iret
    
End Function

Public Function HTTP_Appel_GetFile(ByVal v_FichServeur As String, _
                                   ByVal v_FichLocal As String, _
                                   ByVal v_bLocker As Boolean, _
                                   ByVal v_bMessage As Boolean, _
                                   ByRef r_liberr As String) As Integer
                                   
    Dim FichServeur_chemin As String, FichServeur_fichier As String
    Dim FichServeur_Extension As String, strHTTP As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim iret As Integer
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_GetFile Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_GetFile = HTTP_GET_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_GetFile = HTTP_GET_ERREUR
        Exit Function
    End If
    
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/get_file.php"
    Session = HTTP_RandomAlphaNumString(5)
    
    iret = http_getfile(Session, strHTTP, p_HTTP_CheminDepot, _
                    FichServeur_chemin, FichServeur_fichier, _
                    FichServeur_Extension, v_FichLocal, _
                    v_bLocker, v_bMessage, r_liberr)
        
'    If iret = HTTP_GET_LOCKE Then
'    ElseIf iret = HTTP_GET_ERREUR Then
'    ElseIf iret = HTTP_GET_PAS_COMPLET Then
'    ElseIf iret = HTTP_GET_FIC_INTROUVABLE Then
'    ElseIf iret = HTTP_GET_DEJA_EN_LOCAL Then
'    ElseIf iret = HTTP_OK Then
'    End If
    
    HTTP_Appel_GetFile = iret
    
End Function

Public Function HTTP_Appel_effacer_repertoire(ByVal v_nomrep_srv As String, _
                                              ByRef r_liberr As String) As Integer
                                          
    Dim bresult As Boolean
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, sBuf As String
    Dim strHTTP As String, ligne As String, stLoad As String
    Dim stPost As String, sret As String, buf As String
    Dim nomfich_Serveur As String
    Dim ret As Integer, iret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Initialise Connect
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/effacer_repertoire.php"
    ret = http_InitConnect(strHTTP, http_req)
    If CBool(ret) = False Then
        HTTP_Appel_effacer_repertoire = HTTP_EFFREP_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_NomRep=" & v_nomrep_srv & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Effacer_Repertoire : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_Appel_effacer_repertoire = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "Effacer_Repertoire : Erreur HttpQueryInfo "
        HTTP_Appel_effacer_repertoire = HTTP_EFFREP_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Effacer_Repertoire : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_Appel_effacer_repertoire = HTTP_EFFREP_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) = "OK" Then
        r_liberr = STR_GetChamp(sBuf, "|", 1)
        r_liberr = STR_GetChamp(r_liberr, " ", 0)
        HTTP_Appel_effacer_repertoire = HTTP_OK
    ElseIf left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        HTTP_Appel_effacer_repertoire = HTTP_EFFREP_ERREUR
    Else
        r_liberr = sBuf
        HTTP_Appel_effacer_repertoire = HTTP_EFFREP_ERREUR
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_Appel_est_repertoire(ByVal v_nomrep_srv As String, _
                                          ByVal v_bMessage As Boolean, _
                                          ByRef r_liberr As String) As Integer
                                          
    Dim bresult As Boolean
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, sBuf As String
    Dim stStatusCode As String, stStatusText As String
    Dim strHTTP As String, ligne As String, stLoad As String
    Dim stPost As String, sret As String, buf As String
    Dim nomfich_Serveur As String
    Dim ret As Integer, iret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Initialise Connect
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/est_repertoire.php"
    ret = http_InitConnect(strHTTP, http_req)
    If CBool(ret) = False Then
        HTTP_Appel_est_repertoire = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_NomRep=" & v_nomrep_srv & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Est_Repertoire : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_Appel_est_repertoire = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
        
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "Est_Repertoire : Erreur HttpQueryInfo"
        HTTP_Appel_est_repertoire = HTTP_ESTREP_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Est_Repertoire : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_Appel_est_repertoire = HTTP_ESTREP_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) = "OK" Then
        r_liberr = STR_GetChamp(sBuf, "|", 1)
        r_liberr = STR_GetChamp(r_liberr, " ", 0)
        HTTP_Appel_est_repertoire = HTTP_OK
    ElseIf left(sBuf, 6) = "ERREUR" Then
        HTTP_Appel_est_repertoire = HTTP_ESTREP_ERREUR
        r_liberr = STR_GetChamp(sBuf, "|", 2)
    Else
        r_liberr = sBuf
        HTTP_Appel_est_repertoire = HTTP_ESTREP_ERREUR
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_Appel_fichier_existe(ByVal v_FichServeur As String, _
                                          ByVal v_bMessage As Boolean, _
                                          ByRef r_liberr As String) As Integer
                                          
    Dim bresult As Boolean
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, sBuf As String
    Dim stStatusCode As String, stStatusText As String
    Dim strHTTP As String, ligne As String, stLoad As String
    Dim stPost As String, sret As String, buf As String
    Dim nomfich_Serveur As String
    Dim ret As Integer, iret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    Dim http_req As S_HTTP_REQUEST
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_fichier_existe Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "HTTP_Appel_fichier_existe Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
        
    r_liberr = ""
    
    ' Initialise Connect
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/fichier_existe.php"
    ret = http_InitConnect(strHTTP, http_req)
    If CBool(ret) = False Then
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_CheminFichier=" & FichServeur_chemin & "&" _
            & "v_NomFichier=" & FichServeur_fichier & "&" _
            & "v_ExtensionFichier=" & FichServeur_Extension & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : Erreur HttpQueryInfo"
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) = "OK" Then
        r_liberr = STR_GetChamp(sBuf, "|", 1)
        r_liberr = STR_GetChamp(r_liberr, " ", 0)
        HTTP_Appel_fichier_existe = HTTP_OK
    ElseIf left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
    Else
        r_liberr = sBuf
        HTTP_Appel_fichier_existe = HTTP_EXIST_ERREUR
    End If
    
    Call http_CloseConnect(http_req)

End Function

Public Function HTTP_Appel_EffacerFichier(ByVal v_FichServeur As String, _
                                          ByVal v_bMessage As Boolean, _
                                          ByRef r_liberr As String) As Integer
                                          
    Dim bresult As Boolean
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, sBuf As String
    Dim stStatusCode As String, stStatusText As String
    Dim strHTTP As String, ligne As String, stLoad As String
    Dim stPost As String, sret As String, buf As String
    Dim nomfich_Serveur As String
    Dim ret As Integer, iret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    Dim http_req As S_HTTP_REQUEST
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_EffacerFichier Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_EffacerFichier = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_EffacerFichier = HTTP_DEL_ERREUR
        Exit Function
    End If
    
    r_liberr = ""
    
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    
    ' Initialise Connect
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/effacer_fichiers.php"
    ret = http_InitConnect(strHTTP, http_req)
    If CBool(ret) = False Then
        HTTP_Appel_EffacerFichier = HTTP_DEL_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_CheminDossierServeur=" & FichServeur_chemin & "&" _
            & "v_NomFichierServeur=" & FichServeur_fichier & "&" _
            & "v_ExtFichierServeur=" & FichServeur_Extension & "&" _
            & "v_ForcerTout=1&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_Appel_EffacerFichier = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : Erreur HttpQueryInfo "
        HTTP_Appel_EffacerFichier = HTTP_EXIST_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "Fichier_Existe : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_Appel_EffacerFichier = HTTP_EXIST_ERREUR
        Exit Function
    End If
    
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) = "OK" Then
        r_liberr = STR_GetChamp(sBuf, "|", 1)
        r_liberr = STR_GetChamp(r_liberr, " ", 0)
        HTTP_Appel_EffacerFichier = HTTP_OK
    ElseIf left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        HTTP_Appel_EffacerFichier = HTTP_EXIST_ERREUR
    Else
        r_liberr = sBuf
        HTTP_Appel_EffacerFichier = HTTP_EXIST_ERREUR
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function http_getfile(ByVal v_Session As String, _
                             ByVal v_sURL As String, _
                             ByVal v_chemin As String, _
                             ByVal v_CheminFichier_Serveur As String, _
                             ByVal v_NomFichier_Serveur As String, _
                             ByVal v_ExtensionFichier_Serveur As String, _
                             ByVal v_nomfich_Copie As String, _
                             ByVal v_locker As Boolean, _
                             ByVal v_bool_message As Boolean, _
                             ByRef r_liberr As String) As Integer

    Dim bresult As Boolean, bPrem As Boolean
    Dim stStatusCode As String, stStatusText As String, stPost As String
    Dim stLoad As String, liberr As String, sret As String, buf As String
    Dim strTailleChargement  As String, ligne As String, nomfich_Serveur As String
    Dim nomFicRenomme As String, CheminTmp As String, Locker As String
    Dim ret As Integer, fpIn As Integer, iRetTaille As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long, RetClose As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, ResteSeconde As Long
    Dim TailleChargement As Long, NbSeconde As Long, taille As Long
    Dim TimeDebut As Date, TimePrem As Date
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Récupérer la taille du fichier
    p_HTTP_TailleFichier = 0
    iRetTaille = HTTP_gettaille(v_sURL, v_CheminFichier_Serveur, v_NomFichier_Serveur, _
                                v_ExtensionFichier_Serveur, taille, liberr)
    If iRetTaille <> HTTP_OK Then
        r_liberr = liberr
        http_getfile = iRetTaille
        Exit Function
    End If
        
    p_HTTP_TailleFichier = taille
    maxn = p_HTTP_MaxParPaquet
    TailleChargement = taille + maxn
        
    If g_HTTP_VoirProgression Then
        If iRetTaille = HTTP_OK Then
            strTailleChargement = Int((TailleChargement / 1024))
            If val(strTailleChargement) > 1024 Then
                strTailleChargement = Round(strTailleChargement / 1024, 2) & " M Octets)  "
            Else
                strTailleChargement = Round(strTailleChargement, 2) & " K Octets)  "
            End If
            g_HTTP_FormMaj.lblMaj.Caption = "Chargement à partir du serveur de " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " (" & strTailleChargement
            g_HTTP_FormMaj.PgbarHTTPDTaille.max = TailleChargement
            g_HTTP_FormMaj.PgbarHTTPDTaille.Value = 0
            g_HTTP_FormMaj.lblHTTPDTemps.Caption = "Temps restant"
            g_HTTP_FormMaj.lblHTTPDTaille.Caption = "Volume chargé"
            DoEvents
        End If
    End If
    
    TimeDebut = DateTime.Now()
    
    ' Initialise Connect
    If v_locker Then
        Locker = "O"
    Else
        Locker = "N"
    End If
    v_sURL = v_sURL & "?v_Locker=" & Locker _
            & "&v_Session=" & v_Session
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then GoTo ErrorHandle

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_chemin & "&" _
            & "v_CheminFichier=" & v_CheminFichier_Serveur & "&" _
            & "v_NomFichier=" & v_NomFichier_Serveur & "&" _
            & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        http_getfile = HTTP_GET_ERREUR
        r_liberr = "GetFile : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    ' Status Code
    buf = String(p_HTTP_MaxParPaquet, Chr(0))
    n = p_HTTP_MaxParPaquet
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        http_getfile = HTTP_GET_ERREUR
        r_liberr = "GetFile : Erreur HttpQueryInfo QUERY_STATUS_CODE "
        GoTo ErrorHandle
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        http_getfile = HTTP_GET_ERREUR
        r_liberr = "GetFile : Erreur HttpQueryInfo QUERY_STATUS_CODE sret=" & left(Trim(buf), 3) & " " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
    
    nomfich_Serveur = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur
    
    v_nomfich_Copie = v_nomfich_Copie & "_Session_" & v_Session
    
    ' Création du fichier tampon
    hFileLocal = CreateFile(v_nomfich_Copie, GENERIC_WRITE Or GENERIC_READ, _
                            0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFileLocal < 0 Then
        http_getfile = HTTP_GET_ERREUR
        r_liberr = "GetFile : CreateFile " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
        
    ' Lecture
    lgTotal = 0
    maxn = p_HTTP_MaxParPaquet
    Do While hFileLocal > 0
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
        
        If iRetTaille = HTTP_OK Then
            If p_HTTP_TailleFichier = 0 Then
                'MsgBox v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " est vide"
                Call http_CloseConnect(http_req)
                Call FICH_EffacerFichier(v_nomfich_Copie, False)
                r_liberr = "GetFile : Le Fichier " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " est vide"
                http_getfile = HTTP_GET_ERREUR
                Exit Function
            End If
            If g_HTTP_VoirProgression Then
                If g_HTTP_FormMaj.PgbarHTTPDTaille.max < lgTotal Then
                    g_HTTP_FormMaj.PgbarHTTPDTaille.max = lgTotal
                End If
                g_HTTP_FormMaj.PgbarHTTPDTaille.Value = lgTotal
            End If
            
            TimePrem = DateTime.Now()
            If n > 0 Then
                If g_HTTP_VoirProgression Then
                    NbSeconde = DateDiff("s", TimeDebut, TimePrem)
                    If NbSeconde = 0 Then NbSeconde = 1
                    g_HTTP_FormMaj.PgbarHTTPDTemps.max = NbSeconde / n * p_HTTP_TailleFichier
                    ResteSeconde = NbSeconde / n * (p_HTTP_TailleFichier - lgTotal)
                    If ResteSeconde < 0 Then
                        ResteSeconde = 0
                        g_HTTP_FormMaj.PgbarHTTPDTemps.Value = ResteSeconde
                        g_HTTP_FormMaj.lblHTTPDTemps.Caption = "Terminé"
                    Else
                        g_HTTP_FormMaj.PgbarHTTPDTemps.Value = ResteSeconde
                    End If
                    DoEvents
                End If
            End If
        End If
                        
        If n > 0 Then
            ret = WriteFile(hFileLocal, ByVal buf, n, nb_ecrits, ByVal 0&)
            If nb_ecrits < 1 Then
                ret = GetLastError
            End If
        Else
            RetClose = CloseHandle(hFileLocal)
            If RetClose <> 1 Then
                http_getfile = HTTP_GET_ERREUR
                r_liberr = "GetFile : Impossible de fermer " & v_nomfich_Copie
                Call FICH_EffacerFichier(v_nomfich_Copie, False)
                GoTo ErrorHandle
            End If
            hFileLocal = 0
        End If
    Loop
    If lgTotal = 0 Then
        http_getfile = HTTP_GET_OK_VIDE
        r_liberr = "GetFile : Fichier Vide " & v_nomfich_Copie
    End If
    
    Call http_CloseConnect(http_req)
    
    If http_getfile = HTTP_GET_OK_VIDE Then
        GoTo SuiteGet
    End If
    
    http_getfile = HTTP_OK
    ' Voir si erreur
    If FICH_FichierExiste(v_nomfich_Copie) Then
    
        fpIn = FreeFile
        FICH_OuvrirFichier v_nomfich_Copie, FICH_LECTURE, fpIn
            
        If Not EOF(fpIn) Then
            Line Input #fpIn, ligne
            If left(ligne, 6) = "ERREUR" Then
                If STR_GetChamp(ligne, "|", 1) = 5 Then
                    http_getfile = HTTP_GET_LOCKE
                    r_liberr = STR_GetChamp(ligne, "|", 2)
                ElseIf InStr(LCase(ligne), nomfich_Serveur & " introuvable") > 0 Then
                    http_getfile = HTTP_GET_FIC_INTROUVABLE
                    r_liberr = "GetFile : " & nomfich_Serveur & " introuvable"
                    If v_bool_message Then
                        MsgBox "Erreur " & r_liberr
                    End If
                Else
                    http_getfile = HTTP_GET_ERREUR
                    r_liberr = "GetFile : " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie de " & nomfich_Serveur & " vers " & v_nomfich_Copie
                    If v_bool_message Then
                        MsgBox "Erreur " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie de " & nomfich_Serveur & " vers " & v_nomfich_Copie
                    End If
                End If
            'ElseIf InStr(LCase(ligne), "warning") > 0 Or InStr(LCase(ligne), "parse") > 0 Or InStr(ligne, "404") Then
            '    http_getfile = HTTP_GET_ERREUR
            '    r_liberr = "GetFile : " & ligne
            '    If v_bool_message Then
            '        MsgBox "Erreur " & r_liberr
            '    End If
            Else
                http_getfile = HTTP_OK
            End If
        Else
            http_getfile = HTTP_OK
        End If
        
        If http_getfile = HTTP_OK Then
            If iRetTaille = HTTP_OK Then
                ' comparer la taille des deux fichiers
                If InStr(v_nomfich_Copie, "TraceHTTP.txt_Session_") > 0 Then
                    If g_HTTP_VoirProgression Then
                        g_HTTP_FormMaj.lblMaj.Caption = "Chargement terminé avec succès"
                    End If
                ElseIf p_HTTP_TailleFichier <> FileLen(v_nomfich_Copie) Then
                    http_getfile = HTTP_GET_PAS_COMPLET
                    r_liberr = "Taille du fichier sur le serveur " & p_HTTP_TailleFichier & Chr(13) & Chr(10) & "Taille du fichier chargé " & FileLen(v_nomfich_Copie)
                    If g_HTTP_VoirProgression Then
                        g_HTTP_FormMaj.lblMaj.Caption = r_liberr
                    End If
                Else
                    If g_HTTP_VoirProgression Then
                        g_HTTP_FormMaj.lblMaj.Caption = "Chargement terminé avec succès"
                    End If
                End If
            End If
        End If

        Close (fpIn)
    Else
        http_getfile = HTTP_GET_ERREUR
        r_liberr = "GetFile : Fichier non récupéré : " & v_nomfich_Copie
    End If
    If http_getfile = HTTP_OK Then
SuiteGet:
        ' renommer pour enlever la session
        ' si existe déjà => renommer avec date et heure
        
        If InStr(v_nomfich_Copie, "TraceHTTP.txt_Session_") > 0 Then
            nomFicRenomme = Replace(v_nomfich_Copie, "_Session_" & v_Session, "")
            Call FICH_RenommerFichier(v_nomfich_Copie, nomFicRenomme)
        ElseIf FICH_FichierExiste(Replace(v_nomfich_Copie, "_Session_" & v_Session, "")) Then
            nomFicRenomme = Replace(v_nomfich_Copie, "_Session_" & v_Session, "")
            nomFicRenomme = Replace(nomFicRenomme, "." & v_ExtensionFichier_Serveur, "_Date_" & Format(Date, "yyyy_mm_dd") & "_Heure_" & Format(Time, "hh_nn_ss") & "." & v_ExtensionFichier_Serveur)
            r_liberr = "Le fichier " & Replace(v_nomfich_Copie, "_Session_" & v_Session, "") & " a été renommé en " & nomFicRenomme
            Call FICH_RenommerFichier(Replace(v_nomfich_Copie, "_Session_" & v_Session, ""), nomFicRenomme)
            http_getfile = HTTP_GET_DEJA_EN_LOCAL
            Call FICH_RenommerFichier(v_nomfich_Copie, Replace(v_nomfich_Copie, "_Session_" & v_Session, ""))
        Else
            Call FICH_RenommerFichier(v_nomfich_Copie, Replace(v_nomfich_Copie, "_Session_" & v_Session, ""))
        End If
    ElseIf http_getfile = HTTP_GET_LOCKE Or http_getfile = HTTP_GET_ERREUR Or http_getfile = HTTP_GET_FIC_INTROUVABLE Then
        ' supprimer le fichier généré
        Call FICH_EffacerFichier(v_nomfich_Copie, False)
    End If
    ' supprimer les fichiers temporaires sur le serveur
    If http_getfile = HTTP_OK Then
        'If InStr(v_nomfich_Copie, "TraceHTTP.txt_Session_") > 0 Then
        '    CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_Session_" & v_Session
        'Else
        '    ' CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_Session_" & v_Session
        '    CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_Session_" & v_Session
        'End If
        
        ' essayer avec le doc contenant le numutil
        'CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_Session_" & v_Session
        'If HTTP_Appel_deletefile_simple(CheminTmp, False, r_liberr) = HTTP_DEL_ERREUR Then
        '    ' essayer avec le numutil
        '    CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_" & p_NumUtil & "_Session_" & v_Session
        '    If HTTP_Appel_deletefile_simple(CheminTmp, False, r_liberr) = HTTP_DEL_ERREUR Then
        '        ' essayer en enlever le .mod
        '        CheminTmp = Replace(CheminTmp, ".mod_", ".")
        '        If HTTP_Appel_deletefile_simple(CheminTmp, False, r_liberr) = HTTP_DEL_ERREUR Then
        '            ' essayer en ajoutant .mod
        '            CheminTmp = Replace(CheminTmp, v_NomFichier_Serveur & ".", v_NomFichier_Serveur & ".mod_")
        '            If HTTP_Appel_deletefile_simple(CheminTmp, False, r_liberr) = HTTP_DEL_ERREUR Then
        '                MsgBox "ne s'efface pas"
        '            End If
        '        End If
        '    End If
        'End If
    End If
    
ErrorHandle:
    
    'CheminTmp = v_chemin & v_NomFichier_Serveur & ".mod_" & v_ExtensionFichier_Serveur & "_" & p_NumUtil & "_Session_" & v_Session
    
    CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_Session_" & v_Session
    If HTTP_Appel_deletefile_simple(CheminTmp, False, liberr) = HTTP_DEL_ERREUR Then
        ' essayer avec le numutil
        CheminTmp = v_chemin & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & "_" & p_NumUtil & "_Session_" & v_Session
        If HTTP_Appel_deletefile_simple(CheminTmp, False, liberr) = HTTP_DEL_ERREUR Then
            ' essayer en enlever le .mod
            CheminTmp = Replace(CheminTmp, ".mod_", ".")
            If HTTP_Appel_deletefile_simple(CheminTmp, False, liberr) = HTTP_DEL_ERREUR Then
                ' essayer en ajoutant .mod
                CheminTmp = Replace(CheminTmp, v_NomFichier_Serveur & ".", v_NomFichier_Serveur & ".mod_")
                If HTTP_Appel_deletefile_simple(CheminTmp, False, liberr) = HTTP_DEL_ERREUR Then
                    'MsgBox "ne s'efface pas"
                End If
            End If
        End If
    End If
    
    'Call HTTP_Appel_deletefile_simple(CheminTmp, False, liberr)
    
    Err.Clear
    
    Call http_CloseConnect(http_req)
    
    If g_HTTP_VoirProgression Then
        g_HTTP_FormMaj.lblMaj.Caption = ""
    End If
    
End Function

Public Function http_renamefile(ByVal v_sURL As String, _
                                ByVal v_CheminHTTP As String, _
                                ByVal v_CheminFichSrc As String, _
                                ByVal v_CheminFichDest As String, _
                                ByVal v_Session As String, _
                                ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean, bret As Boolean
    Dim stLoad As String, stPost As String, sret As String
    Dim buf As String, ligne As String, sBuf As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    ' Initialise Connect
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP _
            & "&v_CheminFichierSrc=" & v_CheminFichSrc _
            & "&v_CheminFichierDest=" & v_CheminFichDest _
            & "&v_Session=" & v_Session _
            & "&v_NumUtil=" & p_NumUtil
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        http_renamefile = HTTP_RENAME_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "RenameFile : HttpSendRequest=0 : Apache arrêté ?"
        http_renamefile = HTTP_RENAME_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "RenameFile : Erreur HttpQueryInfo"
        http_renamefile = HTTP_RENAME_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "RenameFile : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        http_renamefile = HTTP_RENAME_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        http_renamefile = HTTP_RENAME_ERREUR
        Exit Function
    End If
    
    Call http_CloseConnect(http_req)
    
    http_renamefile = HTTP_OK

End Function

Public Sub HTTP_SetVoirProgression(ByVal v_voirprog As Boolean, _
                                   ByVal v_form As Form)
                          
    g_HTTP_VoirProgression = v_voirprog
    On Error Resume Next
    Set g_HTTP_FormMaj = v_form
    g_HTTP_FormMaj.FrmMaj.visible = v_voirprog
    g_HTTP_FormMaj.frmHTTPD.visible = v_voirprog
    g_HTTP_FormMaj.PgbarHTTPDTaille.Value = 0
    g_HTTP_FormMaj.PgbarHTTPDTemps.Value = 0
    DoEvents
    
End Sub

Public Function http_listefichiers(ByVal v_sURL As String, _
                                    ByVal v_CheminHTTP As String, _
                                    ByVal v_FichServ_chemin As String, _
                                    ByVal v_FichServ_fichier As String, _
                                    ByVal v_FichServ_extension As String, _
                                    ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean, bret As Boolean
    Dim stLoad As String, stPost As String, sret As String, coderr As String
    Dim buf As String, ligne As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim strRetour As String
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim PremPaquet As Boolean
    Dim http_req As S_HTTP_REQUEST
    
    ' Initialise Connect
    v_sURL = v_sURL
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP _
            & "&v_CheminDossierServeur=" & v_FichServ_chemin _
            & "&v_NomFichierServeur=" & v_FichServ_fichier _
            & "&v_ExtFichierServeur=" & v_FichServ_extension _
            & "&v_NumUtil=" & p_NumUtil
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        http_listefichiers = HTTP_LISTEFICH_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    stPost = "v_CheminHTTP=" & v_CheminHTTP _
            & "&v_CheminDossierServeur=" & v_FichServ_chemin _
            & "&v_NomFichierServeur=" & v_FichServ_fichier _
            & "&v_ExtFichierServeur=" & v_FichServ_extension _
            & "&v_NumUtil=" & p_NumUtil
    stPost = "" ' tout est passé en GET pour l'instant
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "ListeFichiers : HttpSendRequest=0 : Apache arrêté ?"
        http_listefichiers = HTTP_LISTEFICH_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "ListeFichiers : Erreur HttpQueryInfo "
        http_listefichiers = HTTP_LISTEFICH_ERREUR
        GoTo ErrorHandle
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        http_listefichiers = HTTP_LISTEFICH_ERREUR
        r_liberr = "ListeFichiers : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        GoTo ErrorHandle
    End If
    
    lgTotal = 0
    
    hFileLocal = 1
    strRetour = ""
    PremPaquet = True
    http_listefichiers = HTTP_OK
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
        If PremPaquet Then
            ' voir si OK ou pas
            PremPaquet = False
            If left(buf, 2) = "OK" Or strRetour <> "" Then
                strRetour = strRetour & buf
            ElseIf left(buf, 6) = "ERREUR" Then
                coderr = STR_GetChamp(buf, "|", 1)
                If coderr = "1" Then
                    http_listefichiers = HTTP_LISTEFICH_DOSINEX
                ElseIf coderr = "2" Then
                    r_liberr = STR_GetChamp(buf, "|", 2)
                    r_liberr = Replace(r_liberr, Chr(0), "")
                    http_listefichiers = HTTP_LISTEFICH_ERREUR
                Else
                    http_listefichiers = HTTP_LISTEFICH_DOSINACC
                End If
                GoTo ErrorHandle
            Else
                r_liberr = buf
                http_listefichiers = HTTP_LISTEFICH_ERREUR
                GoTo ErrorHandle
            End If
        Else
            strRetour = strRetour & buf
        End If
        If (n = 0) Then
            hFileLocal = 0
        End If
    Loop
    
    ' Pour mettre le retour en forme correcte
    If InStr(strRetour, "!FIN!;") > 0 Then
        strRetour = left(strRetour, InStr(strRetour, "!FIN!;") + Len("!FIN!;") - 1)
    End If
        
    Call http_CloseConnect(http_req)
    
    http_listefichiers = HTTP_OK
    r_liberr = strRetour
    
    Exit Function

ErrorHandle:
    Err.Clear
    Call http_CloseConnect(http_req)

End Function

Public Function HTTP_lock_delock_file(ByVal v_locker As Boolean, _
                                    ByVal v_sURL As String, _
                                    ByVal v_chemin As String, _
                                    ByVal v_CheminFichier_Serveur As String, _
                                    ByVal v_NomFichier_Serveur As String, _
                                    ByVal v_ExtensionFichier_Serveur As String, _
                                    ByVal v_bool_message As Boolean, _
                                    ByVal v_Session As String, _
                                    ByRef r_liberr As String) As Integer
    
    Dim bresult As Boolean
    Dim stStatusCode As String, stStatusText As String, ligne As String
    Dim stLoad As String, stPost As String, sBuf As String
    Dim sret As String, buf As String, smode_lock As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long, n As Long
    Dim hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Initialise Connect
    v_sURL = v_sURL & "?v_Locker=" & IIf(v_locker, "Lock", "DeLock")
    v_sURL = v_sURL & "&v_Session=" & v_Session
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_lock_delock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_chemin & "&" _
            & "v_CheminFichier=" & v_CheminFichier_Serveur & "&" _
            & "v_NomFichier=" & v_NomFichier_Serveur & "&" _
            & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "LockerFile : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_lock_delock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "LockerFile : Erreur HttpQueryInfo "
        HTTP_lock_delock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "LockerFile : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_lock_delock_file = HTTP_LOCK_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        HTTP_lock_delock_file = STR_GetChamp(sBuf, "|", 1)
    Else
        HTTP_lock_delock_file = HTTP_OK
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_gettaille(ByVal v_sURL As String, _
                               ByVal v_CheminFichier_Serveur As String, _
                               ByVal v_NomFichier_Serveur As String, _
                               ByVal v_ExtensionFichier_Serveur As String, _
                               ByRef r_taille As Long, _
                               ByRef r_liberr As String) As Long

    Dim bresult As Boolean
    Dim stStatusCode As String, stStatusText As String, stLoad As String
    Dim stPost As String, sret As String, buf As String, ligne As String
    Dim nomfich_Serveur As String, Locker As String, s As String, sBuf As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    v_sURL = Replace(v_sURL, "get_file", "get_taille")
    v_sURL = Replace(v_sURL, "put_file", "get_taille")

    ' Initialise Connect
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
            & "v_Chemin=" & v_CheminFichier_Serveur & "&" _
            & "v_Fichier=" & v_NomFichier_Serveur & "&" _
            & "v_Extension=" & v_ExtensionFichier_Serveur & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetTaille : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        Exit Function
    End If
    
    ' Status Code
    buf = String(p_HTTP_MaxParPaquet, Chr(0))
    n = p_HTTP_MaxParPaquet
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult Then
        sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
        If val(left(Trim(buf), 1)) > 3 Then
            bresult = False
        End If
    End If
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetTaille : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) <> "OK" Then
        If left(sBuf, 6) = "ERREUR" Then
            r_liberr = STR_GetChamp(buf, "|", 1)
        Else
            r_liberr = sBuf
        End If
        Call http_CloseConnect(http_req)
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        Exit Function
    End If
    
    s = STR_GetChamp(sBuf, "|", 1)
    r_taille = STR_GetChamp(s, " ", 0)
    
    Call http_CloseConnect(http_req)
    
    HTTP_gettaille = HTTP_OK
    
End Function

Private Function http_InitConnect(ByVal v_strURL As String, _
                                 ByRef r_httpreq As S_HTTP_REQUEST) As Boolean
    
    Dim iPort As Integer
    Dim strObject As String
    Dim intPos As Integer
    
    If left$(LCase(v_strURL), 7) = "http://" Then
        v_strURL = Right$(v_strURL, Len(v_strURL) - 7)
    Else
        If left$(LCase(v_strURL), 8) = "https://" Then
            v_strURL = Right$(v_strURL, Len(v_strURL) - 8)
        End If
    End If
    
    intPos = InStr(1, v_strURL, "/")
    If intPos > 0 Then
        strObject = Right$(v_strURL, Len(v_strURL) - intPos + 1)
        v_strURL = left$(v_strURL, intPos - 1)
    End If
    
    intPos = InStr(1, v_strURL, ":")
    If intPos > 0 Then
        iPort = val(Right$(v_strURL, Len(v_strURL) - intPos))
        v_strURL = left$(v_strURL, intPos - 1)
    Else
        iPort = INTERNET_DEFAULT_HTTP_PORT
    End If
    
    r_httpreq.lInternetSession = InternetOpen("KaliDoc", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)

    r_httpreq.lInternetConnect = InternetConnect(r_httpreq.lInternetSession, v_strURL, iPort, _
                         vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    
    r_httpreq.lHttpRequest = HttpOpenRequest(r_httpreq.lInternetConnect, "POST", strObject, _
                     "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
                                    
    http_InitConnect = CBool(r_httpreq.lHttpRequest)
    
End Function

' generates a random alphanumeirc string of a given length
Public Function HTTP_RandomAlphaNumString(ByVal intLen As Integer)
    
    Dim StrReturn As String
    
    Dim x As Integer
    Dim c As Byte
    
    Randomize
    
    For x = 1 To intLen
        c = Int(Rnd() * 127)
    
        If (c >= Asc("0") And c <= Asc("9")) Or _
           (c >= Asc("A") And c <= Asc("Z")) Or _
           (c >= Asc("a") And c <= Asc("z")) Then
           
            StrReturn = StrReturn & Chr(c)
        Else
            x = x - 1
        End If
    Next x
    
    HTTP_RandomAlphaNumString = StrReturn
    
End Function

Public Function HTTP_Appel_PutFile(ByVal v_FichServeur As String, _
                                   ByVal v_FichLocal As String, _
                                   ByVal v_bMessage As Boolean, _
                                   ByVal v_bDeLocker As Boolean, _
                                   ByRef r_liberr As String) As Integer
    
    Dim FichServeur_chemin As String, FichServeur_fichier As String
    Dim FichLocal As String, strChemin As String, sBuf As String
    Dim FichTmp As String, nomfich As String, nomFichTmp As String
    Dim ch As String, strHTTP As String, FichServeur_Extension As String
    Dim liberr As String
    Dim iret As Integer, i As Integer, NumFichier As Integer, NbFichier As Integer
    Dim fpOut As Integer, fpIn As Integer
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    Dim taille As Long, taille_dec As Long, lpFileSizeHight As Long
    Dim TaillePaquet As Double, ResteàTRSF As Double, TailleDéjà As Double
    Dim sTime As String
    Dim b_attendre As Boolean
    Dim leItbl As Integer
    Dim nbSecondeAttente As Integer
    Dim DateDébut As Date
    
    If g_HTTP_VoirProgression Then
        g_HTTP_FormMaj.lblMaj.Caption = "Mise à jour sur le serveur"
        DoEvents
    End If
    
    NbFichier = 1
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_PutFile Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_PutFile = HTTP_PUT_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "HTTP_Appel_PutFile Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_PutFile = HTTP_PUT_ERREUR
        Exit Function
    End If
    FichServeur_Extension = Mid$(strChemin, pos + 1)
    FichServeur_fichier = left$(strChemin, pos - 1)
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    
    If Not FICH_FichierExiste(v_FichLocal) Then
        r_liberr = "HTTP_Appel_PutFile : Le fichier local " & v_FichLocal & " n'existe pas"
        HTTP_Appel_PutFile = HTTP_PUT_ERREUR
        Exit Function
    End If
    taille = GetFileSize(v_FichLocal)
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/put_file.php"

    TailleDéjà = 0
    If val(taille) > p_HTTP_MaxParFichier Then
        NbFichier = Int(val(taille) / p_HTTP_MaxParFichier) + 1
        ResteàTRSF = taille
        fpIn = FreeFile
        'MsgBox "v_FichServeur=" & v_FichServeur
        'MsgBox "V_FichLocal=" & V_FichLocal
        Open v_FichLocal For Binary As #fpIn
        NumFichier = 0
        p_HTTP_TimeDébut = DateTime.Now()
        sTime = Format(Time, "hh_mm_ss")
        
        p_bool_HTTP_Fichiers_Multiples = False
        While (Not EOF(fpIn)) And ResteàTRSF > 0
            If ResteàTRSF > p_HTTP_MaxParFichier Then
                TaillePaquet = p_HTTP_MaxParFichier
            Else
                TaillePaquet = ResteàTRSF
            End If
            If ResteàTRSF > 0 Then
                ch = Space(TaillePaquet)
                Get #fpIn, , ch
                NumFichier = NumFichier + 1
                If p_bool_HTTP_Fichiers_Multiples Then
                    ReDim Preserve p_tbl_HTTP_Fichiers_Multiples(UBound(p_tbl_HTTP_Fichiers_Multiples) + 1)
                Else
                    ReDim Preserve p_tbl_HTTP_Fichiers_Multiples(0)
                End If
                p_tbl_HTTP_Fichiers_Multiples(UBound(p_tbl_HTTP_Fichiers_Multiples)).HTTP_Numero = NumFichier
                p_tbl_HTTP_Fichiers_Multiples(UBound(p_tbl_HTTP_Fichiers_Multiples)).HTTP_FileName = v_FichLocal
                p_tbl_HTTP_Fichiers_Multiples(UBound(p_tbl_HTTP_Fichiers_Multiples)).HTTP_Chargé = False
                p_bool_HTTP_Fichiers_Multiples = True
                leItbl = UBound(p_tbl_HTTP_Fichiers_Multiples)
                ' créer ce fichier
                nomfich = v_FichLocal & "_" & NumFichier
                fpOut = FreeFile
                Open nomfich For Binary As #fpOut
                Put #fpOut, , ch
                Close #fpOut
                taille_dec = GetFileSize(nomfich)
                iret = http_putfile(str(taille_dec), strHTTP, FichServeur_chemin, _
                            FichServeur_fichier, _
                            FichServeur_Extension & "_" & NumFichier, _
                            nomfich, v_FichServeur, v_bDeLocker, NbFichier, _
                            NumFichier, TailleDéjà, r_liberr)
                'MsgBox NumFichier & " iRet=" & iRet
                If iret = HTTP_OK Then
                    TailleDéjà = TailleDéjà + TaillePaquet
                    p_tbl_HTTP_Fichiers_Multiples(leItbl).HTTP_Chargé = True
                Else
                    MsgBox "Erreur PutFile sur fichier n° " & NumFichier & " :" & r_liberr
                    iret = HTTP_PUT_ERREUR
                    GoTo LabSuiteErr
                End If
            End If
            ResteàTRSF = ResteàTRSF - TaillePaquet
        Wend
        Close #fpIn
        
        strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/reconstituer_file.php"
        strHTTP = Replace(strHTTP, "\", "/")
        
        If g_HTTP_VoirProgression Then
            g_HTTP_FormMaj.lblMaj.Caption = "Reconstruction de " & FichServeur_fichier & "." & FichServeur_Extension & " " & IIf(NumFichier > 1, "[" & NumFichier & " fichiers]", "")
            DoEvents
        End If
        iret = HTTP_ReconstituerFile(str(taille), p_HTTP_CheminDepot, _
                    strHTTP, FichServeur_chemin, _
                    FichServeur_fichier, FichServeur_Extension, "", _
                    NbFichier, r_liberr)
        If iret <> HTTP_OK Then
            iret = HTTP_PUT_ERREUR
            If g_HTTP_VoirProgression Then
                g_HTTP_FormMaj.lblHTTPD.Caption = "Reconstruction non réussie" & Chr(13) & Chr(10) & r_liberr
                DoEvents
            End If
        Else
            If g_HTTP_VoirProgression Then
                g_HTTP_FormMaj.lblMaj.Caption = "Reconstruction réussie"
                DoEvents
            End If
        End If
        
LabSuiteErr:
        'MsgBox iret & " Effacer les fichier temporaires locaux"
        For i = 1 To NbFichier
            If FICH_FichierExiste(v_FichLocal & "_" & i) Then
                Call FICH_EffacerFichier(v_FichLocal & "_" & i, True)
            End If
        Next i
        ' Effacer le .mod  22.mod_doc_73
        If iret = HTTP_OK Then
            Call HTTP_Appel_deletefile_simple(FichServeur_chemin & "/" & FichServeur_fichier & ".mod_" & FichServeur_Extension & "_" & p_NumUtil, False, liberr)
        End If
    ' Pas de découpage
    Else
        p_HTTP_TimeDébut = DateTime.Now()
        iret = http_putfile(str(taille), strHTTP, _
                            FichServeur_chemin, FichServeur_fichier, _
                            FichServeur_Extension, v_FichLocal, _
                            v_FichServeur, v_bDeLocker, 1, 1, 0, r_liberr)
    End If
    
    If g_HTTP_VoirProgression Then
        g_HTTP_FormMaj.lblMaj.Caption = ""
        DoEvents
    End If

    HTTP_Appel_PutFile = iret
    
End Function

Public Function http_putfile(ByVal v_Taille As String, _
                             ByVal v_sURL As String, _
                             ByVal v_CheminFichier_Serveur As String, _
                             ByVal v_NomFichier_Serveur As String, _
                             ByVal v_ExtensionFichier_Serveur As String, _
                             ByVal v_nomFichTmp As String, _
                             ByVal v_nomFichDest As String, _
                             ByRef v_DeLocker As Boolean, _
                             ByVal v_NbFichier As Integer, _
                             ByVal v_NumFichier As Integer, _
                             ByVal v_TailleDéjà As Double, _
                             ByRef r_liberr As String) As Integer

    Dim bresult As Boolean, bPrem As Boolean
    Dim stLoad As String, stPost1 As String, stPost2 As String
    Dim strBoundary As String, MimeType As String, sret As String
    Dim buf As String, ligne As String, sURL As String
    Dim strTailleChargement As String, sBuf As String
    Dim ret As Integer, fpIn As Integer
    Dim lBufferLength   As Long, maxn As Long, n As Long
    Dim nb_transmis As Long, hindex As Long, nb_total As Long
    Dim RetClose As Long, hFileLocal As Long, lgTot As Long
    Dim lgTotal As Long, ResteSeconde As Long, TailleChargement As Long
    Dim NbSeconde As Long
    Dim TimeDébut As Date, TimePrem As Date
    Dim Fl As Double
    Dim BufferIn As INTERNET_BUFFERS
    Dim http_req As S_HTTP_REQUEST
    
    p_HTTP_TailleFichier = val(Trim(v_Taille))
    maxn = p_HTTP_MaxParPaquet
    TailleChargement = p_HTTP_TailleFichier + maxn
    
    If g_HTTP_VoirProgression Then
        g_HTTP_FormMaj.lblMaj = "Mise à jour sur le serveur : " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur
        g_HTTP_FormMaj.frmHTTPD.visible = True
        g_HTTP_FormMaj.frmHTTPD.ZOrder 0
        strTailleChargement = Int((TailleChargement / 1024))
        If val(strTailleChargement) > 1024 Then
            strTailleChargement = Round(strTailleChargement / 1024, 2) & " M Octets)  "
        Else
            strTailleChargement = Round(strTailleChargement, 2) & " K Octets)  "
        End If
        g_HTTP_FormMaj.lblMaj.Caption = "Dépot vers le serveur de " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " (" & strTailleChargement & IIf(v_NbFichier > 1, "[" & v_NumFichier & "/" & v_NbFichier & " fichiers]", "")
        g_HTTP_FormMaj.PgbarHTTPDTaille.max = TailleChargement 'v_TailleDéjà + TailleChargement
        g_HTTP_FormMaj.PgbarHTTPDTaille.Value = 0 'v_TailleDéjà
        g_HTTP_FormMaj.lblHTTPDTemps.Caption = "Temps restant"
        g_HTTP_FormMaj.lblHTTPDTaille.Caption = "Volume chargé"
        DoEvents
    End If
    
    'TimeDébut = DateTime.Now()
    TimeDébut = p_HTTP_TimeDébut
    
    MimeType = "application/octet-stream"
    strBoundary = HTTP_RandomAlphaNumString(32)
    
    v_nomFichDest = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur

    stPost1 = "--" & strBoundary & vbCrLf & _
             "Content-Disposition: form-data; " & _
             "name=""" & v_nomFichTmp & """; " & _
             "filename=""" & v_nomFichDest & """" & vbCrLf & _
             "Content-Type: " & MimeType & vbCrLf & vbCrLf
    
    stPost2 = vbCrLf & "--" & strBoundary & "--"
    ' find the length of the request body - this is required for the Content-Length header
    lgTot = Len(stPost1) + FileLen(v_nomFichTmp) + Len(stPost2)
    
    ' headers
    stLoad = "Content-Type: multipart/form-data, boundary=" & strBoundary & vbCrLf & _
             "Content-Length: " & lgTot & vbCrLf & vbCrLf
    
    On Error GoTo lab_fin
    sURL = v_sURL _
        & "?filename=" & v_nomFichDest & "&" _
        & "v_CheminFichier=" & v_CheminFichier_Serveur & "&" _
        & "v_NomFichier=" & v_NomFichier_Serveur & "&" _
        & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&" _
        & "v_CheminHTTP=" & p_HTTP_CheminDepot & "&" _
        & "v_NumUtil=" & p_NumUtil & "&" _
        & "v_NbFichier=" & v_NbFichier & "&" _
        & "v_NumFichier=" & v_NumFichier & "&" _
        & "v_Taille=" & p_HTTP_TailleFichier & "&"
    If v_DeLocker Then
        sURL = sURL & "v_DeLocker=O"
    Else
        sURL = sURL & "v_DeLocker=N"
    End If
    'MsgBox sUrl
InitialConnect:
    ' Initialise Connect
    bresult = http_InitConnect(sURL, http_req)
    If bresult = False Then
        r_liberr = "PutFile : Erreur InitialHttpConnect sret=" & left(Trim(buf), 3) & " " & sURL
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    ' Ajoute Header à la requête
    ret = HttpAddRequestHeaders(http_req.lHttpRequest, stLoad, Len(stLoad), HTTP_ADDREQ_FLAG_ADD)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur HttpAddRequestHeaders"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    ' Envoi début de requête "extended"
    BufferIn.dwStructSize = 40
    BufferIn.Next = 0
    BufferIn.lpcszHeader = 0
    BufferIn.dwHeadersLength = 0
    BufferIn.dwHeadersTotal = 0
    BufferIn.lpvBuffer = 0
    BufferIn.dwBufferLength = 0
    BufferIn.dwBufferTotal = lgTot
    BufferIn.dwOffsetLow = 0
    BufferIn.dwOffsetHigh = 0
    ret = HttpSendRequestEx(http_req.lHttpRequest, BufferIn, 0, 0, 0)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur HttpSendRequestEx"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    ' Paquet début de requête
    nb_transmis = 0
    nb_total = 0
    ret = InternetWriteFile(http_req.lHttpRequest, stPost1, Len(stPost1), nb_transmis)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur InternetWriteFile"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Ouvre le fichier à transférer
    hFileLocal = CreateFile(v_nomFichTmp, GENERIC_READ, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        0&, OPEN_EXISTING, 0, 0)
    
    lgTotal = 0 ' v_TailleDéjà
    Do While hFileLocal > 0
        buf = String(maxn, Chr(0))
        ' Lecture locale
        ret = ReadFile(hFileLocal, ByVal buf, maxn, n, ByVal 0&)
        If n > 0 Then
            ' Ecriture Serveur
            ret = InternetWriteFile(http_req.lHttpRequest, buf, n, nb_transmis)
            If CBool(ret) Then
                lgTotal = lgTotal + n
                If g_HTTP_VoirProgression Then
                     If g_HTTP_FormMaj.PgbarHTTPDTaille.max < lgTotal Then
                         g_HTTP_FormMaj.PgbarHTTPDTaille.max = lgTotal
                     End If
                     g_HTTP_FormMaj.PgbarHTTPDTaille.Value = lgTotal
                     TimePrem = DateTime.Now()
                     If n > 0 Then
                         NbSeconde = DateDiff("s", TimeDébut, TimePrem)
                         If NbSeconde = 0 Then NbSeconde = 1
                         g_HTTP_FormMaj.PgbarHTTPDTemps.max = NbSeconde / n * p_HTTP_TailleFichier
                         ResteSeconde = NbSeconde / n * (p_HTTP_TailleFichier - lgTotal)
                         If ResteSeconde < 0 Then
                             ResteSeconde = 0
                             g_HTTP_FormMaj.PgbarHTTPDTemps.Value = ResteSeconde
                             g_HTTP_FormMaj.lblHTTPDTemps.Caption = "Terminé"
                         Else
                             g_HTTP_FormMaj.PgbarHTTPDTemps.Value = ResteSeconde
                         End If
                     End If
                     DoEvents
                End If
            Else
                r_liberr = "PutFile : Erreur InternetWriteFile"
                http_putfile = HTTP_PUT_ERREUR
                GoTo lab_fin
            End If
            nb_total = nb_total + n
        Else
            RetClose = CloseHandle(hFileLocal)
            If RetClose <> 1 Then
                Call FICH_EffacerFichier(v_nomFichTmp, False)
                r_liberr = "PutFile : Impossible de lire dans " & v_nomFichTmp
                http_putfile = HTTP_PUT_ERREUR
                GoTo lab_fin
            End If
            hFileLocal = -1
        End If
    Loop
    
    ' Paquet fin de fichier
    ret = InternetWriteFile(http_req.lHttpRequest, stPost2, Len(stPost2), nb_transmis)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur InternetWriteFile"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    ' Fin de requête
    ret = HttpEndRequest(http_req.lHttpRequest, 0, 0, 0)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur HttpEndRequest"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        r_liberr = "PutFile : Erreur HttpQueryInfo"
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        r_liberr = "PutFile : Erreur Trim(Buf) " & buf
        http_putfile = HTTP_PUT_ERREUR
        GoTo lab_fin
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 2) <> "OK" Then
        If left(sBuf, 6) = "ERREUR" Then
            r_liberr = STR_GetChamp(buf, "|", 1)
        Else
            r_liberr = sBuf
        End If
        Call http_CloseConnect(http_req)
        http_putfile = HTTP_PUT_ERREUR
        Exit Function
    End If
    
lab_fin:
    If g_HTTP_VoirProgression Then
        g_HTTP_FormMaj.lblMaj.Caption = ""
        DoEvents
    End If
    
    Err.Clear
    Call http_CloseConnect(http_req)
    
    http_putfile = HTTP_OK
    
End Function

Public Function HTTP_Appel_Convert_pdf(ByVal v_FichServeur As String, _
                                        ByVal v_FichLocal As String, _
                                        ByVal v_bMessage As Boolean, _
                                        ByRef r_liberr As String) As Integer
    
    Dim FichServeur_chemin As String, FichServeur_fichier As String, FichServeur_Extension As String
    Dim FichServeurGS_Chemin As String, FichServeurGS_Fichier As String, FichServeurGS_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim FichTmp As String, liberr As String, strHTTP As String
    Dim nomfich As String, nomFichTmp As String, FichServeurGS_Complet As String
    Dim RandomFichier As String
    Dim iret As Integer, i As Integer, fpOut As Integer, fpIn As Integer
    Dim taille As Long, lpFileSizeHight As Long
    Dim pos As Long, pos1 As Long, pos2 As Long, slash As String
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/convert_pdf.php"
    Session = HTTP_RandomAlphaNumString(5)
    
    ' FichServeurGS est le fichier à donner à ghostscript
    ' FichServeurPS est le fichier ps de destination
    
    ' 1) Faire un put_file vers FichGS
    
    ' décomposer FichServeur
    pos1 = InStrRev(v_FichServeur, "/")
    If pos1 > 0 Then
        slash = "/"
        pos = pos1
    End If
    
    If pos1 + pos2 = 0 Then
        r_liberr = "HTTP_Appel_Convert_pdf Pb avec " & v_FichServeur & " : pas de /"
        HTTP_Appel_Convert_pdf = HTTP_CONVERT_ERREUR
        Exit Function
    End If
    
    strChemin = Mid$(v_FichServeur, pos + 1)
    pos = InStrRev(strChemin, ".")
    If pos = 0 Then
        r_liberr = "Pb avec " & v_FichServeur & " : pas d'extension"
        HTTP_Appel_Convert_pdf = HTTP_CONVERT_ERREUR
        Exit Function
    End If
    FichServeurGS_Extension = Mid$(strChemin, pos + 1)
    FichServeurGS_Fichier = left$(strChemin, pos - 1)
    
    RandomFichier = HTTP_RandomAlphaNumString(3)
    FichServeurGS_Fichier = FichServeurGS_Fichier & "_" & RandomFichier
    FichServeurGS_Chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    FichServeurGS_Complet = FichServeurGS_Chemin & "/" & FichServeurGS_Fichier & "." & FichServeurGS_Extension
    FichServeur_chemin = left$(v_FichServeur, Len(v_FichServeur) - Len(strChemin) - 1)
    
    p_HTTP_TimeDébut = DateTime.Now()
    iret = HTTP_Appel_PutFile(FichServeurGS_Complet, v_FichLocal, False, False, liberr)
    If iret = HTTP_OK Then
        ' 2) Appeler convert_file pour convertir en pdf et déplacer au bon endroit
        strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/convert_pdf.php"
        Session = ""
        iret = HTTP_convert_pdf(strHTTP, p_HTTP_CheminDepot, FichServeurGS_Chemin, FichServeurGS_Fichier, FichServeurGS_Extension, RandomFichier, False, Session, r_liberr)
        If iret = HTTP_CONVERT_ERREUR Then
            HTTP_Appel_Convert_pdf = P_ERREUR
            Exit Function
        End If
    Else
        r_liberr = liberr
    End If
    
    HTTP_Appel_Convert_pdf = iret
    
End Function

Public Function HTTP_convert_pdf(ByVal v_sURL As String, _
                                 ByVal v_CheminHTTP As String, _
                                 ByVal v_CheminFichier_Serveur As String, _
                                 ByVal v_NomFichier_Serveur As String, _
                                 ByVal v_ExtensionFichier_Serveur As String, _
                                 ByVal v_RandomFichier_Serveur As String, _
                                 ByVal v_bool_message As Boolean, _
                                 ByVal v_Session As String, _
                                 ByRef r_liberr As String) As Integer
    
    Dim stStatusCode As String, stStatusText As String
    Dim lgTotal As Long
    Dim stLoad As String, sBuf As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim bresult As Boolean
    Dim Locker As String
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    'Initialise Connect
    v_sURL = v_sURL & "?v_Session=" & v_Session
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_convert_pdf = HTTP_CONVERT_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_CheminHTTP & "&" _
            & "v_CheminFichier=" & v_CheminFichier_Serveur & "&" _
            & "v_NomFichier=" & v_NomFichier_Serveur & "&" _
            & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&" _
            & "v_PrgConvPDF=" & p_HTTP_PrgConvPDF & "&" _
            & "v_RandomFichier=" & v_RandomFichier_Serveur & "&" _
            & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "ConvertPDF : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_convert_pdf = HTTP_CONVERT_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "ConvertPDF : Erreur HttpQueryInfo "
        HTTP_convert_pdf = HTTP_CONVERT_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        HTTP_convert_pdf = HTTP_CONVERT_ERREUR
        r_liberr = "ConvertPDF : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, Chr(0), "")
        HTTP_convert_pdf = HTTP_CONVERT_ERREUR
    Else
        HTTP_convert_pdf = HTTP_OK
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_ReconstituerFile(ByVal v_Taille As Double, _
                                      ByVal v_CheminHTTP As String, _
                                      ByVal v_sURL As String, _
                                      ByVal v_CheminFichier_Serveur As String, _
                                      ByVal v_NomFichier_Serveur As String, _
                                      ByVal v_ExtensionFichier_Serveur As String, _
                                      ByVal v_Session As String, _
                                      ByVal v_NbFichier As Integer, _
                                      ByRef r_liberr As String) As Integer
    ' Reconstituer un fichier sur le serveur qui a été découpé en plusieurs fichiers
    Dim bresult As Boolean
    Dim stLoad As String, stPost As String, sret As String
    Dim buf As String, ligne As String, sBuf As String
    Dim ret As Integer, fpIn As Integer
    Dim lgTotal As Long, maxn As Long, hFileLocal As Long
    Dim n As Long, hindex As Long, nb_ecrits As Long, RetClose As Long
    Dim http_req As S_HTTP_REQUEST
    
    ' Initialise Connect
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP _
            & "&v_CheminFichier=" & v_CheminFichier_Serveur _
            & "&v_NomFichier=" & v_NomFichier_Serveur _
            & "&v_ExtensionFichier=" & v_ExtensionFichier_Serveur _
            & "&v_NbFichier=" & v_NbFichier _
            & "&v_NumUtil=" & p_NumUtil _
            & "&v_Taille=" & v_Taille
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_ReconstituerFile = HTTP_RECONST_ERREUR
        Exit Function
    End If
    
    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "HTTP_ReconstituerFile : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_ReconstituerFile = HTTP_RECONST_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "HTTP_ReconstituerFile : Erreur HttpQueryInfo "
        HTTP_ReconstituerFile = HTTP_RECONST_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "HTTP_ReconstituerFile : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_ReconstituerFile = HTTP_RECONST_ERREUR
        Exit Function
    End If
    
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        HTTP_ReconstituerFile = HTTP_RECONST_ERREUR
    Else
        HTTP_ReconstituerFile = HTTP_OK
    End If
    
    Call http_CloseConnect(http_req)

End Function

Public Function HTTP_Appel_renamefile(ByVal v_FichServSrc As String, _
                                      ByVal v_FichServDest As String, _
                                      ByRef r_liberr As String) As Integer
    
    Dim strHTTP As String, Session As String
    Dim iret As Integer
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/rename_file.php"
    
    iret = http_renamefile(strHTTP, p_HTTP_CheminDepot, v_FichServSrc, v_FichServDest, Session, r_liberr)
        
    HTTP_Appel_renamefile = iret
    
End Function

Private Function http_CloseConnect(ByRef v_httpreq As S_HTTP_REQUEST)
    
    InternetCloseHandle (v_httpreq.lHttpRequest)
    InternetCloseHandle (v_httpreq.lInternetConnect)
    InternetCloseHandle (v_httpreq.lInternetSession)

End Function

Public Function HTTP_Appel_GetDos(ByVal v_NomDos As String, _
                                  ByVal v_CheminLocal As String, _
                                  ByVal v_CheminServeur As String, _
                                  ByRef r_liberr As String) As Integer
    
    Dim nomdos As String, arbreRep As String, Session As String
    Dim strHTTP As String
    Dim iret As Integer
    
    nomdos = v_CheminLocal & "\" & v_NomDos
    If FICH_EstRepertoire(nomdos, False) Then
        HTTP_RecRmDir (nomdos)
    End If
    If Not FICH_EstRepertoire(nomdos, False) Then
        MkDir (nomdos)
    End If
    
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/get_dos.php"
    Session = HTTP_RandomAlphaNumString(5)
    
    iret = HTTP_getdos(strHTTP, p_HTTP_CheminDepot, v_NomDos, v_CheminLocal, v_CheminServeur, False, False, Session, r_liberr)
    
    HTTP_Appel_GetDos = iret
    
End Function

Public Function HTTP_getdos(ByVal v_sURL As String, _
                            ByVal v_CheminHTTP As String, _
                            ByVal v_NomDossier As String, _
                            ByVal v_CheminDossier_Local As String, _
                            ByVal v_CheminDossier_Serveur As String, _
                            ByVal v_bool_message As Boolean, _
                            ByVal v_locker As Boolean, _
                            ByVal v_Session As String, _
                            ByRef r_liberr As String) As Integer

    Dim stStatusCode As String, stStatusText As String, sBuf As String
    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim iret As Integer
    Dim nomDos_Serveur As String
    Dim bresult As Boolean
    Dim Locker As String
    Dim NomFichierDir As String
    Dim stype As String, snom As String
    Dim http_req As S_HTTP_REQUEST
    
    ' Retourner à l'appelant la liste des fichiers et dossiers d'un dossier
    
    r_liberr = ""
    
    ' Initialise Connect
    If v_Session <> "" Then
        v_sURL = v_sURL & "?v_Session=" & v_Session
    End If
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_getdos = HTTP_GETDOS_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_CheminHTTP & "&"
    stPost = stPost & "v_CheminDossierServeur=" & v_CheminDossier_Serveur & "/" & v_NomDossier & "&"
    stPost = stPost & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetDos : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_getdos = HTTP_GETDOS_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetDos : Erreur HttpQueryInfo " & NomFichierDir
        HTTP_getdos = HTTP_GETDOS_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetDos : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3) & " " & NomFichierDir
        HTTP_getdos = HTTP_GETDOS_ERREUR
        Exit Function
    End If
    
    ' Création fichier local
    NomFichierDir = p_CheminDossierTravailLocal & "\FichierDir_" & v_NomDossier & Rnd() & ".txt" ' & "_Session_" & v_Session
    hFileLocal = CreateFile(NomFichierDir, GENERIC_WRITE Or GENERIC_READ, _
                        0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If (hFileLocal < 0) Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetDos : CreateFile " & NomFichierDir
        HTTP_getdos = HTTP_GETDOS_ERREUR
        Exit Function
    End If
    
    lgTotal = 0
    
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
        If (n > 0) Then
            ret = WriteFile(hFileLocal, ByVal buf, n, nb_ecrits, ByVal 0&)
            If (nb_ecrits < 1) Then
                ret = GetLastError
            End If
        Else
            RetClose = CloseHandle(hFileLocal)
            If RetClose <> 1 Then
                Call http_CloseConnect(http_req)
                HTTP_getdos = HTTP_GETDOS_ERREUR
                r_liberr = "GetDos : Impossible de fermer " & NomFichierDir
                Call FICH_EffacerFichier(NomFichierDir, False)
                Exit Function
            End If
            hFileLocal = 0
        End If
    Loop
    If lgTotal = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "GetDos : Fichier Vide " & NomFichierDir
        HTTP_getdos = HTTP_GETDOS_VIDE
        Exit Function
    End If
    
    Call http_CloseConnect(http_req)
    
    HTTP_getdos = HTTP_OK
    ' Voir si erreur
    If FICH_FichierExiste(NomFichierDir) Then
        fpIn = FreeFile
        FICH_OuvrirFichier NomFichierDir, FICH_LECTURE, fpIn
        If Not EOF(fpIn) Then
            Line Input #fpIn, ligne
            If left(ligne, 6) = "ERREUR" Then
                If STR_GetChamp(ligne, "|", 1) = 5 Then
                    HTTP_getdos = HTTP_GETDOS_LOCKE
                    r_liberr = Mid(STR_GetChamp(ligne, "|", 2), InStr(STR_GetChamp(ligne, "|", 2), "mod_"))
                    r_liberr = Replace(r_liberr, "mod_", "")
                Else
                    HTTP_getdos = HTTP_GETDOS_ERREUR
                    r_liberr = "GetDos : " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie vers " & NomFichierDir
                    If v_bool_message Then
                        MsgBox "Erreur " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie vers " & NomFichierDir
                    End If
                End If
            'ElseIf InStr(LCase(ligne), "warning") > 0 Or InStr(ligne, "404") Then
            '    HTTP_getdos = HTTP_GETDOS_ERREUR
            '    r_liberr = "GetDos : " & ligne
            '    If v_bool_message Then
            '        MsgBox "Erreur " & r_liberr
            '    End If
            Else
                HTTP_getdos = HTTP_OK
            End If
        Else
            HTTP_getdos = HTTP_OK
        End If
        Close (fpIn)
    Else
        HTTP_getdos = HTTP_GETDOS_ERREUR
        r_liberr = "GetDos : FichierDir non récupéré : " & NomFichierDir
    End If
    If HTTP_getdos = HTTP_OK Or HTTP_getdos = HTTP_GET_LOCKE Then
        ' créer les dossiers et les fichiers
        FICH_OuvrirFichier NomFichierDir, FICH_LECTURE, fpIn
        'MsgBox NomFichierDir
        While Not EOF(fpIn)
            Line Input #fpIn, ligne
            'MsgBox ligne
            If ligne <> "" Then
                snom = STR_GetChamp(ligne, "|", 0)
                If snom <> "." And snom <> ".." Then
                    stype = STR_GetChamp(ligne, "|", 1)
                    'MsgBox sNom
                    If stype = "dir" Then
                        ' créer le dossier
                        MkDir (v_CheminDossier_Local & "/" & v_NomDossier & "/" & snom)
                        iret = HTTP_getdos(v_sURL, v_CheminHTTP, snom, v_CheminDossier_Local & "/" & v_NomDossier, v_CheminDossier_Serveur & "/" & v_NomDossier, v_bool_message, v_locker, "", r_liberr)
                        If iret = HTTP_OK Then
                            HTTP_getdos = iret
                        Else
                            HTTP_getdos = iret
                            Exit Function
                        End If
                    ElseIf stype = "file" Then
                        ' charger le fichier
                        iret = HTTP_Appel_GetFile(v_CheminDossier_Serveur & "/" & v_NomDossier & "/" & snom, _
                                    v_CheminDossier_Local & "/" & v_NomDossier & "/" & snom, _
                                    False, False, r_liberr)
                    
                        If iret = HTTP_OK Or HTTP_GET_OK_VIDE Or iret = HTTP_GET_DEJA_EN_LOCAL Then
                            ' C'est OK
                            HTTP_getdos = HTTP_OK
                        Else
                            HTTP_getdos = HTTP_GETDOS_ERREUR
                            Exit Function
                        End If
                    End If
                End If
            End If
        Wend
        Close (fpIn)
        
        Call FICH_EffacerFichier(NomFichierDir, False)
    End If
    
End Function

Public Function HTTP_Appel_PutDos(ByVal v_CheminDosLocal As String, _
                                  ByVal v_CheminDosServeur As String, _
                                  ByVal v_bCreerDossier As Boolean, _
                                  ByVal v_bViderAvant As Boolean, _
                                  ByRef r_liberr As String) As Integer
    ' copie le contenu de v_CheminDosLocal dans v_CheminDosServeur
    '
    ' si besoin, créer d'abord le dossier de destination
    'HTTP_creerDos(strHTTP, p_HTTP_CheminDepot, <dossier serveur de destination>, <nom du dossier>, r_liberr)
    'If iret = HTTP_CREERDOS_EXISTE_DEJA Then
    '    HTTP_PutDos = HTTP_OK
    'ElseIf iret = HTTP_OK Or iret = HTTP_PUTDOS_DEJA Then
    '    HTTP_PutDos = HTTP_OK
    'Else
    '    HTTP_PutDos = HTTP_PUTDOS_ERREUR
    'End If

    HTTP_Appel_PutDos = HTTP_PutDos(v_CheminDosLocal, _
                            v_CheminDosServeur, _
                            v_bCreerDossier, v_bViderAvant, r_liberr)

End Function

Public Function HTTP_PutDos(ByVal v_CheminDosLocal As String, _
                            ByVal v_CheminDosServeur As String, _
                            ByVal v_bCreerDossier As Boolean, _
                            ByVal v_bViderAvant As Boolean, _
                            ByRef r_liberr As String) As Integer
    
    Dim fd As Integer
    Dim bPremier As Boolean
    Dim ret As Integer
    Dim Session As String, liberr As String
    Dim strHTTP As String
    Dim sURL As String
    Dim CheminDossier_Local As String
    Dim CheminDossier_Serveur As String
    Dim FichierTmp As String
    Dim ligne As String
    Dim NomDossierLocal As String
    Dim TypeFic As String, nomfic As String
    Dim nomIn_Chemin As String
    Dim nomIn_Fichier As String
    Dim nomIn_Extension As String
    Dim nomInCpy As String
    Dim iret As Integer
    Dim FichServeur As String, FichLocal As String
    Dim nomrep As String
    Dim fso As FileSystemObject
    Dim Dossier As Variant
    Dim fileItem As Variant
    Dim nomdos_loc As String, nomdos_srv As String
    Dim s As String
        
    strHTTP = "http://" & p_AdrServeur & "/TRSF_HTTP/put_dos.php"
    
    ' chemin local
    If FICH_EstRepertoire(v_CheminDosLocal, False) Then
        ' le créer sur le serveur
        'If v_bCreerDossier Then
        '    nomdos_srv = left$(v_CheminDosServeur, InStrRev(v_CheminDosServeur, "/") - 1)
        '    iret = HTTP_creerDos(strHTTP, p_HTTP_CheminDepot, nomdos_srv, "", r_liberr)
        '    If iret = HTTP_CREERDOS_EXISTE_DEJA Then
        '        HTTP_PutDos = HTTP_OK
        '    ElseIf iret = HTTP_OK Or iret = HTTP_PUTDOS_DEJA Then
        '        HTTP_PutDos = HTTP_OK
        '    Else
        '        HTTP_PutDos = HTTP_PUTDOS_ERREUR
        '    End If
        'End If
        Set fso = CreateObject("Scripting.FileSystemObject")
        ' copier les fichiers
        For Each Dossier In fso.GetFolder(v_CheminDosLocal).Files
            'MsgBox Dossier.Name
                
            FichLocal = v_CheminDosLocal & "/" & Dossier.Name
            FichServeur = v_CheminDosServeur & "/" & Dossier.Name
            iret = HTTP_Appel_PutFile(FichServeur, FichLocal, False, False, liberr)
                
            If iret = HTTP_PUT_ERREUR Then
                r_liberr = "Impossible de recopier le fichier " & FichLocal & " vers " & p_AdrServeur & " " & FichServeur & Chr(13) & Chr(10) & liberr
                HTTP_PutDos = HTTP_PUTDOS_ERREUR
            End If
        Next
        ' Lire les sous répertoires
        For Each Dossier In fso.GetFolder(v_CheminDosLocal).SubFolders
            Set fileItem = fso.GetFolder(Dossier)
            'MsgBox "dossier:" & Dossier.Name
            ' le créer
            iret = HTTP_creerDos(strHTTP, p_HTTP_CheminDepot, v_CheminDosServeur, Dossier.Name, r_liberr)
            If iret = HTTP_CREERDOS_EXISTE_DEJA Then
                HTTP_PutDos = HTTP_OK
                If v_bViderAvant Then
                    iret = HTTP_Appel_EffacerFichier(v_CheminDosServeur & "/" & Dossier.Name, False, r_liberr)
                End If
            ElseIf iret = HTTP_OK Or iret = HTTP_PUTDOS_DEJA Then
                HTTP_PutDos = HTTP_OK
            Else
                HTTP_PutDos = HTTP_PUTDOS_ERREUR
            End If
            
            Dim nomdest As String
            s = Mid(v_CheminDosServeur, Len(v_CheminDosServeur), 1)
            If s = "\" Or s = "/" Then
                nomdest = v_CheminDosServeur & Dossier.Name
            Else
                nomdest = v_CheminDosServeur & "/" & Dossier.Name
            End If
            s = Mid(v_CheminDosLocal, Len(v_CheminDosLocal), 1)
            If s = "\" Or s = "/" Then
                nomdos_loc = v_CheminDosLocal & Dossier.Name
            Else
                nomdos_loc = v_CheminDosLocal & "\" & Dossier.Name
            End If
            iret = HTTP_PutDos(nomdos_loc, nomdest, _
                                  False, False, r_liberr)
        Next
    Else
        r_liberr = "Chemin Local " & nomdos_loc & " inexistant"
        HTTP_PutDos = HTTP_PUTDOS_ERREUR
    End If
    
End Function

Public Sub HTTP_RecRmDir(ByVal vsFolder As Variant)
    
    ' destruction récursive d'un dossier Local (Client)
    
    Dim sName As Variant
    Dim oKillElements As Collection
    Dim ret As Integer
    
   On Local Error Resume Next
   If VarType(vsFolder) <> vbString Then
       Err.Raise 5
   Else
       If Right$(vsFolder, 1) = "\" Then
           vsFolder = left$(vsFolder, Len(vsFolder) - 1)
       End If
       Set oKillElements = New Collection
       sName = Dir$(vsFolder & "\*.*", vbDirectory Or vbReadOnly Or vbHidden Or vbSystem)
       Do While Len(sName)
            If (sName <> "..") And (sName <> ".") Then
                oKillElements.Add vsFolder & "\" & sName
           End If
           sName = Dir$()
       Loop
       For Each sName In oKillElements
           If GetAttr(sName) And vbDirectory Then
               HTTP_RecRmDir sName
               RmDir sName
           Else
               SetAttr sName, vbNormal
               Kill sName
           End If
       Next sName
       RmDir (vsFolder)
   End If

End Sub

Public Function HTTP_creerDos(ByVal v_sURL As String, _
                              ByVal v_chemin As String, _
                              ByVal v_CheminDossier_Serveur As String, _
                              ByVal v_NomDossier_Serveur As String, _
                              ByRef r_liberr As String) As Integer
    
    Dim stStatusCode As String, stStatusText As String
    Dim lgTotal As Long, sBuf As String
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim nomfich_Serveur As String
    Dim bresult As Boolean
    Dim Locker As String
    Dim http_req As S_HTTP_REQUEST
    
    r_liberr = ""
    
    ' Initialise Connect
    v_sURL = v_sURL
    ret = http_InitConnect(v_sURL, http_req)
    If CBool(ret) = False Then
        HTTP_creerDos = HTTP_PUTDOS_ERREUR
        Exit Function
    End If

    ' Send Request
    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_chemin & "&"
    stPost = stPost & "v_CheminDossier=" & v_CheminDossier_Serveur & "&"
    stPost = stPost & "v_NomDossier=" & v_NomDossier_Serveur & "&"
    stPost = stPost & "v_NumUtil=" & p_NumUtil
    ret = HttpSendRequest(http_req.lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        Call http_CloseConnect(http_req)
        r_liberr = "LockerFile : HttpSendRequest=0 : Apache arrêté ?"
        HTTP_creerDos = HTTP_PUTDOS_ERREUR
        Exit Function
    End If
    
    maxn = p_HTTP_MaxParPaquet
    
    ' Status Code
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(http_req.lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        Call http_CloseConnect(http_req)
        r_liberr = "CréerDos : Erreur HttpQueryInfo "
        HTTP_creerDos = HTTP_PUTDOS_ERREUR
        Exit Function
    End If
    If val(left(Trim(buf), 1)) > 3 Then
        Call http_CloseConnect(http_req)
        r_liberr = "CréerDos : Erreur HttpQueryInfo sret=" & left(Trim(buf), 3)
        HTTP_creerDos = HTTP_PUTDOS_ERREUR
        Exit Function
    End If
    
    ' Lecture
    sBuf = ""
    Do
        buf = String(p_HTTP_MaxParPaquet, Chr(0))
        ret = InternetReadFile(http_req.lHttpRequest, buf, p_HTTP_MaxParPaquet, n)
        sBuf = sBuf & left$(buf, n)
    Loop Until n = 0
    lgTotal = Len(sBuf)
    If left(sBuf, 6) = "ERREUR" Then
        r_liberr = STR_GetChamp(sBuf, "|", 2)
        r_liberr = Replace(r_liberr, "mod_", "")
        HTTP_creerDos = STR_GetChamp(sBuf, "|", 1)
    Else
        HTTP_creerDos = HTTP_OK
    End If
    
    Call http_CloseConnect(http_req)
    
End Function

Public Function HTTP_Ajouter_Tbl(ByVal v_nomLocal As String, _
                                 ByVal v_nomServeur, _
                                 ByVal v_Type_DosDoc, _
                                 ByVal v_Locké As Boolean) As Integer
    
    ' enregistrer dans le tableau les fichiers chargés par HTTP
    ' pour pouvoir les remettre ensuite
    Dim laDim As Integer
    Dim i As Integer
    Dim bDéjà As Boolean
    Dim CheminServeur As String
    Dim CheminTableau As String
    
    bDéjà = False
    If Not p_bool_HTTP_Fichiers_Chargés Then
        laDim = 0
    Else
        laDim = UBound(p_tbl_HTTP_Fichiers_Chargés(), 1)
        For i = 0 To laDim
            CheminTableau = p_tbl_HTTP_Fichiers_Chargés(i).HTTP_Fullname_Serveur
            CheminServeur = v_nomServeur
            'MsgBox CheminTableau
            If CheminServeur = CheminTableau Then
                bDéjà = True
                p_tbl_HTTP_Fichiers_Chargés(i).HTTP_Chargé = True
                Exit For
            End If
        Next i
        laDim = UBound(p_tbl_HTTP_Fichiers_Chargés(), 1) + 1
    End If
    If Not bDéjà Then
        ReDim Preserve p_tbl_HTTP_Fichiers_Chargés(laDim)
        p_tbl_HTTP_Fichiers_Chargés(laDim).HTTP_Chargé = True
        p_tbl_HTTP_Fichiers_Chargés(laDim).HTTP_Fullname_Local = v_nomLocal
        p_tbl_HTTP_Fichiers_Chargés(laDim).HTTP_Fullname_Serveur = v_nomServeur
        p_tbl_HTTP_Fichiers_Chargés(laDim).HTTP_Locké = v_Locké
        p_tbl_HTTP_Fichiers_Chargés(laDim).HTTP_Type_DosDoc = v_Type_DosDoc
        p_bool_HTTP_Fichiers_Chargés = True
        HTTP_Ajouter_Tbl = laDim
    Else
        p_tbl_HTTP_Fichiers_Chargés(i).HTTP_Chargé = True
        p_tbl_HTTP_Fichiers_Chargés(i).HTTP_Locké = v_Locké
        HTTP_Ajouter_Tbl = i
    End If
    
End Function

