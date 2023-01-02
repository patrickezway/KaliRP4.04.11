VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End

'======================================================================================
' Classe : clsDir
' Auteur : Jack
' Source : http://www.vbfrance.com/code.aspx?ID=37859
'======================================================================================

' Cette classe permet de remplacer la fonction DIR native de VB.
' Elle possède (presque) les mêmes paramètres et principe d'utilisation
'   + quelques avantages

' Attention : Contrairement à la fonction DIR, cette classe ne renvoie QUE les
'               résultats dans le type demandé.
'             Les paramètres du DIR classique listent les fichiers
'               demandés + les fichiers sans attributs (je trouve ça un peu galère)
'             Avec l'instruction DIR classique, l'exemple suivant ne renvoie pas
'               que les répertoires, il faut retester le résultat fourni avec
'               un GetAttr pour en être sûr.
'                      sTemp = Dir("C:\mon Répertoire\*.*", vbDirectory)
'                      If GetAttr("C:\mon Répertoire\" & sTemp) And vbDirectory Then ...
'             (idem avec vbHidden ...)

' Exemple d'utilisation de DIR traditionnel :
'     sFichier = Dir("C:\mon répertoire\*.txt")
'     Do While Len(sFichier) > 0
'         Debug.Print sFichier
'         sFichier = Dir
'         DoEvents
'     Loop

' La fonction DIR classique est très pratique puisque, pour lister les fichiers
'   suivants, il suffit de faire un simple  'Chaine = Dir'.
'   On va garder son principe bien pratique.
' Les zinconvénients de cette fonction DIR classique sont les suivants :
'   - Il est impossible d'imbriquer deux recherches dans une même boucle.
'     Exemple de Bug pas toujours facile à déceler :
'       Vous ètes en train de faire votre boucle de lecture (Voir exemple ci-dessus)
'       Comme il y a un 'DoEvents', le programme redonne la main au système qui va
'         pouvoir gérer son affichage, son clavier ainsi que les évènements qui se
'         déclenchent.
'       Si par exemple un Timer de votre application se déclenche et que lui-même
'         utilise aussi la fonction 'Dir' : papatra, le 'Dir' de votre 1ère boucle
'         a été réinitialisé et ne pointe plus sur le bon répertoire !
'   - Impossible de dissocier (nativement) les répertoires des fichiers marqués 'Normal'
'       (voir le 'Attention' plus haut)

' L'utilisation d'une Classe permet de créer une nouvelle instance à chaque
'   utilisation et par conséquent de personnaliser le paramétrage de chaque instance.

' Utilisation de la Classe :
'     Dim monDir As New clsDir2
'     sFichier = monDir.Dir2("C:\mon répertoire\*.txt")
'     Do While Len(sFichier) > 0
'         Debug.Print sFichier
'         sFichier = monDir.Dir2
'         DoEvents
'     Loop
'     Set monDir = Nothing
'
'======================================================================================
Option Explicit

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * 260
    cAlternate          As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
                                    ByVal lpFileName As String, _
                                    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
                                    ByVal hFindFile As Long, _
                                    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" ( _
                                    ByVal hFindFile As Long) As Long

' Liste des types de recherche proposés optionnellement dans la fonction 'Dir2'
Public Enum typeRecherche
    [Fichier standard] = 0                  ' Option par défaut
    [Fichier caché] = 1
    [Fichier système] = 2
    [Fichier lecture seule] = 3
    [Fichier compressé (sur disque)] = 4
    [Répertoire standard] = 5
    [Répertoire caché] = 6
    [Répertoire système] = 7
    [Répertoire lecture seule] = 8
    [Répertoire compressé (sur disque)] = 9
End Enum
' Attributs (additionnables) d'un fichier/répertoire renvoyés par les APIs
Private Const mNormal = 128                 ' Fichier sans attribut
Private Const mArchive = 32                 ' Fichier modifié
Private Const mLectureSeule = 1
Private Const mCaché = 2
Private Const mSystème = 4
Private Const mRépertoire = 16
Private Const mCompressionDisque = 2048     ' Compressé au sens Windows/disque (pas Zip, RAR ...)

' Les variables
Private mHandle         As Long             ' Mémo de handle de la recherche en cours
Private mRechercheQuoi  As Long             ' Type de recherche à effectuer
Private aFileStructure  As WIN32_FIND_DATA  ' Elément récupéré par les APIs
'

Public Function Dir2(Optional sRépertoire As String = "", _
                     Optional lRechercheQuoi As typeRecherche = [Fichier standard]) As String

    ' Renvoie :
    ' - Le nom du premier fichier du répertoire si sRépertoire est fourni
    ' - Le nom du fichier suivant si sRépertoire n'est pas fourni

    Dim sTemp As String
    
    ' Est-ce qu'on a fourni le répertoire ?
    If Len(sRépertoire) > 0 Then
        ' Alors, il s'agit d'une initialisation
        '--------------------------------------
        ' Referme une éventuelle précédente recherche
        If mHandle > 0 Then Call FindClose(mHandle)
        ' Ouverture de session
        mHandle = FindFirstFile(sRépertoire, aFileStructure)
        ' On ressort si la fonction a échoué (renvoie -1)
        If mHandle <= 0 Then Exit Function
        ' Mémorise le paramètre du filtre de type
        mRechercheQuoi = lRechercheQuoi
        ' Isole le nom du fichier/répertoire
        sTemp = IsoleString(aFileStructure.cFileName)
        If Len(sTemp) = 0 Then Exit Function
        ' En tête de recherche, on a tjrs les "." et ".." (sauf sur une racine)
        If sTemp = ".." Or sTemp = "." Then
            Dir2 = FichierSuivant
        Else
            If TestTypeFichier Then
                ' Le fichier est du même type que celui recherché (youpi)
                Dir2 = sTemp
            Else
                ' Pas le bon type : Continue la recherche
                Dir2 = FichierSuivant
            End If
        End If
    Else
        ' Pas de paramètre --> On veut connaitre le fichier suivant
        '----------------------------------------------------------
        ' Recherche occurence suivante (c'est bien dit, hein ?)
        Dir2 = FichierSuivant
    End If

End Function

Private Function FichierSuivant() As String

    ' Cette routine interne à la classe permet de rechercher le fichier suivant
    
    Dim sTemp As String
    
    ' Bah, qu'est c'est qui fait ? Il a pas initialiser la recherche !
    If mHandle <= 0 Then Exit Function
    
    Do While 1 = 1      ' On ne ressort qu'avec un 'Exit Do'
        If FindNextFile(mHandle, aFileStructure) Then
            ' Récupère le nom du fichier
            sTemp = IsoleString(aFileStructure.cFileName)
            If Len(sTemp) = 0 Then Exit Do  ' L'a pô trouvé
            ' Si ce n'est pas l'entête de répertoire,
            If sTemp <> ".." And sTemp <> "." Then
                ' Vérifie qu'il est du type recherché
                If TestTypeFichier Then
                    ' Ca y est, on en tient un
                    FichierSuivant = sTemp
                    Exit Do
                End If
            End If
        Else
            ' Pas de fichier suivant trouvé. On sort
            Exit Do
        End If
    Loop
    
End Function

Private Function TestTypeFichier() As Boolean

    ' Cette fonction renvoie True si le type de fichier correspond
    '   au type recherché
    ' Le type de fichier renvoyé par les APIs est presque identique à celui
    '   définit dans VB dans VbFileAttribute, sauf :
    '                 'Normal'  'Compressed'
    '                ---------  ------------
    '     Sous VB         0     n'existe pas
    '     Avec API       128        2048

    ' La différence entre Normal et Archive est assez nébuleuse.
    ' D'après certaines docs :
    '   Normal  : Fichier sans attribut
    '   Archive : Item modifié
    ' Bof, des fichiers EXE en Archive, je ne comprends pas trop

    ' La variable à comparer est toujours aFileStructure.dwFileAttributes
    '   donc, pas utile de le transmettre comme paramètre
    ' Le choix de comparaison est dans mRechercheQuoi
    
    ' Transfère l'attribut à tester - syntaxe plus courte pour la suite
    Dim Attribut As Long
    Attribut = aFileStructure.dwFileAttributes
    
    Select Case mRechercheQuoi
    
        '************* Fichiers *************
        Case [Fichier standard]
            ' Normal et/ou Archive
            ' On autorise aussi les LectureSeule
            '   mais pas les Caché ni Système
            aFileStructure.cFileName = aFileStructure.cFileName
            Attribut = Attribut
            If Not CompareAttribut(mRépertoire) And _
               (CompareAttribut(mNormal) Or CompareAttribut(mArchive)) And _
               Not CompareAttribut(mCaché) And _
               Not CompareAttribut(mSystème) Then
                    TestTypeFichier = True
            End If
        Case [Fichier caché]
            ' On autorise les Système et LectureSeule
            If Not CompareAttribut(mRépertoire) And _
               CompareAttribut(mCaché) Then
                    TestTypeFichier = True
            End If
        Case [Fichier système]
            ' On autorise les Cachés et LectureSeule
            If Not CompareAttribut(mRépertoire) And _
               CompareAttribut(mSystème) Then
                    TestTypeFichier = True
            End If
        Case [Fichier lecture seule]
            ' On autorise les Cachés et Système
            If Not CompareAttribut(mRépertoire) And _
               CompareAttribut(mLectureSeule) Then
                    TestTypeFichier = True
            End If
        Case [Fichier compressé (sur disque)]
            ' On autorise les Caché, Système et LectureSeule
            If Not CompareAttribut(mRépertoire) And _
               CompareAttribut(mCompressionDisque) Then
                    TestTypeFichier = True
            End If
            
        '************* Répertoires *************
        ' Pas de notion de 'Normal' avec les répertoires
        Case [Répertoire standard]
            ' Archive
            ' On autorise les Archives, Lecture seule et Compressé
            ' On n'autorise pas les Caché ni Système
            If CompareAttribut(mRépertoire) And _
               Not CompareAttribut(mCaché) And _
               Not CompareAttribut(mSystème) Then
                    TestTypeFichier = True
            End If
        Case [Répertoire caché]
            ' On autorise les Système, LectureSeule et Compressé
            If CompareAttribut(mRépertoire) And _
               CompareAttribut(mCaché) Then
                    TestTypeFichier = True
            End If
        Case [Répertoire système]
            ' On autorise les Cachés, LectureSeule et Compressé
            If CompareAttribut(mRépertoire) And _
               CompareAttribut(mSystème) Then
                    TestTypeFichier = True
            End If
        Case [Répertoire lecture seule]
            ' On autorise les Cachés, Système et Compressé
            If CompareAttribut(mRépertoire) And _
               CompareAttribut(mLectureSeule) Then
                    TestTypeFichier = True
            End If
        Case [Répertoire compressé (sur disque)]
            ' On autorise les Caché, Système et LectureSeule
            If CompareAttribut(mRépertoire) And _
               CompareAttribut(mCompressionDisque) Then
                    TestTypeFichier = True
            End If
    End Select
    
End Function

Private Function CompareAttribut(Attribut_a_Comparer As Long) As Boolean

    ' Quand on veut comparer bit à bit deux mots, il est facile d'utiliser la syntaxe
    '     Mot1 And Mot2     qui renverra Mot2 si le bit correspondant (2^x) est monté
    '   Exemples : 38 And 4 renvoie 4    (38 = 32 + 4 + 2)
    '            : 35 And 4 renvoie 0    (35 = 32 + 2 + 1)
    ' Par contre, pour faire un Not, ça ne marche plus, car le Not est interprété
    '   'numériquement', et pas 'logiquement'.
    '   Exemples : Not(38 And 4) renvoie -5 --> Considéré comme True dans un If
    '              Not(35 And 4) renvoie -1 --> Considéré comme True dans un If
    ' Il suffit donc de convertir l'équation dans le bon type Boolean avant de
    '   faire le Not.
    ' On va le faire ici pour éviter de surcharger la syntaxe de la Fonction précédente

    ' La variable à comparer est toujours aFileStructure.dwFileAttributes
    '   donc, pas utile de le transmettre comme paramètre

    CompareAttribut = CBool(aFileStructure.dwFileAttributes And Attribut_a_Comparer)

End Function

Private Function IsoleString(sChaine As String) As String

    ' Le résultat de cetaines API (comme ici) renvoie une chaine de longueur
    '   définie à l'avance (260). Dans ce cas, la fin de la partie de texte
    '   interressante est marquée par un Chr$(0) = caractère ascii 0
    ' Ici, on va isoler le texte interressant
    
    Dim lPositionChar As Long
    
    lPositionChar = InStr(sChaine, Chr$(0))     ' Chr(0) final
    If lPositionChar > 0 Then IsoleString = left$(sChaine, lPositionChar - 1)
    
End Function

Private Sub Class_Terminate()

    ' Cloture la recherche proprement
    If mHandle > 0 Then Call FindClose(mHandle)
    
End Sub






