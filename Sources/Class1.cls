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
' Elle poss?de (presque) les m?mes param?tres et principe d'utilisation
'   + quelques avantages

' Attention : Contrairement ? la fonction DIR, cette classe ne renvoie QUE les
'               r?sultats dans le type demand?.
'             Les param?tres du DIR classique listent les fichiers
'               demand?s + les fichiers sans attributs (je trouve ?a un peu gal?re)
'             Avec l'instruction DIR classique, l'exemple suivant ne renvoie pas
'               que les r?pertoires, il faut retester le r?sultat fourni avec
'               un GetAttr pour en ?tre s?r.
'                      sTemp = Dir("C:\mon R?pertoire\*.*", vbDirectory)
'                      If GetAttr("C:\mon R?pertoire\" & sTemp) And vbDirectory Then ...
'             (idem avec vbHidden ...)

' Exemple d'utilisation de DIR traditionnel :
'     sFichier = Dir("C:\mon r?pertoire\*.txt")
'     Do While Len(sFichier) > 0
'         Debug.Print sFichier
'         sFichier = Dir
'         DoEvents
'     Loop

' La fonction DIR classique est tr?s pratique puisque, pour lister les fichiers
'   suivants, il suffit de faire un simple  'Chaine = Dir'.
'   On va garder son principe bien pratique.
' Les zinconv?nients de cette fonction DIR classique sont les suivants :
'   - Il est impossible d'imbriquer deux recherches dans une m?me boucle.
'     Exemple de Bug pas toujours facile ? d?celer :
'       Vous ?tes en train de faire votre boucle de lecture (Voir exemple ci-dessus)
'       Comme il y a un 'DoEvents', le programme redonne la main au syst?me qui va
'         pouvoir g?rer son affichage, son clavier ainsi que les ?v?nements qui se
'         d?clenchent.
'       Si par exemple un Timer de votre application se d?clenche et que lui-m?me
'         utilise aussi la fonction 'Dir' : papatra, le 'Dir' de votre 1?re boucle
'         a ?t? r?initialis? et ne pointe plus sur le bon r?pertoire !
'   - Impossible de dissocier (nativement) les r?pertoires des fichiers marqu?s 'Normal'
'       (voir le 'Attention' plus haut)

' L'utilisation d'une Classe permet de cr?er une nouvelle instance ? chaque
'   utilisation et par cons?quent de personnaliser le param?trage de chaque instance.

' Utilisation de la Classe :
'     Dim monDir As New clsDir2
'     sFichier = monDir.Dir2("C:\mon r?pertoire\*.txt")
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

' Liste des types de recherche propos?s optionnellement dans la fonction 'Dir2'
Public Enum typeRecherche
    [Fichier standard] = 0                  ' Option par d?faut
    [Fichier cach?] = 1
    [Fichier syst?me] = 2
    [Fichier lecture seule] = 3
    [Fichier compress? (sur disque)] = 4
    [R?pertoire standard] = 5
    [R?pertoire cach?] = 6
    [R?pertoire syst?me] = 7
    [R?pertoire lecture seule] = 8
    [R?pertoire compress? (sur disque)] = 9
End Enum
' Attributs (additionnables) d'un fichier/r?pertoire renvoy?s par les APIs
Private Const mNormal = 128                 ' Fichier sans attribut
Private Const mArchive = 32                 ' Fichier modifi?
Private Const mLectureSeule = 1
Private Const mCach? = 2
Private Const mSyst?me = 4
Private Const mR?pertoire = 16
Private Const mCompressionDisque = 2048     ' Compress? au sens Windows/disque (pas Zip, RAR ...)

' Les variables
Private mHandle         As Long             ' M?mo de handle de la recherche en cours
Private mRechercheQuoi  As Long             ' Type de recherche ? effectuer
Private aFileStructure  As WIN32_FIND_DATA  ' El?ment r?cup?r? par les APIs
'

Public Function Dir2(Optional sR?pertoire As String = "", _
                     Optional lRechercheQuoi As typeRecherche = [Fichier standard]) As String

    ' Renvoie :
    ' - Le nom du premier fichier du r?pertoire si sR?pertoire est fourni
    ' - Le nom du fichier suivant si sR?pertoire n'est pas fourni

    Dim sTemp As String
    
    ' Est-ce qu'on a fourni le r?pertoire ?
    If Len(sR?pertoire) > 0 Then
        ' Alors, il s'agit d'une initialisation
        '--------------------------------------
        ' Referme une ?ventuelle pr?c?dente recherche
        If mHandle > 0 Then Call FindClose(mHandle)
        ' Ouverture de session
        mHandle = FindFirstFile(sR?pertoire, aFileStructure)
        ' On ressort si la fonction a ?chou? (renvoie -1)
        If mHandle <= 0 Then Exit Function
        ' M?morise le param?tre du filtre de type
        mRechercheQuoi = lRechercheQuoi
        ' Isole le nom du fichier/r?pertoire
        sTemp = IsoleString(aFileStructure.cFileName)
        If Len(sTemp) = 0 Then Exit Function
        ' En t?te de recherche, on a tjrs les "." et ".." (sauf sur une racine)
        If sTemp = ".." Or sTemp = "." Then
            Dir2 = FichierSuivant
        Else
            If TestTypeFichier Then
                ' Le fichier est du m?me type que celui recherch? (youpi)
                Dir2 = sTemp
            Else
                ' Pas le bon type : Continue la recherche
                Dir2 = FichierSuivant
            End If
        End If
    Else
        ' Pas de param?tre --> On veut connaitre le fichier suivant
        '----------------------------------------------------------
        ' Recherche occurence suivante (c'est bien dit, hein ?)
        Dir2 = FichierSuivant
    End If

End Function

Private Function FichierSuivant() As String

    ' Cette routine interne ? la classe permet de rechercher le fichier suivant
    
    Dim sTemp As String
    
    ' Bah, qu'est c'est qui fait ? Il a pas initialiser la recherche !
    If mHandle <= 0 Then Exit Function
    
    Do While 1 = 1      ' On ne ressort qu'avec un 'Exit Do'
        If FindNextFile(mHandle, aFileStructure) Then
            ' R?cup?re le nom du fichier
            sTemp = IsoleString(aFileStructure.cFileName)
            If Len(sTemp) = 0 Then Exit Do  ' L'a p? trouv?
            ' Si ce n'est pas l'ent?te de r?pertoire,
            If sTemp <> ".." And sTemp <> "." Then
                ' V?rifie qu'il est du type recherch?
                If TestTypeFichier Then
                    ' Ca y est, on en tient un
                    FichierSuivant = sTemp
                    Exit Do
                End If
            End If
        Else
            ' Pas de fichier suivant trouv?. On sort
            Exit Do
        End If
    Loop
    
End Function

Private Function TestTypeFichier() As Boolean

    ' Cette fonction renvoie True si le type de fichier correspond
    '   au type recherch?
    ' Le type de fichier renvoy? par les APIs est presque identique ? celui
    '   d?finit dans VB dans VbFileAttribute, sauf :
    '                 'Normal'  'Compressed'
    '                ---------  ------------
    '     Sous VB         0     n'existe pas
    '     Avec API       128        2048

    ' La diff?rence entre Normal et Archive est assez n?buleuse.
    ' D'apr?s certaines docs :
    '   Normal  : Fichier sans attribut
    '   Archive : Item modifi?
    ' Bof, des fichiers EXE en Archive, je ne comprends pas trop

    ' La variable ? comparer est toujours aFileStructure.dwFileAttributes
    '   donc, pas utile de le transmettre comme param?tre
    ' Le choix de comparaison est dans mRechercheQuoi
    
    ' Transf?re l'attribut ? tester - syntaxe plus courte pour la suite
    Dim Attribut As Long
    Attribut = aFileStructure.dwFileAttributes
    
    Select Case mRechercheQuoi
    
        '************* Fichiers *************
        Case [Fichier standard]
            ' Normal et/ou Archive
            ' On autorise aussi les LectureSeule
            '   mais pas les Cach? ni Syst?me
            aFileStructure.cFileName = aFileStructure.cFileName
            Attribut = Attribut
            If Not CompareAttribut(mR?pertoire) And _
               (CompareAttribut(mNormal) Or CompareAttribut(mArchive)) And _
               Not CompareAttribut(mCach?) And _
               Not CompareAttribut(mSyst?me) Then
                    TestTypeFichier = True
            End If
        Case [Fichier cach?]
            ' On autorise les Syst?me et LectureSeule
            If Not CompareAttribut(mR?pertoire) And _
               CompareAttribut(mCach?) Then
                    TestTypeFichier = True
            End If
        Case [Fichier syst?me]
            ' On autorise les Cach?s et LectureSeule
            If Not CompareAttribut(mR?pertoire) And _
               CompareAttribut(mSyst?me) Then
                    TestTypeFichier = True
            End If
        Case [Fichier lecture seule]
            ' On autorise les Cach?s et Syst?me
            If Not CompareAttribut(mR?pertoire) And _
               CompareAttribut(mLectureSeule) Then
                    TestTypeFichier = True
            End If
        Case [Fichier compress? (sur disque)]
            ' On autorise les Cach?, Syst?me et LectureSeule
            If Not CompareAttribut(mR?pertoire) And _
               CompareAttribut(mCompressionDisque) Then
                    TestTypeFichier = True
            End If
            
        '************* R?pertoires *************
        ' Pas de notion de 'Normal' avec les r?pertoires
        Case [R?pertoire standard]
            ' Archive
            ' On autorise les Archives, Lecture seule et Compress?
            ' On n'autorise pas les Cach? ni Syst?me
            If CompareAttribut(mR?pertoire) And _
               Not CompareAttribut(mCach?) And _
               Not CompareAttribut(mSyst?me) Then
                    TestTypeFichier = True
            End If
        Case [R?pertoire cach?]
            ' On autorise les Syst?me, LectureSeule et Compress?
            If CompareAttribut(mR?pertoire) And _
               CompareAttribut(mCach?) Then
                    TestTypeFichier = True
            End If
        Case [R?pertoire syst?me]
            ' On autorise les Cach?s, LectureSeule et Compress?
            If CompareAttribut(mR?pertoire) And _
               CompareAttribut(mSyst?me) Then
                    TestTypeFichier = True
            End If
        Case [R?pertoire lecture seule]
            ' On autorise les Cach?s, Syst?me et Compress?
            If CompareAttribut(mR?pertoire) And _
               CompareAttribut(mLectureSeule) Then
                    TestTypeFichier = True
            End If
        Case [R?pertoire compress? (sur disque)]
            ' On autorise les Cach?, Syst?me et LectureSeule
            If CompareAttribut(mR?pertoire) And _
               CompareAttribut(mCompressionDisque) Then
                    TestTypeFichier = True
            End If
    End Select
    
End Function

Private Function CompareAttribut(Attribut_a_Comparer As Long) As Boolean

    ' Quand on veut comparer bit ? bit deux mots, il est facile d'utiliser la syntaxe
    '     Mot1 And Mot2     qui renverra Mot2 si le bit correspondant (2^x) est mont?
    '   Exemples : 38 And 4 renvoie 4    (38 = 32 + 4 + 2)
    '            : 35 And 4 renvoie 0    (35 = 32 + 2 + 1)
    ' Par contre, pour faire un Not, ?a ne marche plus, car le Not est interpr?t?
    '   'num?riquement', et pas 'logiquement'.
    '   Exemples : Not(38 And 4) renvoie -5 --> Consid?r? comme True dans un If
    '              Not(35 And 4) renvoie -1 --> Consid?r? comme True dans un If
    ' Il suffit donc de convertir l'?quation dans le bon type Boolean avant de
    '   faire le Not.
    ' On va le faire ici pour ?viter de surcharger la syntaxe de la Fonction pr?c?dente

    ' La variable ? comparer est toujours aFileStructure.dwFileAttributes
    '   donc, pas utile de le transmettre comme param?tre

    CompareAttribut = CBool(aFileStructure.dwFileAttributes And Attribut_a_Comparer)

End Function

Private Function IsoleString(sChaine As String) As String

    ' Le r?sultat de cetaines API (comme ici) renvoie une chaine de longueur
    '   d?finie ? l'avance (260). Dans ce cas, la fin de la partie de texte
    '   interressante est marqu?e par un Chr$(0) = caract?re ascii 0
    ' Ici, on va isoler le texte interressant
    
    Dim lPositionChar As Long
    
    lPositionChar = InStr(sChaine, Chr$(0))     ' Chr(0) final
    If lPositionChar > 0 Then IsoleString = left$(sChaine, lPositionChar - 1)
    
End Function

Private Sub Class_Terminate()

    ' Cloture la recherche proprement
    If mHandle > 0 Then Call FindClose(mHandle)
    
End Sub






