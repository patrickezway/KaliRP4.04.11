Attribute VB_Name = "ModuleDir"
'*********************************************************************************************************************'
'*********************************************************************************************************************'
'**                                                                                                                 **'
'**                                             DIRECTORY LISTING CLASS                                             **'
'**                                                                                                                 **'
'*********************************************************************************************************************'
'*********************************************************************************************************************'
'--------------------------------------------------   ATTRIBUTES   ---------------------------------------------------'
'Author:        Santiago Diez
'Email:         santiago.diez@caoba.fr
'Website:       http://santiago.diez.free.fr
'Date:          2008-05-06  07:28:20
'Version:       1.1.2
'Copyright:     None
'Description:   This class provides a powerfull and fast way to list files in a directory.
'Bugs:          No bug reported
'Sources:       Renfield's file listing Class (http://www.vbfrance.com/code.aspx?ID=43640)
'Requirements:  msvbvm60.dll, VB6.OLB, VB6FR.DLL (Always required)
'----------------------------------------------------   OPTIONS   ----------------------------------------------------'
Option Base 1
Option Compare Binary
Option Explicit
'---------------------------------------------------   CONSTANTS   ---------------------------------------------------'
Const SELF_DIRECTORY As String = "." & vbNullChar & vbNullChar
Const SELF_DIRECTORY_LEN As Long = 3
Const PARENT_DIRECTORY As String = "." & vbNullChar & "." & vbNullChar & vbNullChar
Const PARENT_DIRECTORY_LEN As Long = 5
Const MAXDEPTH As Long = 32
'------------------------------------------------   ENUMS AND TYPES   ------------------------------------------------'
'dlFileAttributes Constants
'   Constants used to enumerate file attributes.
'---------------------------------------------------------------------------------------------------------------------'
Enum dlFileAttributes
    dlReadOnly = 1
    dlHidden = 2
    dlSystem = 4
    'dlVolume = 8           'Not supported
    dlDirectory = 16
    dlArchive = 32
    'dlAlias = 64           'Not supported
    dlNormal = 128
    dlTemporary = 256
    dlSparseFile = 512
    dlReparsePoint = 1024
    dlCompressed = 2048
    dlOffline = 4096
    dlNotIndexed = 8192
    dlEncrypted = 16384
    dlRecursive = 32768     'Not applicable to file attribute. Just makes the search recursive
End Enum
'---------------------------------------------------------------------------------------------------------------------'
'SYSTEMTIME Data Type
'   Type used to represent date/time in the system.
'---------------------------------------------------------------------------------------------------------------------'
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'---------------------------------------------------------------------------------------------------------------------'
'FILETIME Data Type
'   Type used in WIN32_FIND_DATA Data Type.
'---------------------------------------------------------------------------------------------------------------------'
Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type
'---------------------------------------------------------------------------------------------------------------------'
'WIN32_FIND_DATA Data Type
'   Type used in declared function FindFirstFileW.
'---------------------------------------------------------------------------------------------------------------------'
Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * 520
    cAlternate          As String * 28
End Type
'---------------------------------------------------------------------------------------------------------------------'
'Search Data Type
'   Type used to stack recursive searches.
'---------------------------------------------------------------------------------------------------------------------'
Private Type Search
    Handle      As Long
    Path        As String
End Type
'-----------------------------------------------   GLOBAL VARIABLES   ------------------------------------------------'
Dim SearchMask As String            'Search mask (eg. "*.cls")
Dim SearchAttributes As Long        'Search attributes
Dim SearchRecursive As Long         'Search in subfolders (zero or cdRecursive)
Dim Search() As Search              'Stack of recursive searches
Dim CurSearch As Long               'Position of the current search in the stack
Dim fileinfo As WIN32_FIND_DATA     'Description of a found file
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                      EVENTS                                                       +'
'+-------------------------------------------------------------------------------------------------------------------+'
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                 DECLARED FUNCTIONS                                                +'
'+-------------------------------------------------------------------------------------------------------------------+'
'FindFirstFileW Function
'   <Description>
'       Searches a directory for a file or subdirectory with a name that matches a specific name.
'   <Syntax>
'       FindFirstFileW(*pathname*, *fileinfo*)
'   <Parameters>
'       *pathname*  Required, Long Pointer  to a String. The  directory or path, and  the file name, which  can include
'                   wildcard characters, for example, an asterisk (*) or a question mark (?).
'       *fileinfo*  Required. A pointer to  the WIN32_FIND_DATA structure that receives information  about a found file
'                   or subdirectory.
'   <Returned value>
'       If the function  succeeds, the return value  is a search handle used  in a subsequent call  to FindNextFileW or
'       FindClose. If it fails, the return value is INVALID_HANDLE_VALUE (-1).
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal PathName As Long, fileinfo As WIN32_FIND_DATA) As Long
'---------------------------------------------------------------------------------------------------------------------'
'FindNextFile Function
'   <Description>
'       Continues a file search from a previous call to the FindFirstFileW function.
'   <Syntax>
'       FindNextFileW(*SearchHandle*, *fileinfo*)
'   <Parameters>
'       *SearchHandle*  Required, Long. The search handle returned by a previous call to the FindFirstFileW function.
'       *fileinfo*      Required. A pointer to the WIN32_FIND_DATA structure  that receives information about the found
'                       file or subdirectory.
'   <Returned value>
'       If the function succeeds, the return value is nonzero. If it fails, the return value is zero (0).
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Function FindNextFileW Lib "kernel32" (ByVal SearchHandle As Long, fileinfo As WIN32_FIND_DATA) As Long
'---------------------------------------------------------------------------------------------------------------------'
'FindClose Function
'   <Description>
'       Closes a file search handle opened by the FindFirstFileW function.
'   <Syntax>
'       FindClose(*SearchHandle*)
'   <Parameters>
'       *SearchHandle*  Required, Long. The search handle returned by a previous call to the FindFirstFileW function.
'   <Returned value>
'       If the function succeeds, the return value is nonzero. If it fails, the return value is zero (0).
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Function FindClose Lib "kernel32" (ByVal SearchHandle As Long) As Long
'---------------------------------------------------------------------------------------------------------------------'
'FileTimeToLocalFileTime Function
'   <Description>
'       Converts a file time to a local file time.
'   <Syntax>
'       FileTimeToLocalFileTime(*FileTime_*, *LocalFileTime_*)
'   <Parameters>
'       *FileTime_*         Required.  A pointer  to a  FILETIME structure  containing the  UTC-based file  time to  be
'                           converted into a local file time.
'       *LocalFileTime_*    Required. A pointer to a FILETIME structure to  receive the converted local file time. This
'                           parameter cannot be the same as the *FileTime_* parameter.
'   <Returned value>
'       If the function succeeds, the return value is nonzero. If it fails, the return value is zero.
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef FileTime_ As FILETIME, ByRef LocalFileTime_ As _
FILETIME) As Long
'---------------------------------------------------------------------------------------------------------------------'
'FileTimeToSystemTime Function
'   <Description>
'       Converts a file time to system time format.
'   <Syntax>
'       FileTimeToSystemTime(*FileTime_*, *SystemTime_*)
'   <Parameters>
'       *FileTime_*     Required. A pointer to a FILETIME structure containing  the file time to convert to system date
'                       and time format.
'       *SystemTime_*   Required. A pointer to a SYSTEMTIME structure to receive the converted file time.
'   <Returned value>
'       If the function succeeds, the return value is nonzero. If it fails, the return value is zero.
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef FileTime_ As FILETIME, ByRef SystemTime_ As _
SYSTEMTIME) As Long
'---------------------------------------------------------------------------------------------------------------------'
'RtlMoveMemory Method
'   <Description>
'       The RtlMoveMemory  routine moves memory  either forward or  backward, aligned  or unaligned, in  4-byte blocks,
'       followed by any remaining bytes.
'   <Syntax>
'       RtlMoveMemory *Destination*, *Source*, *Length*
'   <Parameters>
'       *Destination*   Required. Pointer to the destination of the move.
'       *Source*        Required. Pointer to the memory to be copied.
'       *Lenth*         Required. Specifies the number of bytes to be copied.
'---------------------------------------------------------------------------------------------------------------------'
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                  ERROR HANDLING                                                   +'
'+-------------------------------------------------------------------------------------------------------------------+'
'ErrRaise Method
'   <Description>
'       Raises a class specific, user defined error.
'   <Syntax>
'       ErrRaise *number*
'   <Parameters>
'       *number*    Required, Long. Integer that  identifies the nature of the error. The  range 513-65535 is available
'                   for user-defined errors. Outside this range, a standard Visual Basic error is generated.
'---------------------------------------------------------------------------------------------------------------------'
Private Sub ErrRaise(number As Long)
    Const Source As String = "DirListing"
    Dim description As String
    Select Case number
        Case Is < 513, Is > 65535
            Err.Raise number
        Case 513
            description = "Incorrect number of arguments"
        Case 514
            description = "No more files"
        Case Else
            description = "Undescribed error"
    End Select
    Err.Raise vbObjectError + number, Source, description
End Sub
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                     FUNCTIONS                                                     +'
'+-------------------------------------------------------------------------------------------------------------------+'
'GetStrFromUnicode Function
'   <Description>
'       Transform a unicode null terminated string into a readable string.
'   <Syntax>
'       GetStrFromUnicode(*Str*)
'   <Parameters>
'       *Str*   Required. The unicode string that will b evaluated
'---------------------------------------------------------------------------------------------------------------------'
Private Function GetStrFromUnicode(Str As String) As String
    GetStrFromUnicode = StrConv(left$(Str, InStr(Str, String$(2, 0))), vbFromUnicode)
End Function
'---------------------------------------------------------------------------------------------------------------------'
'SplitPathName Method
'   <Description>
'       Splits a path (relative or absolute) into two parts:  1) the absolute, fully qualified path without the name or
'       mask. 2) the name or mask.
'   <Syntax>
'       SplitPathName *PathName*, *Path*, *Mask*
'   <Parameters>
'       *PathName*  Required. Path specification to split.
'       *Path*      Required. A string variable in whitch the path will be stored.
'       *Mask*      Required. A string variable in whitch the mask will be stored.
'---------------------------------------------------------------------------------------------------------------------'
Private Sub SplitPathName(ByVal PathName As String, Path As String, Mask As String)
    Dim i As Long, fileparent As WIN32_FIND_DATA
    'Build absolute path
    If left$(PathName, 2) = "\\" Then       'This is a UNC path, there is nothing to append
    ElseIf left$(PathName, 1) = "\" Then    'This is a relative path from the root of the current drive
        PathName = left$(CurDir$, 2) & PathName
    ElseIf InStr(PathName, ":") Then        'This is an absolute path, there is nothing to append
        'Change every ":" into ":\" because Dir("C:WINDOWS") is interpreted as Dir("C:\WINDOWS")
        'If there are other ":", it will anyway lead to no find items.
        'If there's already a ":\" it will become ":\\" which will anyway be interpreted as ":\"
        PathName = Replace(PathName, ":", ":\")
    Else                                    'This is a relative path from the current directory
        PathName = CurDir$ & "\" & PathName
    End If
    'Get mask and pathname's parent folder
    If Right$(PathName, 1) = "\" Then
        Mask = "*"
        PathName = PathName & "."
    ElseIf Right$(PathName, 2) = "\." Then
        Mask = "*"
        'PathName = PathName
    ElseIf Right$(PathName, 3) = "\.." Then
        Mask = "*"
        'PathName = PathName
    Else
        i = InStrRev(PathName, "\")
        If i Then
            Mask = Mid$(PathName, i + 1)
            PathName = left$(PathName, i) & "."
        Else    'Cases like "C:" or "D:"
            Mask = "*"
            PathName = PathName & "\."
        End If
    End If
    'Built fully qualified path
    Path = ""
    Do While FindFirstFileW(StrPtr(PathName), fileparent) > 0
        Path = GetStrFromUnicode(fileparent.cFileName) & "\" & Path
        PathName = PathName & "\.."
    Loop
    'If path is still empty, PathName is either "C:\." or inexistant. So we check the existence:
    On Error GoTo ErrWrongPath
    If Path = "" Then GetAttr PathName
    'Append drive letter
    Path = left$(PathName, 3) & Path
    '# needs further work to take into consideration UNC paths) #'
ErrWrongPath:
End Sub
'---------------------------------------------------------------------------------------------------------------------'
'FileTimeToDate Function
'   <Description>
'       Converts a File Time into a Date.
'   <Syntax>
'       FileTimeToDate(*FileTime_*)
'   <Parameters>
'       *FileTime_* Required. A pointer to a FILETIME structure containing  the file time to convert to a Date.
'---------------------------------------------------------------------------------------------------------------------'
Private Function FileTimeToDate(ByRef FileTime_ As FILETIME) As Date
    Dim LocalFileTime_ As FILETIME
    Dim SystemTime_ As SYSTEMTIME
    FileTimeToLocalFileTime FileTime_, LocalFileTime_
    FileTimeToSystemTime LocalFileTime_, SystemTime_
    With SystemTime_
        If .wMilliseconds >= 500 Then .wSecond = .wSecond + 1
        FileTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                               BUILDER AND DESTROYER                                               +'
'+-------------------------------------------------------------------------------------------------------------------+'
Private Sub Class_Initialize()
    ReDim Search(MAXDEPTH)
End Sub
Private Sub Class_Terminate()
    Dim i As Long
    For i = 1 To CurSearch
        If Search(i).Handle Then FindClose Search(i).Handle
    Next
    CurSearch = 0
End Sub
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                    PROPERTIES                                                     +'
'+-------------------------------------------------------------------------------------------------------------------+'
'Name, FullPath, Extension Properties
'   <Description>
'       .......
'---------------------------------------------------------------------------------------------------------------------'
Property Get Name() As String
    If CurSearch Then Name = GetStrFromUnicode(fileinfo.cFileName)
End Property
Property Get RelPath() As String
    If CurSearch Then RelPath = GetStrFromUnicode(fileinfo.cFileName)
End Property
Property Get FullPath() As String
    If CurSearch Then FullPath = Search(CurSearch).Path & Name
End Property
Property Get Extension() As String
    Dim i As Long, j As Long, Str As String
    Str = fileinfo.cFileName
    If CurSearch And Not CBool(fileinfo.dwFileAttributes And dlDirectory) Then
        i = InStr(Str, vbNullChar & vbNullChar)
        j = InStrRev(Str, "." & vbNullChar, i)
        If j Then Extension = StrConv(Mid$(Str, j + 2, i - j - 1), vbFromUnicode)
    End If
End Property
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
Property Get Attributes() As Long
    If CurSearch Then Attributes = fileinfo.dwFileAttributes
End Property
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
Property Get DateCreated() As Date
    If CurSearch Then DateCreated = FileTimeToDate(fileinfo.ftCreationTime)
End Property
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
Property Get DateLastModified() As Date
    If CurSearch Then DateLastModified = FileTimeToDate(fileinfo.ftLastWriteTime)
End Property
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
Property Get DateLastAccessed() As Date
    If CurSearch Then DateLastAccessed = FileTimeToDate(fileinfo.ftLastAccessTime)
End Property
'---------------------------------------------------------------------------------------------------------------------'
'---------------------------------------------------------------------------------------------------------------------'
Property Get FileSize() As Currency
    Dim FileSize2(2) As Long
    'File bigger than 4 GB!!!
    If fileinfo.nFileSizeHigh Then
        'Get size as two longs
        FileSize2(1) = fileinfo.nFileSizeLow
        FileSize2(2) = fileinfo.nFileSizeHigh
        'Copy to a Currency
        RtlMoveMemory FileSize, FileSize2(1), 8
    Else
        'File smaller than 4 GB
        FileSize = fileinfo.nFileSizeLow
    End If
End Property
Property Get EOF() As Boolean
    EOF = CurSearch = 0
End Property
'+-------------------------------------------------------------------------------------------------------------------+'
'+                                                 SEARCH FUNCTIONS                                                  +'
'+-------------------------------------------------------------------------------------------------------------------+'
'MoveNext Method
'   <Description>
'       Moves to next file in search directory that matches attributes specifications.
'   <Syntax>
'       MoveNext
'   <Parameters>
'       None.
'---------------------------------------------------------------------------------------------------------------------'
Private Sub MoveNext()
    Dim Handle As Long, Found As Boolean, Flag As Boolean
    Do
        Handle = Search(CurSearch).Handle
        Do
            'Move to next found item...
            If Handle Then  '... continuing current search
                Found = FindNextFileW(Handle, fileinfo)
            Else            '... or starting a new search
                Handle = FindFirstFileW(StrPtr(Search(CurSearch).Path & SearchMask), fileinfo)
                Search(CurSearch).Handle = Handle
                Found = Handle > 0
            End If
        'Loops until the found item matches attributes specifications or there is no more items
        Loop Until (CBool(fileinfo.dwFileAttributes And (SearchAttributes Or SearchRecursive)) _
        And (left$(fileinfo.cFileName, SELF_DIRECTORY_LEN) <> SELF_DIRECTORY) _
        And (left$(fileinfo.cFileName, PARENT_DIRECTORY_LEN) <> PARENT_DIRECTORY)) _
        Or Not Found
        'If no more items, resume parent search
        If Not Found Then
            FindClose Search(CurSearch).Handle
            CurSearch = CurSearch - 1
            Flag = CurSearch = 0
        'If found item is a folder and recursive search is on, create a new search
        ElseIf (fileinfo.dwFileAttributes And SearchRecursive) Then
            CurSearch = CurSearch + 1
            Search(CurSearch).Handle = 0
            Search(CurSearch).Path = Search(CurSearch - 1).Path & Name & "\"
            Flag = fileinfo.dwFileAttributes And SearchAttributes
        Else
            Flag = True
        End If
    'Loops until the found item matches attributes specifications or there is no more search in progress
    Loop Until Flag
End Sub
'---------------------------------------------------------------------------------------------------------------------'
'.......
'   <Description>
'       .......
'   <Syntax>
'       Dir [*pathname*[, *attributes*]]
'   <Parameters>
'       .......
'---------------------------------------------------------------------------------------------------------------------'
Function ModDir(Optional PathName, Optional Attributes As dlFileAttributes) As Boolean
    'If PathName is missing, continue current search
    If IsMissing(PathName) Then
        'You can't change attributes specification during a search
        If Attributes Then ErrRaise 513
        'Calling Dir after the last item returns False, calling Dir again raises an error
        If CurSearch = 0 Then ErrRaise 514
        'Move to next item
        MoveNext
    'Start new search
    Else
        'Stop current searches
        Class_Terminate
        'Save search parameters
        ReDim Preserve Search(1)
        SplitPathName PathName, Search(1).Path, SearchMask
        If Search(1).Path = "" Then
            CurSearch = 0
        Else
            If Attributes = 0 Then SearchAttributes = &HFFFFFF Else SearchAttributes = Attributes
            If SearchAttributes And dlRecursive Then SearchRecursive = dlDirectory Else SearchRecursive = 0
            CurSearch = 1
            Search(1).Handle = 0
            'Move to next item
            MoveNext
        End If
    End If
    ModDir = CurSearch
End Function

