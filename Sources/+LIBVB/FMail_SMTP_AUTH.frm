VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FMail_SMTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Sender"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1755
   Icon            =   "FMail_SMTP_AUTH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   1755
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   600
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FMail_SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nomsrc As String
Private g_adrsrc As String
Private g_nomdest As String
Private g_adrdest As String
Private g_subject As String
Private g_message As Variant
Private g_cr As Integer

Private g_smtp_user As String
Private g_smtp_pass As String

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_AUTH
    MAIL_USER
    MAIL_PASS
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State
Private m_strEncodedFiles As String


Public Function EnvoiMessage(ByVal v_nomsrc As String, _
                             ByVal v_adrsrc As String, _
                             ByVal v_nomdest As String, _
                             ByVal v_adrdest As String, _
                             ByVal v_subject As String, _
                             ByVal v_message As Variant, _
                             ByVal v_filename As String, _
                             Optional v_smtp_user As String = "", _
                             Optional v_smtp_pass As String = "") As Integer
    
    Dim ColonPos As Integer, lngPort As Long
    
    g_nomsrc = v_nomsrc
    g_adrsrc = v_adrsrc
    g_nomdest = v_nomdest
    g_adrdest = v_adrdest
    g_subject = v_subject
    g_message = v_message

    If v_filename <> "" Then
        m_strEncodedFiles = UUEncodeFile(v_filename) & vbCrLf
    End If
    
    g_smtp_user = v_smtp_user
    g_smtp_pass = v_smtp_pass
    
    'find out if the sender is using a Proxy server
    ColonPos = InStr(p_smtp_adrsrv, ":")
    If ColonPos = 0 Then
        'no proxy so use standard SMTP port
        Winsock.Connect p_smtp_adrsrv, 25
    Else
        'Proxy, so get proxy port number and parse out the server name or IP address
        lngPort = CLng(Right$(p_smtp_adrsrv, Len(p_smtp_adrsrv) - ColonPos))
        p_smtp_adrsrv = left$(p_smtp_adrsrv, ColonPos - 1)
        Winsock.Connect p_smtp_adrsrv, lngPort
    End If
    m_State = MAIL_CONNECT
    
    g_cr = 0
    While g_cr = 0
        DoEvents
        SYS_Sleep (1)
    Wend
    
    EnvoiMessage = g_cr
    
End Function

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse As String, strResponseCode As String
    Dim strDataToSend As String, strMessage As String
    Dim pos As Integer
    
    'Retrive data from winsock buffer
    Winsock.GetData strServerResponse

    Debug.Print strServerResponse
    
    'Get server response code (first three symbols)
    strResponseCode = left(strServerResponse, 3)
    
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                'Retrieve mailbox name from e-mail address
                pos = InStr(1, g_adrdest, "@")
                If pos = 0 Then
                    Winsock.Close
                    Call MsgBox("Adresse mail incorrecte : " & g_adrdest, _
                            vbInformation, "")
                    Call quitter(False)
                    Exit Sub
                End If
                strDataToSend = left$(g_adrdest, pos - 1)
                'Send HELO command to the server
                Winsock.SendData "HELO " & strDataToSend & vbCrLf
'                Debug.Print "HELO " & strDataToSend
            Case MAIL_HELO
                If g_smtp_user <> "" Then
                    m_State = MAIL_AUTH
                    Winsock.SendData "AUTH LOGIN" & vbCrLf
                Else
                    m_State = MAIL_FROM
                    Winsock.SendData "MAIL FROM:<" & g_adrsrc & ">" & vbCrLf
                End If
            Case MAIL_FROM
                'Change current state of the session
                m_State = MAIL_RCPTTO
                'Send RCPT TO command to the server
                Winsock.SendData "RCPT TO:<" & g_adrdest & ">" & vbCrLf
'                Debug.Print "RCPT TO:" & g_adrdest
            Case MAIL_RCPTTO
                'Change current state of the session
                m_State = MAIL_DATA
                'Send DATA command to the server
                Winsock.SendData "DATA" & vbCrLf
'                Debug.Print "DATA"
            Case MAIL_DATA
                'Change current state of the session
                m_State = MAIL_DOT
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf - This is wrong, it should be vbCrLf
                'see   http://cr.yp.to/docs/smtplf.html       for details
                
                'Send Subject line
                Winsock.SendData "From: " & g_nomsrc & " <" & g_adrsrc & ">" & vbCrLf
                Winsock.SendData "To: " & g_nomdest & " <" & g_adrdest & ">" & vbCrLf
                
                Winsock.SendData "Subject: " & g_subject & vbCrLf
'                Debug.Print "Subject: " & g_subject
                Winsock.SendData "Date: " & DATE_ToRFC822(Date + Time) & vbCrLf
                Winsock.SendData "MIME-Version: 1.0" & vbCrLf
                Winsock.SendData "Content-Type: text/plain" & vbCrLf
                If g_adrsrc <> "" Then
                    Winsock.SendData "Reply-To: " & g_nomsrc & " <" & g_adrsrc & ">" & vbCrLf & vbCrLf
                End If
                
                'Add attachments
                strMessage = g_message & vbCrLf & vbCrLf & m_strEncodedFiles
                'clear memory
                m_strEncodedFiles = ""
                'Debug.Print Len(strMessage)
                'These lines aren't needed, see
                '
                'http://cr.yp.to/docs/smtplf.html for details
                '
                '*****************************************
                'Parse message to get lines (for VB6 only)
                'varLines() = Split(strMessage, vbNewLine)
                'Parse message to get lines (for VB5 and lower)
                'SplitMessage strMessage, varLines()
                'clear memory
                'strMessage = ""
                '
                'Send each line of the message
                'For i = LBound(varLines()) To UBound(varLines())
                '    Winsock.SendData varLines(i) & vbCrLf
                '    '
                '    Debug.Print varLines(i)
                'Next
                '
                '******************************************
                Winsock.SendData strMessage & vbCrLf
                strMessage = ""
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock.SendData "." & vbCrLf
'                Debug.Print "."
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                'Send QUIT command to the server
                Winsock.SendData "QUIT" & vbCrLf
'                Debug.Print "QUIT"
            Case MAIL_QUIT
                'Close connection
                Winsock.Close
                Call quitter(True)
        End Select
    ElseIf strResponseCode = "334" Then
        Select Case m_State
            Case MAIL_AUTH
                m_State = MAIL_USER
                Winsock.SendData Base64Encode(g_smtp_user) & vbCrLf
                Debug.Print Base64Encode(g_smtp_user) & vbCrLf
            Case MAIL_USER
                m_State = MAIL_PASS
                Winsock.SendData Base64Encode(g_smtp_pass) & vbCrLf
                Debug.Print Base64Encode(g_smtp_pass) & vbCrLf
'STR_Decrypter (rs("UAPP_MotPasse").Value)
        
        End Select
    ElseIf strResponseCode = "235" Then
        Select Case m_State
            Case MAIL_PASS
                m_State = MAIL_FROM
                'Send MAIL FROM command to the server
                Winsock.SendData "MAIL FROM:<" & g_adrsrc & ">" & vbCrLf
        End Select
    Else
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        Winsock.Close
        If Not m_State = MAIL_QUIT Then
            Call MsgBox("Erreur SMTP: " & strServerResponse, _
                    vbInformation, "")
            Call quitter(False)
        Else
            ' Call MsgBox("Message sent successfuly.", vbInformation, "")
            Call quitter(True)
        End If
    End If
    
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Call MsgBox("Erreur Winsock " & Number & vbCrLf & _
            Description, vbExclamation, "")
    Winsock.Close
    Call quitter(False)
    
End Sub

Private Sub quitter(ByVal v_fok As Boolean)

    g_cr = IIf(v_fok, 1, -1)
    
End Sub

Private Sub SplitMessage(strMessage As String, strlines() As String)
    
    Dim intAccs As Long
    Dim i
    Dim lngSpacePos As Long, lngStart As Long
    
    strMessage = Trim$(strMessage)
    lngSpacePos = 1
    lngSpacePos = InStr(lngSpacePos, strMessage, vbNewLine)
    Do While lngSpacePos
        intAccs = intAccs + 1
        lngSpacePos = InStr(lngSpacePos + 1, strMessage, vbNewLine)
    Loop
    ReDim strlines(intAccs)
    lngStart = 1
    For i = 0 To intAccs
        lngSpacePos = InStr(lngStart, strMessage, vbNewLine)
        If lngSpacePos Then
            strlines(i) = Mid(strMessage, lngStart, lngSpacePos - lngStart)
            lngStart = lngSpacePos + Len(vbNewLine)
        Else
            strlines(i) = Right(strMessage, Len(strMessage) - lngStart + 1)
        End If
    Next

End Sub

Private Function UUEncodeFile(strFilePath As String) As String

    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFilename     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    Dim strResult       As String
    '
    'Get file name
    strFilename = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    '
    'Insert first marker: "begin 664 ..."
    strResult = "begin 664 " + strFilename + vbCrLf
    '
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize \ 45 + 1
    '
    'Prepare buffer to retrieve data from
    'the file by 45 symbols chunks
    strFileData = Space(45)
    '
    intFile = FreeFile
    '
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            'Read file data by 45-bytes cnunks
            '
            If i = lEncodedLines Then
                'Last line of encoded data often is not
                'equal to 45, therefore we need to change
                'size of the buffer
                strFileData = Space(lFileSize Mod 45)
            End If
            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            '
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                'If the last line is processed and length of
                'source data is not a number divisible by 3, add one or two
                'blankspace symbols
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            strResult = strResult + strTempLine + vbCrLf
            'reset line buffer
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    strResult = strResult & "`" & vbCrLf + "end" + vbCrLf
    'asign return value
    UUEncodeFile = strResult
    
End Function

Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, i
  
  'For each group of 3 bytes
  For i = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, i, 1)) + _
      &H100 * MyASC(Mid(inData, i + 1, 1)) + MyASC(Mid(inData, i + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function




