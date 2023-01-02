VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Sender"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Reply To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   2640
      Width           =   4725
      Begin VB.TextBox txtReplyToName 
         Height          =   285
         Left            =   1635
         TabIndex        =   4
         Text            =   "txtReplyToName"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtReplyTo 
         Height          =   285
         Left            =   1635
         TabIndex        =   5
         Text            =   "txtReplyTo"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "E-mail address:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1305
      End
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1875
      TabIndex        =   23
      Text            =   "txtHost"
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1875
      TabIndex        =   6
      Text            =   "txtSubject"
      Top             =   3795
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "Recipient:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   4725
      Begin VB.TextBox txtRecipient 
         Height          =   285
         Left            =   1635
         TabIndex        =   3
         Text            =   "txtRecipient"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtRecipientName 
         Height          =   285
         Left            =   1635
         TabIndex        =   2
         Text            =   "txtRecipientName"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "E-mail address:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sender:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   4725
      Begin VB.TextBox txtSenderName 
         Height          =   285
         Left            =   1635
         TabIndex        =   0
         Text            =   "txtSenderName"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtSender 
         Height          =   285
         Left            =   1635
         TabIndex        =   1
         Text            =   "txtSender"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "E-mail address:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attach files"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   6720
      Width           =   6375
      Begin VB.ListBox lstAttachments 
         Height          =   840
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send message"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Message"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtMessage 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmSendMail.frx":030A
      Top             =   4200
      Width           =   6375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   330
      TabIndex        =   15
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SMTP Host:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State
Private m_strEncodedFiles As String
'

Private Sub cmdAddFile_Click()
Dim varFilePath As Variant
Dim strMyDocs As String
strMyDocs = GetSpecialFolderLocation(Me.hWnd)
varFilePath = CommFileDialog(strMyDocs, , , , , "File to attach", Me.hWnd)
If Not IsNull(varFilePath) Then lstAttachments.AddItem CStr(varFilePath)
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdNew_Click()

    txtRecipient = ""
    txtSubject = ""
    txtMessage = ""
    
End Sub

Private Sub cmdRemove_Click()

    On Error Resume Next
    
    lstAttachments.RemoveItem lstAttachments.ListIndex

End Sub

Private Sub cmdSend_Click()
    '
    Dim i As Integer
    Dim strServer As String, ColonPos As Integer, lngPort As Long
    '
    'prepare attachments
    '
    For i = 0 To lstAttachments.ListCount - 1
        'lstAttachments.ListIndex = i
        m_strEncodedFiles = m_strEncodedFiles & _
                         UUEncodeFile(lstAttachments.List(i)) & vbCrLf
    Next i
    
    strServer = Trim(txtHost)
    'find out if the sender is using a Proxy server
    ColonPos = InStr(strServer, ":")
    If ColonPos = 0 Then
        'no proxy so use standard SMTP port
        Winsock1.Connect strServer, 25
    Else
        'Proxy, so get proxy port number and parse out the server name or IP address
        lngPort = CLng(Right$(strServer, Len(strServer) - ColonPos))
        strServer = Left$(strServer, ColonPos - 1)
        Winsock1.Connect strServer, lngPort
    End If
    m_State = MAIL_CONNECT
    '
End Sub

Private Sub Form_Load()
    '
    'clear all textboxes
    '
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_colAttachments = Nothing
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:<" & Trim$(txtSender) & ">" & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(txtSender)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:<" & Trim$(txtRecipient) & ">" & vbCrLf
                '
                Debug.Print "RCPT TO:" & Trim$(txtRecipient)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf - This is wrong, it should be vbCrLf
                'see   http://cr.yp.to/docs/smtplf.html       for details
                '
                'Send Subject line
                Winsock1.SendData "From:" & txtSenderName & " <" & txtSender & ">" & vbCrLf
                Winsock1.SendData "To:" & txtRecipientName & " <" & txtRecipient & ">" & vbCrLf
                
                '
                Debug.Print "Subject: " & txtSubject
                '
                If Len(txtReplyTo.Text) > 0 Then
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf
                    Winsock1.SendData "Reply-To:" & txtReplyToName & " <" & txtReplyTo & ">" & vbCrLf & vbCrLf
                Else
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf & vbCrLf
                End If
                'Dim varLines() As String
                'Dim varLine As String
                Dim strMessage As String
                'Dim i
                '
                'Add atacchments
                strMessage = txtMessage & vbCrLf & vbCrLf & m_strEncodedFiles
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
                '    Winsock1.SendData varLines(i) & vbCrLf
                '    '
                '    Debug.Print varLines(i)
                'Next
                '
                '******************************************
                Winsock1.SendData strMessage & vbCrLf
                strMessage = ""
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.Close
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.Close
        '
        If Not m_State = MAIL_QUIT Then
            MsgBox "SMTP Error: " & strServerResponse, _
                    vbInformation, "SMTP Error"
        Else
            MsgBox "Message sent successfuly.", vbInformation
        End If
        '
    End If
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error number " & Number & vbCrLf & _
            Description, vbExclamation, "Winsock Error"

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
