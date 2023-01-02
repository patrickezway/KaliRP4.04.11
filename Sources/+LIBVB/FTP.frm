VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FTP 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet inet 
      Left            =   480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://anonymous@"
      UserName        =   "anonymous"
      Password        =   "fred"
   End
End
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function PutFile(ByVal v_nomsrc As String, _
                        ByVal v_dirdest As String) As Integer

    Dim nom_simple As String
    
    nom_simple = Mid$(v_nomsrc, InStrRev(v_nomsrc, "\") + 1)
    
    Call ftp_connect
    
    inet.Execute , "cd """ & v_dirdest & """"
    Do Until inet.StillExecuting = 0
        DoEvents
    Loop
    
    inet.Execute , "put """ & v_nomsrc & """ " & nom_simple
    Do Until inet.StillExecuting = 0
        DoEvents
    Loop

    ftp_quit
    
End Function

Public Function GetFile(ByVal v_nomsrc As String, _
                        ByVal v_dirdest As String) As Integer

    Dim nom_simple As String
    
    ChDir (v_dirdest)
    
    Call ftp_connect
    
    inet.Execute , "get """ & v_nomsrc & """ " & "mon_essai"
    Do Until inet.StillExecuting = 0
        DoEvents
    Loop

    ftp_quit
    
End Function

Private Sub ftp_connect()

    inet.RemoteHost = "192.168.101.238"
    inet.UserName = "kalidoc"
    inet.Password = "kalidoc"

End Sub

Private Sub ftp_quit()

    inet.Execute , "quit"

End Sub

Private Sub inet_StateChanged(ByVal State As Integer)

    Dim s As String
    
    Select Case State
    Case 0
       s = "No state information is available."
    Case 1
       s = "Looking up the IP address for the remote server."
    Case 2
       s = "Found the IP address for the remote server."
    Case 3
       s = "Connecting to the remote server."
    Case 4
       s = "Connected to the remote server."
    Case 5
       s = "Requesting information from the remote server."
    Case 6
       s = "The request was sent successfully to the remote server."
    Case 7
       s = "Receiving a response from the remote server."
    Case 8
       s = "The response was received successfully from the remote server."
    Case 9
       s = "Disconnecting from the remote server."
    Case 10
       s = "Disconnected from the remote server."
    Case 11
       s = "An error has occurred while communicating with the remote server."
    Case 12
       s = "The request complete, files have been received."
    Case Else
       s = "Unknown state: " & FormatNumber(s, 0)
    End Select
End Sub
