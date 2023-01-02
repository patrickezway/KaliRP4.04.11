VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tableau de Bord"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1695
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Logo 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox pct2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   4080
      Picture         =   "Splash.frx":0588
      ScaleHeight     =   1485
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox pct1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1080
      Picture         =   "Splash.frx":928E
      ScaleHeight     =   975
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Timer tmr 
      Left            =   0
      Top             =   2640
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "V ???"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_lsec As Long
Private g_ldelai As Long

Private Sub Form_Load()
    
    ' Initialisation de l'appli
    pct1.Visible = True
    pct2.Visible = True
    lblVersion.Visible = True
                                    
    ' Initialisation du timer
    tmr.Enabled = False
    tmr.Interval = 1000
    
    ' Initialisation du compteur
    g_lsec = 0
    
    ' Form au 1er plan
    Call FRM_AuPremierPlan(Me.hwnd)
    
End Sub

Public Sub CloseAfter(ByVal v_lattente As Long)
    
    ' en haut à gauche
    Me.Top = 10
    Me.left = 10
    tmr.Enabled = True
    g_ldelai = v_lattente
    
End Sub

Private Sub tmr_Timer()
    
    g_lsec = g_lsec + 1
   ' Me.Print sec
    If g_lsec >= g_ldelai Then Unload Me
    
End Sub
