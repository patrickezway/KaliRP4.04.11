VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Com_ChoixFichier2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   15
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Com_ChoixFichier2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nomfich As String
Private g_pattern As String

Public Function AppelFrm(ByVal v_titre As String, _
                         ByVal v_drive As String, _
                         ByVal v_path As String, _
                         ByVal v_pattern As String, _
                         ByVal v_selrep As Boolean) As String
                         
    With cdlg
        .InitDir = v_path
        .DialogTitle = v_titre
        .CancelError = False
        .FLAGS = cdlOFNExplorer Or cdlOFNHideReadOnly 'or cdlOFNFileMustExist
        .Filter = "(" & v_pattern & ")|" & v_pattern
    End With
    
    g_pattern = v_pattern
    
    Me.Show 1
    
    AppelFrm = g_nomfich
    
End Function

Private Function verif_ext(ByVal v_nomfich As String) As Boolean

    Dim sext As String
    Dim pos As Integer
    
    pos = InStrRev(v_nomfich, ".")
    If pos = 0 Then
        verif_ext = False
        Exit Function
    End If
    sext = LCase(Mid$(v_nomfich, pos))
    If sext = ".xls" Or sext = ".xlsx" Then
        verif_ext = True
    Else
        verif_ext = False
    End If
    
End Function

Private Sub Form_Activate()

    Call FRM_ResizeForm(Me, 0, 0)
    
    With cdlg
        On Error GoTo lab_fin
lab_show:
        .ShowOpen
        On Error GoTo 0
        If Len(.FileName) = 0 Then
            g_nomfich = ""
            Unload Me
            Exit Sub
        End If
        If Not verif_ext(.FileName) Then
            Call MsgBox("Cette extension n'est pas autorisée.", vbInformation + vbOKOnly, "")
            .FileName = ""
            GoTo lab_show
        End If
        g_nomfich = .FileName
    End With
    
lab_fin:
    Unload Me
    
End Sub

