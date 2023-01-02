VERSION 5.00
Begin VB.Form PrmFormAction 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   16725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   16725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Critères d'extraction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   10065
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16845
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   9375
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   16725
         Begin VB.CommandButton cmdAussi 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   11160
            Picture         =   "PrmFormAction.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Etendre cette condition à d'autres filtres"
            Top             =   960
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.CommandButton cmdBoucle 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Poser toutes les Questions en boucle"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Choisir un opérateur"
            Top             =   240
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   3675
         End
         Begin VB.TextBox txtCnd 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox txtChp 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   930
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtOper 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Choisir un opérateur"
            Top             =   930
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.CommandButton cmdOper 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   5160
            Picture         =   "PrmFormAction.frx":0457
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Choisir un opérateur"
            Top             =   960
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.TextBox txtVal 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   5640
            TabIndex        =   5
            ToolTipText     =   "Choisir une valeur"
            Top             =   960
            Visible         =   0   'False
            Width           =   5085
         End
         Begin VB.CommandButton cmdVal 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   10800
            Picture         =   "PrmFormAction.frx":069C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Choisir une valeur"
            Top             =   960
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   320
         End
         Begin VB.Label lblFF 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   11520
            TabIndex        =   16
            Top             =   960
            Width           =   5055
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Champ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Opérateur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   10
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valeur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   5640
            TabIndex        =   9
            Top             =   720
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   800
      Left            =   0
      TabIndex        =   0
      Top             =   10080
      Width           =   16875
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   480
         Picture         =   "PrmFormAction.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   15960
         Picture         =   "PrmFormAction.frx":0D3A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Quitter sans tenir compte des modifications"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmFormAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1

Private g_FF_Num As Long
Private g_FF_Indice As Long
Private g_sQ As String
Private g_Trait As String
Private g_FF_oper As String
Private g_initFait As Boolean

Private g_txt_avant As String
Private g_mode_saisie As Boolean
Private g_form_active As Boolean
Private IndLigneCourrante As Integer
Private ChpCourrant As String
Private NbLigTotal As Integer

Private Bool_Faire_TvVal_Gotfocus As Boolean

Public Function AppelFrm(ByVal v_FF_Num As Long, _
                         ByVal v_FF_Indice As Long, _
                         ByVal v_FF_oper As String, _
                         ByVal v_sQ As String, _
                         ByVal v_Trait As String)
                    
    g_FF_Num = v_FF_Num
    g_FF_Indice = v_FF_Indice
    g_FF_oper = v_FF_oper
    g_sQ = v_sQ
    g_Trait = v_Trait
    
    On Error Resume Next
    Me.Show 1
    
   
End Function


Private Function choisir_fonctions(ByVal v_sfct As String) As String

    Dim sql As String, nomFct As String, s As String
    Dim trouve As Boolean
    Dim n As Integer, i As Integer, lig As Integer, btn_sortie As Integer
    Dim numfct As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    n = STR_GetNbchamp(v_sfct, ";")
    For i = 0 To n - 1
        numfct = Mid$(STR_GetChamp(v_sfct, ";", i), 2)
        Call Odbc_RecupVal("select ft_libelle from fcttrav where ft_num=" & numfct, nomFct)
        Call CL_AddLigne(nomFct, numfct, "", True)
    Next i
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_fonctions = ""
        Exit Function
    End If
    While Not rs.EOF
        trouve = False
        For i = 0 To n - 1
            If Mid$(STR_GetChamp(v_sfct, ";", i), 2) = rs("FT_Num").Value Then
                trouve = True
                Exit For
            End If
        Next i
        If Not trouve Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Choix des fonctions", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
    Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    
    ' Sortie
    If CL_liste.retour = 1 Then
        choisir_fonctions = ""
        Exit Function
    End If
    
    s = ""
    For i = 0 To UBound(CL_liste.lignes())
        If CL_liste.lignes(i).selected Then
            s = s & "F" & CL_liste.lignes(i).num & ";"
        End If
    Next i
    
    choisir_fonctions = s
    
End Function


Private Sub choisir_oper(ByVal v_indice As Integer, v_Trait As String)
    Dim IndTbl As Integer
    Dim s As String
    
    If v_indice > 0 Then
    
        Call CL_Init
    
        ' quel type ?
        IndTbl = Me.txtChp(v_indice).tag
        'MsgBox tbl_Demande(IndTbl).DemandType
        If InStr(tbl_Demande(IndTbl).DemandFctValid, "NUMSERVICE") > 0 Then
            s = "="
            GoTo LabChoisirValeur
        ElseIf tbl_Demande(IndTbl).DemandType = "TEXT" Then
            Call CL_AddLigne("Egal à", 0, "=", False)
            Call CL_AddLigne("Différent de", 0, "!", False)
            Call CL_AddLigne("Supérieur ou égal", 0, ">=", False)
            Call CL_AddLigne("Inférieur ou égal", 0, "<=", False)
            If InStr(tbl_Demande(IndTbl).DemandFctValid, "%DATE") > 0 Then
                Call CL_AddLigne("compris entre", 0, "COMPRIS", False)
            End If
        ElseIf tbl_Demande(IndTbl).DemandType = "SELECT" Or tbl_Demande(IndTbl).DemandType = "RADIO" Or tbl_Demande(IndTbl).DemandType = "CHECK" Then
            Call CL_AddLigne("Egal", 0, "=", False)
            Call CL_AddLigne("Différent", 0, "!", False)
        ElseIf tbl_Demande(IndTbl).DemandType = "HIERARCHIE" Then
            'Call CL_AddLigne("Egal", 0, "=", False)
            Call CL_AddLigne("Parmis", 0, "=", False)
        End If
        Call CL_InitTitreHelp(tbl_Demande(IndTbl).DemandChpStr & " " & tbl_Demande(IndTbl).DemandChpStrPlus, "")
        Call CL_InitTaille(0, -5)
        Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        ChoixListe.Show 1
        ' Quitter
        If CL_liste.retour = 1 Then
            Exit Sub
        End If
    
        txtOper(v_indice).tag = CL_liste.lignes(CL_liste.pointeur).tag
        txtOper(v_indice).Text = CL_liste.lignes(CL_liste.pointeur).texte
                
        cmd(CMD_OK).Enabled = True
        
        If v_Trait <> "Boucle" And Me.txtVal(v_indice).tag = "" Then
            s = CL_liste.lignes(CL_liste.pointeur).tag
LabChoisirValeur:
            Call choisir_valeur(v_indice, "Direct", s)
        End If
        ' calcul de la condition
        FctCalculCondition (v_indice)
        'Passer au suivant
        If v_Trait = "Boucle" Then
            IndLigneCourrante = v_indice
            ChpCourrant = "OPER"
            FctPasserSuivant IndLigneCourrante, ChpCourrant, v_Trait
        End If
    End If
End Sub

Private Sub FctCalculCondition(ByVal v_indice As Integer)
    Dim IndTbl  As Integer
    Dim laS As String, laSPF As String
    
    If v_indice > 0 Then
        ' quel type ?
        IndTbl = Me.txtChp(v_indice).tag
        ' MsgBox "FctCalculCondition " & tbl_Demande(IndTbl).DemandenSQL
        laS = "CHP:" & tbl_Demande(IndTbl).DemandChpNum & ":" & FctNomChp(tbl_Demande(IndTbl).DemandChpNum) & "¤"
        laS = laS & "OP:" & Me.txtOper(v_indice).tag & "¤"
        If InStr(tbl_Demande(IndTbl).DemandFctValid, "%DATE") > 0 Then
            laS = laS & "DATE:" & Me.txtVal(v_indice).tag
        ElseIf InStr(tbl_Demande(IndTbl).DemandFctValid, "%NUMSERVICE") > 0 Then
            laS = laS & "NUMSERVICE:" & Me.txtVal(v_indice).tag
        ElseIf InStr(tbl_Demande(IndTbl).DemandFctValid, "%NUMFCT") > 0 Then
            laS = laS & "NUMFCT:" & Me.txtVal(v_indice).tag
        ElseIf InStr("SELECT%RADIO%CHECK", tbl_Demande(IndTbl).DemandType) > 0 Then
            laS = laS & tbl_Demande(IndTbl).DemandType & ":" & Me.txtVal(v_indice).tag
        ElseIf InStr("HIERARCHIE", tbl_Demande(IndTbl).DemandType) > 0 Then
            laS = laS & tbl_Demande(IndTbl).DemandType & ":" & Me.txtVal(v_indice).tag
        Else
            Me.txtVal(v_indice).tag = Me.txtVal(v_indice).Text
            laS = laS & "VAL:" & Me.txtVal(v_indice).tag
        End If
        If Me.txtOper(v_indice).tag <> "" And Me.txtVal(v_indice).tag <> "" Then
            tbl_Demande(IndTbl).DemandFait = True
        Else
            tbl_Demande(IndTbl).DemandFait = False
        End If
        tbl_Demande(IndTbl).DemandPasFrancais = laS
        Me.txtCnd(v_indice).Text = laS
    End If
End Sub

Private Function FctNomChp(ByVal V_ChpNum As String)
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from FormEtapeChp where FOREC_Num=" & V_ChpNum
    If Odbc_SelectV(sql, rs) <> P_ERREUR Then
        If Not rs.EOF Then
            FctNomChp = rs("Forec_Nom")
        End If
        rs.Close
    End If
End Function
Private Sub FctPasserSuivant(ByRef r_IndLigneCourrante, ByRef r_ChpCourrant, ByVal v_Trait As String)
    'MsgBox r_IndLigneCourrante & " " & r_ChpCourrant
    If r_ChpCourrant = "VAL" Then
        ' ligne suivante
        r_IndLigneCourrante = r_IndLigneCourrante + 1
        If r_IndLigneCourrante > NbLigTotal Then
            'MsgBox "Fini"
            Exit Sub
        End If
        choisir_oper r_IndLigneCourrante, v_Trait
    End If
    If r_ChpCourrant = "OPER" Then
        Call choisir_valeur(r_IndLigneCourrante, v_Trait, txtOper(r_IndLigneCourrante).tag)
    End If
End Sub

Private Function choisir_une_fonction() As String

    Dim sql As String
    Dim n As Integer
    Dim num As Long
    Dim rs As rdoResultset
    
    Call CL_Init
    
    sql = "select * from FctTrav" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_une_fonction = ""
        Exit Function
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        choisir_une_fonction = ""
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -25)
lab_choix:
    ChoixListe.Show 1
    ' Sortie
    If CL_liste.retour = 1 Then
        choisir_une_fonction = ""
        Exit Function
    End If
    
    num = CL_liste.lignes(CL_liste.pointeur).num
    choisir_une_fonction = CL_liste.lignes(CL_liste.pointeur).num & vbTab & CL_liste.lignes(CL_liste.pointeur).texte
    
End Function

Private Function choisir_Aussi(v_index As Integer) As String

    Dim sql As String, sin As String, ff_fornums As String, s As String
    Dim n As Integer, iff As Integer
    Dim num As Long
    Dim rs As rdoResultset
    Dim iD As Integer
    Dim i As Integer
    Dim iDem As Integer
    Dim new_iD As Integer
    Dim snew_ID As String
    Dim sclé As String
    Dim ya As Boolean
    Dim libFF As String
    Dim libF As String
    Dim j As Integer
    Dim LibType As String
    Dim bMettre As Boolean
    Dim strya As String
    Dim strya2 As String
    Dim chpactuel As String
    Dim strAussi As String
    Dim newAussi As String
    Dim àGarder As Boolean
    Dim clé As String
    
    Call CL_Init
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)

    newAussi = ""
    ' choisir les filtres possibles (sauf lui même)
    strya = cmdAussi(v_index).tag
    strAussi = strya
    For i = 0 To PiloteExcelBis.grdForm.Rows - 1
        iDem = Me.lblFF(v_index).tag
        bMettre = True
        If tbl_Demande(iDem).DemandFFNum = PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_Num) Then
            If tbl_Demande(iDem).DemandFormInd = PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_NumIndice) Then
                ' c'est lui même
                bMettre = False
            End If
        End If
        If bMettre Then
            clé = PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_Num) & ":" & PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_NumIndice)
            libFF = PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_Lib) & " " & PiloteExcelBis.grdForm.TextMatrix(i, GrdForm_FF_Titre)
            ' Quel sont les champs élligibles
            sclé = clé
            For j = 0 To STR_GetNbchamp(strya, ";")
                If Mid(STR_GetChamp(strya, ";", j), 1, Len(clé)) = clé Then
                    sclé = STR_GetChamp(strya, ";", j)
                    Exit For
                End If
            Next j
            If InStr(strya, clé & ":") = 0 Then
                Call CL_AddLigne(i & "->" & libFF, i, sclé, False)
            Else
                Call CL_AddLigne(i & "->" & libFF, i, sclé, True)
            End If
            n = n + 1
        End If
    Next i
    If n > 0 Then
        Call CL_InitTitreHelp("Appliquer aussi à ", "")
        Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        Call CL_InitTaille(0, -25)
        ChoixListe.Show 1
        ' Sortie
        If CL_liste.retour = 1 Then
            choisir_Aussi = ""
            Exit Function
        End If
        
        ' Pour chaque filtre, choisir le champ
        strya = ""
        For i = 0 To UBound(CL_liste.lignes())
            If CL_liste.lignes(i).selected Then
                strya = strya & CL_liste.lignes(i).num & ";"
                strya2 = strya2 & CL_liste.lignes(i).tag & ";"
            End If
        Next i
        For i = 0 To STR_GetNbchamp(strya, ";") - 1
            If STR_GetChamp(strya, ";", i) <> "" Then
                Call CL_Init
                Call CL_InitMultiSelect(True, False)
                Call CL_InitGererTousRien(False)
                j = STR_GetChamp(strya, ";", i)
                clé = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Num) & ":" & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_NumIndice)
                libFF = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Lib) & " " & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Titre)
                ' Quel sont les champs élligibles pour ce formulaire
                sql = "select ff_fornums from filtreform where ff_num=" & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Num)
                Call Odbc_RecupVal(sql, ff_fornums)
                n = STR_GetNbchamp(ff_fornums, "*")
                sin = "("
                For iff = 1 To n - 1
                    s = STR_GetChamp(ff_fornums, "*", iff)
                    If iff > 1 Then sin = sin + ","
                    sin = sin + s
                Next iff
                sin = sin + ")"
                sql = "select distinct(forec_num),* from formetapechp where forec_fornum in " & sin
                If InStr(tbl_Demande(iDem).DemandFctValid, "%DATE") > 0 Then
                    sql = sql & " and Forec_FctValid like '%DATE%'"
                    sql = sql & " and forec_type = '" & tbl_Demande(iDem).DemandType & "'"
                ElseIf InStr("SELECT*RADIO*CHECK", tbl_Demande(iDem).DemandType) > 0 Then
                    sql = sql & " and Forec_valeurs_possibles = " & tbl_Demande(iDem).DemandValeursPossibles
                Else
                    sql = sql & " and Forec_FctValid = '" & tbl_Demande(iDem).DemandFctValid & "'"
                    sql = sql & " and forec_type = '" & tbl_Demande(iDem).DemandType & "'"
                End If
                sql = sql & " order by forec_numetape,forec_ordre"
                
                Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
                If rs.EOF Then
                    Call MsgBox("pas de champs de type" & tbl_Demande(iDem).DemandFctValid & " pour le filtre " & Chr(13) & Chr(10) & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Lib) & " : " & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Titre))
                    GoTo LabNextI
                End If
                
                chpactuel = STR_GetChamp(strya2, ";", i)
                chpactuel = STR_GetChamp(chpactuel, ":", 2)
                While Not rs.EOF
                    àGarder = True
                    LibType = ""
                    If rs("FOREC_Type") = "TEXT" Then
                        If InStr(rs("forec_fctvalid"), "%DATE") > 0 Then
                            libF = rs("forec_label")
                            àGarder = True
                            LibType = "Date"
                        ElseIf rs("forec_fctvalid") = "%NUMSERVICE" Then
                            libF = rs("forec_label")
                            àGarder = True
                            LibType = "Service"
                        End If
                    ElseIf rs("FOREC_Type") = "RADIO" Or rs("FOREC_Type") = "CHECK" Or rs("FOREC_Type") = "SELECT" Then
                        libF = rs("forec_label")
                        àGarder = True
                        LibType = "Liste"
                    End If
                    If àGarder Then
                        libF = libF & " Etape " & rs("forec_numetape")
                        If InStr(strAussi, rs("forec_num") & ";") = 0 Then
                            Call CL_AddLigne(libF & "  (" & LibType & ")", i, clé & ":" & rs("forec_num"), False)
                        Else
                            Call CL_AddLigne(libF & "  (" & LibType & ")", i, clé & ":" & rs("forec_num"), True)
                        End If
                        n = n + 1
                    End If
                    rs.MoveNext
                Wend
            End If
            libFF = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Lib) & " " & PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Titre)
            Call CL_InitTitreHelp("Champ pour " & libFF, "")
            Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
            Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
            Call CL_InitTaille(0, -25)
            ChoixListe.Show 1
            ' Sortie
            If CL_liste.retour = 1 Then
                choisir_Aussi = ""
                Exit Function
            End If
            For j = 0 To UBound(CL_liste.lignes())
                If CL_liste.lignes(j).selected Then
                    If CL_liste.lignes(j).tag <> "" Then
                        newAussi = newAussi & CL_liste.lignes(j).tag & ";"
                        Exit For
                    End If
                End If
            Next j
LabNextI:
        Next i
    End If
    cmdAussi(v_index).tag = newAussi
    
    iD = Me.txtChp(v_index).tag ' id de tbl_demande
    If n = 0 Then
        choisir_Aussi = ""
        Exit Function
    End If
    tbl_Demande(iD).DemandAussiStr = newAussi
    tbl_Demande(iD).DemandAussiBool = IIf(tbl_Demande(iD).DemandAussiStr = "", False, True)
    If tbl_Demande(iD).DemandAussiBool Then
        cmdAussi(v_index).BackColor = cmdBoucle.BackColor
        cmdAussi(v_index).Width = 600
    Else
        cmdAussi(v_index).BackColor = cmdBoucle.BackColor
        cmdAussi(v_index).Width = 320
    End If
End Function

Private Sub choisir_valeur(ByVal v_indice As Integer, v_Trait As String, v_operateur As String)

    Dim IndTbl As Integer, stype As String
    Dim sql As String, rs As rdoResultset
    Dim sqlH As String, rsH As rdoResultset
    Dim sfct_valid As String
    Dim n As Integer, i As Integer, lig As Integer
    Dim II As String
    Dim frm As Form
    Dim sret As String, stag As String, stext As String, num_srv As String, nom_srv As String
    Dim num_fct As String, nom_fct As String, sval As String, numlst As String
    Dim nbDate As Integer
    Dim s As String
    Dim d1 As String, d2 As String
        
    'MsgBox v_indice
    nbDate = 1
    If v_indice > 0 Then
    
        ' quel type ?
        IndTbl = Me.txtChp(v_indice).tag
        stype = tbl_Demande(IndTbl).DemandType
        If stype = "TEXT" Then
            If InStr(tbl_Demande(IndTbl).DemandFctValid, "%DATE") > 0 Then
                stext = Date
encoreDate:
                If v_operateur = "COMPRIS" Then
                    If nbDate = 1 Then
                        s = " Date Du"
                        stext = IIf(STR_GetChamp(Me.txtVal(v_indice).Text, " ", 0) = "", Date, STR_GetChamp(Me.txtVal(v_indice).Text, " ", 0))
                    End If
                    If nbDate = 2 Then
                        s = " Date Au"
                        stext = IIf(STR_GetChamp(Me.txtVal(v_indice).Text, " ", 1) = "", Date, STR_GetChamp(Me.txtVal(v_indice).Text, " ", 1))
                    End If
                End If
                stext = InputBox("Valeur pour " & Chr(13) & Chr(10) & tbl_Demande(IndTbl).DemandChpStr & " " & tbl_Demande(IndTbl).DemandChpStrPlus & " " & s, "Valeur de type Texte", stext)
                If stext = "" Then Exit Sub
                If Not IsDate(stext) Then
                    MsgBox "Format de date invalide"
                    GoTo encoreDate
                End If
                If v_operateur = "COMPRIS" Then
                    If nbDate = 1 Then
                        d1 = stext
                        nbDate = 2
                        GoTo encoreDate
                    Else
                        d2 = stext
                    End If
                End If
            ElseIf tbl_Demande(IndTbl).DemandFctValid = "" Then
                stext = ""
                stext = InputBox("Valeur pour " & Chr(13) & Chr(10) & tbl_Demande(IndTbl).DemandChpStr & " " & tbl_Demande(IndTbl).DemandChpStrPlus, "Valeur de type Texte", stext)
            ElseIf InStr(tbl_Demande(IndTbl).DemandFctValid, "%NUMSERVICE") > 0 Then
                sfct_valid = "%NUMSERVICE"
                GoTo Lab_NumService
            ElseIf InStr(tbl_Demande(IndTbl).DemandFctValid, "%NUMFCT") > 0 Then
                sfct_valid = "%NUMFCT"
                GoTo Lab_NumFct
            ElseIf InStr(tbl_Demande(IndTbl).DemandFctValid, "%ENTIER") > 0 Or InStr(tbl_Demande(IndTbl).DemandFctValid, "%MONTANT") > 0 Or InStr(tbl_Demande(IndTbl).DemandFctValid, "%NUM") > 0 Then
                stext = ""
encoreNum:
                stext = InputBox("Valeur pour " & Chr(13) & Chr(10) & tbl_Demande(IndTbl).DemandChpStr & " " & tbl_Demande(IndTbl).DemandChpStrPlus, "Valeur de type Numérique", stext)
                If Not IsNumeric(stext) Then
                    MsgBox "Format invalide"
                    GoTo encoreNum
                End If
            Else
                MsgBox "Cas FCT"
                stext = InputBox("Valeur pour " & Chr(13) & Chr(10) & tbl_Demande(IndTbl).DemandChpStr & " " & tbl_Demande(IndTbl).DemandChpStrPlus, "Valeur de type Texte")
            End If
            If v_operateur = "COMPRIS" Then
                g_form_active = False
                Me.txtVal(v_indice).Text = d1 & " " & d2
                Me.txtVal(v_indice).tag = d1 & " " & d2
                g_form_active = True
            Else
                g_form_active = False
                Me.txtVal(v_indice).Text = stext
                Me.txtVal(v_indice).tag = stext
                g_form_active = True
            End If
            GoTo lab_fin
        Else
            sql = "select FOREC_Type, FOREC_Valeurs_Possibles, forec_fctvalid from FormEtapeChp where FOREC_Num=" & tbl_Demande(IndTbl).DemandChpNum
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If

            If stype = "CHECK" Or stype = "RADIO" Or stype = "SELECT" Then
                sval = rs("FOREC_Valeurs_Possibles")
                GoTo Lab_ValChp
            ElseIf stype = "HIERARCHIE" Then
                sval = rs("FOREC_Valeurs_Possibles")
                GoTo Lab_ValHierar
            ElseIf sfct_valid = "%NUMSERVICE" Then
                sfct_valid = "%NUMSERVICE"
                GoTo Lab_NumService
            ElseIf sfct_valid = "%NUMFCT" Then
                sfct_valid = "%NUMFCT"
                GoTo Lab_NumFct
            End If
        End If
        Call CL_InitTitreHelp(tbl_Demande(IndTbl).DemandChpStr, "")
        Call CL_InitTaille(0, -5)
        Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        ChoixListe.Show 1
        ' Quitter
        If CL_liste.retour = 1 Then
            Exit Sub
        End If
    
        txtOper(v_indice).tag = CL_liste.lignes(CL_liste.pointeur).tag
        txtOper(v_indice).Text = CL_liste.lignes(CL_liste.pointeur).texte
        
        
    End If
    
    
Fin:
    Exit Sub
    
Lab_NumService:
    If sfct_valid = "%NUMSERVICE" Then
        Call CL_Init
        n = STR_GetNbchamp(txtVal(v_indice).tag, ";")
        For i = 0 To n - 1
            II = STR_GetChamp(txtVal(v_indice).tag, ";", i)
            Call CL_AddLigne(II, 0, STR_GetChamp(txtVal(v_indice).tag, ";", i), True, True)
        Next i
        Set frm = KS_PrmService
        sret = KS_PrmService.AppelFrm("Choix des services", "C", True, "", "S", False)
        Set frm = Nothing
        If sret = "" Then
            Exit Sub
        End If
        If sret = "N0" Then
            'Call Odbc_RecupVal("select L_Code from Laboratoire", nom_srv)
            'stag = stag & "S0;"
            stag = "N0"
            nom_srv = "Tout le site"
            stext = nom_srv
            GoTo lab_affiche
        Else
            stag = ""
            stext = ""
            For lig = 0 To UBound(CL_liste.lignes())
                num_srv = Mid$(STR_GetChamp(CL_liste.lignes(lig).texte, ";", STR_GetNbchamp(CL_liste.lignes(lig).texte, ";") - 1), 2)
                stag = stag & "S" & num_srv & ";"
                If stext <> "" Then
                    stext = stext & ", "
                End If
                Call P_RecupSrvNom(num_srv, nom_srv)
                stext = stext & nom_srv
            Next lig
            GoTo lab_affiche
        End If
    End If

Lab_NumFct:
    If sfct_valid = "%NUMFCT" Then
        sret = choisir_fonctions(txtVal(v_indice).tag)
        stag = ""
        stext = ""
        If sret <> "" Then
            n = STR_GetNbchamp(sret, ";")
            For i = 0 To n - 1
                num_fct = Mid$(STR_GetChamp(sret, ";", i), 2)
                stag = stag & "F" & num_fct & ";"
                If stext <> "" Then
                    stext = stext & ", "
                End If
                Call Odbc_RecupVal("select ft_libelle from fcttrav where ft_num=" & num_fct, nom_fct)
                stext = stext & nom_fct
            Next i
        End If
        GoTo lab_affiche
    End If
    
Lab_ValChp:
    numlst = sval
    sql = "select VC_Num, VC_Lib from ValChp" _
        & " where VC_LVCNum=" & numlst _
        & " order by VC_Ordre"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Aucune valeur n'a été trouvée.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    Call CL_Init
    Call CL_AddLigne("<Non renseigné>", 0, 0, IIf(InStr(txtVal(v_indice).tag, "<NR>;") > 0, True, False))
    While Not rs.EOF
        Call CL_AddLigne(rs("VC_Lib").Value, rs("VC_Num").Value, "", IIf(InStr(txtVal(v_indice).tag, "V" & rs("VC_Num").Value & ";") > 0, True, False))
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Liste des valeurs", p_chemin_appli + "\help\kalidoc.chm" & ";" & "form_etape.htm")
    If UBound(CL_liste.lignes) < 10 Then
        n = UBound(CL_liste.lignes) + 1
    Else
        n = 10
    End If
    Call CL_InitTaille(0, -n)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    stag = ""
    stext = ""
    For lig = 0 To UBound(CL_liste.lignes())
        If CL_liste.lignes(lig).selected Then
            If CL_liste.lignes(lig).num = 0 Then
                stag = "<NR>;"
            Else
                stag = stag & "V" & CL_liste.lignes(lig).num & ";"
            End If
            If stext <> "" Then
                stext = stext & ", "
            End If
            stext = stext & CL_liste.lignes(lig).texte
        End If
    Next lig
    GoTo lab_affiche
    
Lab_ValHierar:
    Dim selected As Boolean
    Dim Detail As Boolean
    Dim laS As String
    Dim UneS As String
    
    numlst = sval
    sql = "select * from Hierarvalchp" _
        & " where HVC_LHCNum=" & numlst _
        & " and HVC_Numpere=0" _
        & " order by HVC_Ordre"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Aucune valeur n'a été trouvée.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    Call CL_Init
    Call CL_AddLigne("<Non renseigné>", 0, 0, IIf(InStr(txtVal(v_indice).tag, "<NR>;") > 0, True, False))
    s = txtVal(v_indice).tag
    Detail = True   ' IIf(v_operateur = "(D)", True, False)
    While Not rs.EOF
        selected = False
        For i = 0 To STR_GetNbchamp(s, ";")
            UneS = STR_GetChamp(s, ";", i)
            If UneS <> "" Then
                If UneS = "M" & rs("HVC_Num").Value Then
                    selected = True
                End If
            End If
        Next i
        Call CL_AddLigne(rs("HVC_Nom").Value, rs("HVC_Num").Value, "", selected)
        Call AjoutValHierar(rs("HVC_Num"), "...", selected, Detail, s)
        rs.MoveNext
    Wend
    rs.Close
    
    Call CL_InitTitreHelp("Liste des valeurs", "")
    If UBound(CL_liste.lignes) < 10 Then
        n = UBound(CL_liste.lignes) + 1
    Else
        n = 10
    End If
    Call CL_InitTaille(0, -n)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    stag = ""
    stext = ""
    For lig = 0 To UBound(CL_liste.lignes())
        If CL_liste.lignes(lig).selected Then
            If CL_liste.lignes(lig).num = 0 Then
                stag = "<NR>;"
            Else
                stag = stag & "M" & CL_liste.lignes(lig).num & ";"
            End If
            If stext <> "" Then
                stext = stext & ", "
            End If
            stext = stext & Replace(CL_liste.lignes(lig).texte, "...", "")
        End If
    Next lig
    ' on met toujours en mode détail
    stag = stag & "_DET"
    
lab_affiche:
    txtVal(v_indice).tag = stag
    
    txtVal(v_indice).Text = stext
    
    Bool_Faire_TvVal_Gotfocus = False
    txtVal(v_indice).SetFocus
    Bool_Faire_TvVal_Gotfocus = False
    
    cmd(CMD_OK).Enabled = True

lab_fin:
    cmd(CMD_OK).Enabled = True
        
    ' calcul de la condition
    FctCalculCondition (v_indice)
    'Passer au suivant
    If v_Trait = "Boucle" Then
        IndLigneCourrante = v_indice
        ChpCourrant = "VAL"
        FctPasserSuivant IndLigneCourrante, ChpCourrant, v_Trait
    End If
End Sub

Private Function AjoutValHierar(ByVal HVC_Num As String, ByVal Padd As String, ByVal selected As Boolean, ByVal Detail As Boolean, laS As String)
    Dim sqlH As String, rsH As rdoResultset
    Dim i As Integer
    Dim UneS As String
    Dim selectedF As Boolean
    
    sqlH = "Select * from Hierarvalchp where HVC_numpere=" & HVC_Num
    Call Odbc_SelectV(sqlH, rsH)
    If Not rsH.EOF Then
        'If Detail And selected Then
        '    selectedF = True
        'End If
        While Not rsH.EOF
            selectedF = False
            For i = 0 To STR_GetNbchamp(laS, ";")
                UneS = STR_GetChamp(laS, ";", i)
                If UneS <> "" Then
                    If UneS = "M" & rsH("HVC_Num").Value Then
                        selectedF = True
                    End If
                End If
            Next i
            Call CL_AddLigne(Padd & " " & rsH("HVC_Nom").Value, rsH("HVC_Num").Value, "", selectedF)
            Call AjoutValHierar(rsH("HVC_Num"), Padd & "...", selectedF, Detail, laS)
            rsH.MoveNext
        Wend
    End If
End Function
Private Sub initialiser()

    Dim nbdem As Integer, iD As Integer
    Dim LeTop As Integer
    Dim sql As String, rs As rdoResultset
    Dim idLig As Integer
    Dim numfct As Long
    Dim TagOper As String, TagVal As String, TagChp As String
    Dim n As Integer, i As Integer, j As Integer
    Dim iDem As Integer
    Dim nb As Integer, nbs As Integer
    Dim numval As String, s As String, stext As String
    'If g_initFait Then Exit Sub
    
    Me.lblFF(0).Visible = False
    On Error GoTo err_TabDem
    nbdem = UBound(tbl_Demande)
    GoTo SuiteDem
err_TabDem:
    Resume Fin
SuiteDem:
    On Error GoTo 0
    LeTop = Me.txtChp(0).Top
    idLig = 1
    
    nbdem = STR_GetNbchamp(g_sQ, ";")
    For iDem = 0 To nbdem
        s = STR_GetChamp(g_sQ, ";", iDem)
        If s <> "" Then
            iD = s
'        If tbl_Demande(iD).DemandFFNum = g_FF_Num Then
'            If tbl_Demande(iD).DemandGlobale Or tbl_Demande(iD).DemandFormInd = g_FF_Indice Then
                ' Déjà posée ?
                'If Not tbl_Demande(iD).DemandFait Then
                    sql = "select * from formetapechp where Forec_Num = " & tbl_Demande(iD).DemandChpNum
                    'MsgBox sql
                    If Not Odbc_SelectV(sql, rs) = P_ERREUR Then
                        If Not rs.EOF Then
                            tbl_Demande(iD).DemandType = rs("Forec_Type")
                            tbl_Demande(iD).DemandFctValid = rs("Forec_FctValid")
                            Load Me.txtChp(idLig)
                            Me.txtChp(idLig).Top = LeTop
                            Me.txtChp(idLig).Visible = True
                            'tbl_Demande(iD).DemandChpStr = rs("Forec_Label")
                            Me.txtChp(idLig).Text = tbl_Demande(iD).DemandChpStr & " " & tbl_Demande(iD).DemandChpStrPlus   ' rs("Forec_Label")
                            rs.Close
                            ' sert de lien avec le tableau
                            Me.txtChp(idLig).tag = iD
                            ' retrouver la désignation du filtre
                            For j = 0 To PiloteExcelBis.grdForm.Rows - 1
                                If tbl_Demande(iD).DemandFFNum = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Num) Then
                                    If tbl_Demande(iD).DemandFormInd = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_NumIndice) Then
                                        Load Me.lblFF(idLig)
                                        Me.lblFF(idLig).Top = LeTop
                                        Me.lblFF(idLig).Visible = True
                                        Me.lblFF(idLig).Caption = PiloteExcelBis.grdForm.TextMatrix(j, GrdForm_FF_Lib)
                                        Me.lblFF(idLig).tag = iD
                                        Exit For
                                    End If
                                End If
                            Next j
                            
                            Load Me.txtCnd(idLig)
                            Me.txtCnd(idLig).Top = LeTop
                            Me.txtCnd(idLig).Visible = False
                            Me.txtCnd(idLig).Text = tbl_Demande(iD).DemandenSQL
                            
                            TagChp = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 0)
                            TagChp = STR_GetChamp(TagChp, ":", 1)
                            TagOper = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 1)
                            TagOper = STR_GetChamp(TagOper, ":", 1)
                            TagVal = STR_GetChamp(tbl_Demande(iD).DemandPasFrancais, "¤", 2)
                            TagVal = STR_GetChamp(TagVal, ":", 1)
                            
                            Load Me.txtOper(idLig)
                            Me.txtOper(idLig).Top = LeTop
                            Me.txtOper(idLig).Visible = True
                            Me.txtOper(idLig).tag = TagOper
                            If TagOper <> "" Then
                                If TagOper = "=" Then Me.txtOper(idLig).Text = "Egal"
                                If TagOper = "!" Then Me.txtOper(idLig).Text = "Différent"
                                If TagOper = ">" Then Me.txtOper(idLig).Text = "Supérieur"
                                If TagOper = ">=" Then Me.txtOper(idLig).Text = "Supérieur ou Egal"
                                If TagOper = "<" Then Me.txtOper(idLig).Text = "Inférieur"
                                If TagOper = "<=" Then Me.txtOper(idLig).Text = "Inférieur ou Egal"
                                If TagOper = "COMPRIS" Then Me.txtOper(idLig).Text = "Compris entre "
                            End If
                            
                            Load Me.txtVal(idLig)
                            Me.txtVal(idLig).Top = LeTop
                            Me.txtVal(idLig).Visible = True
                            Me.txtVal(idLig).tag = TagVal
                            If TagVal <> "" Then
                                ' selon le type
                                'MsgBox tbl_Demande(iD).DemandType
                                If tbl_Demande(iD).DemandType = "TEXT" Then
                                    Me.txtVal(idLig).Text = TagVal
                                    Me.txtVal(idLig).tag = TagVal
                                    If TagVal = "N0" Then
                                        Me.txtVal(idLig).Text = "Tout le site"
                                    Else
                                        If tbl_Demande(iD).DemandFctValid = "%NUMSERVICE" Then
                                            nbs = STR_GetNbchamp(TagVal, ";")
                                            s = STR_GetChamp(TagVal, ";", 0)
                                            s = Replace(s, "S", "")
                                            s = Replace(s, ";", "")
                                            If IsNumeric(s) Then
                                                Call P_RecupSrvNom(s, s)
                                                Me.txtVal(idLig).Text = s & IIf(nbs > 1, " , ...", "")
                                            Else
                                                MsgBox "service " & TagVal & " invalide"
                                                Me.txtVal(idLig).Text = ""
                                            End If
                                        ElseIf tbl_Demande(iD).DemandFctValid = "%NUMFCT" Then
                                            s = Replace(TagVal, "F", "")
                                            n = STR_GetNbchamp(TagVal, ";")
                                            For i = 0 To n - 1
                                                numfct = STR_GetChamp(TagVal, ";", i)
                                                numfct = Replace(numfct, "F", "")
                                                Call P_RecupNomFonction(s, s)
                                            Next i
                                            Me.txtVal(idLig).Text = s
                                        End If
                                    End If
                                ElseIf tbl_Demande(iD).DemandType = "SELECT" Or tbl_Demande(iD).DemandType = "RADIO" Or tbl_Demande(iD).DemandType = "CHECK" Then
                                    n = STR_GetNbchamp(TagVal, ";")
                                    For i = 0 To n - 1
                                        'numval = Mid$(STR_GetChamp(TagVal, ";", i), 2)
                                        numval = STR_GetChamp(TagVal, ";", i)
                                        If numval = "<NR>" Then
                                            If stext <> "" Then
                                                stext = stext + ", "
                                            End If
                                            stext = stext + "non renseignée"
                                        Else
                                            numval = Replace(numval, "V", "")
                                            sql = "select VC_Lib from ValChp" _
                                                & " where VC_Num=" & numval
                                            s = ""
                                            Call Odbc_RecupVal(sql, s)
                                            If stext <> "" Then
                                                stext = stext + ", "
                                            End If
                                            stext = stext + s
                                        End If
                                    Next i
                                    Me.txtVal(idLig).Text = stext
                                ElseIf tbl_Demande(iD).DemandType = "HIERARCHIE" Then
                                    n = STR_GetNbchamp(TagVal, ";")
                                    For i = 0 To n - 1
                                        'numval = Mid$(STR_GetChamp(TagVal, ";", i), 2)
                                        numval = STR_GetChamp(TagVal, ";", i)
                                        If numval = "_DET" Then
                                        ElseIf numval = "<NR>" Then
                                            If stext <> "" Then
                                                stext = stext + ", "
                                            End If
                                            stext = stext + "non renseignée"
                                        Else
                                            numval = Replace(numval, "M", "")
                                            sql = "select HVC_Nom from HierarValChp" _
                                                & " where HVC_Num=" & numval
                                            s = ""
                                            Call Odbc_RecupVal(sql, s)
                                            If stext <> "" Then
                                                stext = stext + ", "
                                            End If
                                            stext = stext + s
                                        End If
                                    Next i
                                    If InStr(TagVal, "_DET") > 0 Then
                                        stext = "(Détail) " & stext
                                    End If
                                    Me.txtVal(idLig).Text = stext
                                End If
                            End If
                            
                            Load Me.cmdOper(idLig)
                            Me.cmdOper(idLig).Top = LeTop
                            Me.cmdOper(idLig).Visible = True
                            
                            Load Me.cmdVal(idLig)
                            Me.cmdVal(idLig).Top = LeTop
                            Me.cmdVal(idLig).Visible = True
                            
                            Load Me.cmdAussi(idLig)
                            Me.cmdAussi(idLig).Top = LeTop
                            If PeutGlobal(val(tbl_Demande(iD).DemandChpNum)) Then
                                Me.cmdAussi(idLig).Visible = True
                                Me.cmdAussi(idLig).tag = tbl_Demande(iD).DemandAussiStr
                                If tbl_Demande(iD).DemandAussiStr <> "" Then
                                    cmdAussi(idLig).BackColor = cmdBoucle.BackColor
                                    cmdAussi(idLig).Width = 600
                                Else
                                    cmdAussi(idLig).BackColor = cmdBoucle.BackColor
                                    cmdAussi(idLig).Width = 320
                                End If
                            Else
                                Me.cmdAussi(idLig).Visible = False
                            End If
                            On Error Resume Next
                            Me.lblFF(idLig).left = cmdAussi(idLig).left + 600 + 100
                            
                            LeTop = LeTop + 400
                            idLig = idLig + 1
                        End If
                    End If
                'End If
            End If
        'End If
    Next iDem
Fin:
    On Error GoTo 0
    
    If idLig = 1 Then
        Call quitter(False)
        Exit Sub
    End If
    Call FRM_ResizeForm(Me, Me.Width, Me.Height)
    
    g_mode_saisie = False
    g_initFait = True
    g_form_active = True
    
    ' on commence
    NbLigTotal = idLig - 1
    Me.cmdBoucle.Visible = True
    frm.Caption = g_FF_oper
End Sub

Private Function EstMemeType(v_numchp1, v_numchp2)
    Dim sql As String
    Dim rs As rdoResultset
    Dim numchp As Long
    Dim s As String
    
    If v_numchp1 = v_numchp2 Then
        EstMemeType = True
        Exit Function
    End If
    sql = "select * from formetapechp where forec_num=" & v_numchp1 & " OR forec_num=" & v_numchp2
    
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)

    If Not rs.EOF Then
        s = rs("Forec_FctValid")
        On Error GoTo LabErr
        rs.MoveNext
        If s = rs("Forec_FctValid") Then
            EstMemeType = True
        Else
            EstMemeType = False
        End If
    End If
    Exit Function
LabErr:
    On Error GoTo 0
    EstMemeType = False
End Function

Private Function PeutGlobal(v_numchp As Long)
    Dim sql As String
    Dim rs As rdoResultset
    Dim numchp As Long
    
    sql = "select * from formetapechp where forec_num=" & v_numchp
    
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)

    If Not rs.EOF Then
        'MsgBox rs("Forec_FctValid") & " " & rs("Forec_Type")
        If rs("Forec_FctValid") = "%NUMSERVICE" Then
            PeutGlobal = True
        ElseIf InStr(rs("Forec_FctValid"), "%DATE") > 0 Then
            PeutGlobal = True
        ElseIf InStr("SELECT*CHECK*RADIO", rs("Forec_Type")) > 0 And rs("forec_valeurs_possibles") > 0 Then
            PeutGlobal = True
        Else
            PeutGlobal = False
        End If
    End If

End Function


Private Sub quitter(ByVal v_bforce As Boolean)

    Dim reponse As Integer
    
    If v_bforce Then
        Unload Me
        Exit Sub
    End If
        
    If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
    End If
    
    Unload Me
    
End Sub

Private Sub valider()

    Call P_MAJ(g_sQ)
    Call quitter(True)

End Sub

Private Sub MAJ()
    Dim iDem As Integer, nbdem As Integer, iD As Integer, idLig As Integer
    Dim iSc As Integer
    Dim i_tbl_RDOF As Integer
    Dim s As String, sql As String
    Dim i As Integer
    Dim rs As rdoResultset
    Dim nomAnc As String
    Dim nomS As String
    Dim i_F As Integer
    Dim j As Integer
    Dim laS As String
    Dim sqlRet As String, opSQL As String
    Dim ff_num As Long, ff_indice As Long
    Dim chpnum As String
    
    For i_F = 0 To UBound(tbl_rdoF())
        tbl_rdoF(i_F).RDOF_AussiQuestionsFait = ""
        tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = ""
        tbl_rdoF(i_F).RDOF_AussiChpNum = ""
        tbl_rdoF(i_F).RDOF_AussiChpType = ""
        tbl_rdoF(i_F).RDOF_AussiFctValid = ""
    Next i_F
    nbdem = UBound(tbl_Demande)
    For iDem = 0 To nbdem
        If tbl_Demande(iDem).DemandChpNum = -1 Then GoTo Lab_Next_iDem
        If tbl_Demande(iDem).DemandAussiBool Then
            If tbl_Demande(iDem).DemandAussiStr <> "" Then
                ' ex : 183:1:3170;69:1:137;   ff_num  ff_indice   forec_num
                For iSc = 0 To STR_GetNbchamp(tbl_Demande(iDem).DemandAussiStr, ";")
                    s = STR_GetChamp(tbl_Demande(iDem).DemandAussiStr, ";", iSc)
                    If s <> "" Then
                        ' on applique à l'autre (RDO_F)
                        Call Odbc_RecupVal("select forec_nom from formetapechp where forec_num=" & tbl_Demande(iDem).DemandChpNum, nomAnc)
                        chpnum = STR_GetChamp(s, ":", 2)
                        ' remplacer le champ
                        ff_num = STR_GetChamp(s, ":", 0)
                        ff_indice = STR_GetChamp(s, ":", 1)
                        For i_F = 0 To UBound(tbl_rdoF())
                            If tbl_rdoF(i_F).RDOF_num = ff_num Then
                                If tbl_rdoF(i_F).RDOF_FormIndice = ff_indice Then
                                    Call Odbc_RecupVal("select forec_nom from formetapechp where forec_num=" & chpnum, nomS)
                                    s = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL
                                    tbl_rdoF(i_F).RDOF_Aussi_iDem = tbl_rdoF(i_F).RDOF_Aussi_iDem & IIf(s = "", "", "¤") & iDem
                                    If Not tbl_Demande(iDem).DemandFait Then
                                        tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & "TBL_DEMANDE:" & iDem
                                    Else
                                        tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & Replace(tbl_Demande(iDem).DemandenSQL, nomAnc, nomS)
                                        ' modifier le champ
                                        'laS = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL
                                        's = STR_GetChamp(STR_GetChamp(laS, ":", 0), "|", 0) & ":" & nomS & "|" & STR_GetChamp(laS, "|", 1) & "|" & STR_GetChamp(laS, "|", 2)
                                        'tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = s
                                    End If
                                    'If g_Trait = "PARAM" Then
                                    '    ' on enregistre seulement un pointeur vers le tbl_demande
                                    '    tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & "TBL_DEMANDE:" & iDem
                                    'Else
                                    '    tbl_rdoF(i_F).RDOF_AussiQuestionsSQL = tbl_rdoF(i_F).RDOF_AussiQuestionsSQL & IIf(s = "", "", "¤") & Replace(tbl_Demande(iDem).DemandenSQL, nomAnc, nomS)
                                    'End If
                                    tbl_rdoF(i_F).RDOF_AussiQuestionsFait = tbl_rdoF(i_F).RDOF_AussiQuestionsFait & IIf(s = "", "", "¤") & IIf(tbl_Demande(iDem).DemandFait, "T", "F")
                                    tbl_rdoF(i_F).RDOF_AussiChpNum = tbl_rdoF(i_F).RDOF_AussiChpNum & IIf(s = "", "", "¤") & chpnum
                                    tbl_rdoF(i_F).RDOF_AussiChpType = tbl_rdoF(i_F).RDOF_AussiChpType & IIf(s = "", "", "¤") & tbl_Demande(iDem).DemandType
                                    tbl_rdoF(i_F).RDOF_AussiFctValid = tbl_rdoF(i_F).RDOF_AussiFctValid & IIf(s = "", "", "¤") & tbl_Demande(iDem).DemandFctValid
                                End If
                            End If
                        Next i_F
                        'If Not tbl_Demande(iDem).DemandFait Then
                        '    If g_Trait <> "PARAM" Then
                        '        tbl_Demande(iDem).DemandFait = True
                        '    End If
                        'End If
                    End If
                Next iSc
            End If
        End If
Lab_Next_iDem:
    Next iDem
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_QUITTER
        Call quitter(False)
    End Select
    
End Sub

Private Sub cmdAussi_Click(Index As Integer)

    Call choisir_Aussi(Index)

End Sub

Private Sub cmdBoucle_Click()
    IndLigneCourrante = 0
    ChpCourrant = "VAL"
    FctPasserSuivant IndLigneCourrante, ChpCourrant, "Boucle"
End Sub

Private Sub cmdOper_Click(Index As Integer)

    Call choisir_oper(Index, "Direct")

End Sub

Private Sub cmdVal_Click(Index As Integer)

    Call choisir_valeur(Index, "Direct", txtOper(Index).tag)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then
            Call valider
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False
    Call initialiser
    
End Sub

Private Sub txt_Change(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
End Sub

Private Sub txtOper_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call choisir_oper(Index, "Direct")
    End If
    
End Sub

Private Sub txtVal_Change(Index As Integer)
    Dim IndTbl As Integer
    Dim stype As String
    Dim stext As String
    
    If g_form_active Then
        If Me.txtVal(Index).Text <> "" Then
            IndTbl = Me.txtChp(Index).tag
            stype = tbl_Demande(IndTbl).DemandType
            If stype = "TEXT" Then
                If InStr(tbl_Demande(IndTbl).DemandFctValid, "%DATE") > 0 Then
                    cmdVal_Click (Index)
                End If
                'Me.txtVal(index).tag = Me.txtVal(index).Text
                ' calcul de la condition
                FctCalculCondition (Index)
            End If
        End If
    End If
End Sub

Private Sub txtVal_GotFocus(Index As Integer)
    Dim IndTbl As Integer
    Dim stype As String
    Dim stext As String
    
    If Me.txtVal(Index).Text = "" Then
        IndTbl = Me.txtChp(Index).tag
        stype = tbl_Demande(IndTbl).DemandType
        If stype = "TEXT" Then
            stext = InputBox("Valeur pour " & tbl_Demande(IndTbl).DemandChpStr, "Valeur de type Texte", Me.txtVal(Index).Text)
            Me.txtVal(Index).Text = stext
            Me.txtVal(Index).tag = stext
            ' calcul de la condition
            FctCalculCondition (Index)
        End If
        If stype = "SELECT" Or stype = "RADIO" Or stype = "CHECK" Then
            Call choisir_valeur(Index, "Direct", txtOper(Index).tag)
        End If
    Else
        IndTbl = Me.txtChp(Index).tag
        stype = tbl_Demande(IndTbl).DemandType
        If stype = "TEXT" And tbl_Demande(IndTbl).DemandFctValid <> "%NUMSERVICE" And tbl_Demande(IndTbl).DemandFctValid <> "%NUMFCT" Then
            Me.txtVal(Index).tag = Me.txtVal(Index).Text
            ' calcul de la condition
            FctCalculCondition (Index)
        End If
    End If
End Sub

Private Sub txtVal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call choisir_valeur(Index, "Direct", txtOper(Index).tag)
    End If
    
End Sub
