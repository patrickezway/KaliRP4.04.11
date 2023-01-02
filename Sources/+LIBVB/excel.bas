Attribute VB_Name = "Mexcel"
Option Explicit

Public Exc_obj As Excel.Application
Public Exc_doc As Excel.Workbook

Public Function Excel_Fusionner(ByVal v_nomdoc As String, _
                           ByVal v_nommod As String, _
                           ByVal v_garder_styles As Boolean, _
                           ByVal v_nomdata As String, _
                           ByVal v_garder_bookmark As Boolean, _
                           ByVal v_nomdest As String, _
                           ByVal v_ecraser As Boolean, _
                           ByVal v_passwd As String, _
                           ByVal v_word_visible As Boolean, _
                           ByVal v_word_mode As Integer, _
                           ByVal v_nbex As Integer, _
                           ByVal v_deb_mode As Integer, _
                           ByVal v_fin_mode As Integer, _
                           Optional v_frm As Variant) As Integer
    
    Const PREMIER = 1
    Const DERNIER = 2
    Const AUTRE = 3
    
    Dim sval As String, NomFicStru As String, StrStru As String
    Dim StrDon As String, str As String, tb_Stru() As String, tb_Don() As String
    Dim b_fairefusion As Boolean, a_redim As Boolean
    Dim FeuilleEntete As Integer
    Dim I As Integer, j As Integer, fdStru As Integer, nbStru As Integer, nbDon As Integer
    Dim ichp As Integer, n As Integer, iUbound As Integer, tb_page(5) As Integer
    Dim doc2 As Object
    Dim wrk_doc As Excel.Workbook, wrk_mod As Excel.Workbook
    '
    If v_deb_mode = WORD_DEB_CROBJ Then
        If Excel_Init() = P_ERREUR Then
            Excel_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    If v_deb_mode <> WORD_DEB_RIEN Then
        If Excel_OuvrirDocW(v_nommod, v_passwd, wrk_mod) = P_ERREUR Then
            Excel_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    b_fairefusion = True
    If v_nomdata = "" Then
        b_fairefusion = False
        GoTo lab_fin_fusion
    End If
    
'v_word_visible = True
    If v_word_visible Then
        Exc_obj.visible = True
        a_redim = False
        On Error Resume Next
        If Exc_obj.WindowState <> wdWindowStateMaximize Then
            a_redim = True
        End If
        Exc_obj.Activate
        If a_redim Then
            Exc_obj.WindowState = wdWindowStateMaximize
        End If
        On Error GoTo 0
    Else
        Exc_obj.visible = False
    End If
    
    If v_nomdest <> "" Then
        On Error GoTo err_sav_dest
        If v_ecraser Then
            If v_passwd <> "" Then
                wrk_mod.Password = v_passwd
            End If
            'Call FICH_EffacerFichier(v_nomdest, False)
            'wrk_mod.SaveAs Filename:=v_nomdest
            On Error GoTo 0
        Else
            If v_deb_mode <> WORD_DEB_RIEN Then
                If Excel_OuvrirDocW(v_nomdest, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                On Error GoTo 0
                Call Excel_Doc.Close(savechanges:=wdDoNotSaveChanges)
                Set Excel_Doc = Exc_obj.ActiveDocument
            Else
                If Excel_OuvrirDocW(v_nommod, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                On Error GoTo 0
                Call doc2.Close(savechanges:=wdDoNotSaveChanges)
                Set doc2 = Nothing
            End If
        End If
    End If
    
    ' ouvrir le document
    If Excel_OuvrirDocW(v_nomdoc, v_passwd, wrk_doc) = P_ERREUR Then
        Excel_Fusionner = P_ERREUR
        Exit Function
    End If
    ' ouvrir le modèle
    ' enregistrer ses caractéristiques
    n = wrk_mod.Sheets.Count
    tb_page(PREMIER) = 1
    If n > 2 Then
        tb_page(AUTRE) = 2
        tb_page(DERNIER) = 3
    ElseIf n = 2 Then
        tb_page(AUTRE) = 2
        tb_page(DERNIER) = 2
    ElseIf n = 1 Then
        tb_page(AUTRE) = 1
        tb_page(DERNIER) = 1
    End If
    
    ' ouvrir le fichier de structure
    ' la première ligne contient la structure
    ' puis une ligne par donnée
    fdStru = FreeFile
    NomFicStru = v_nomdata
    Open NomFicStru For Input As #fdStru
    ' charger le fichier de structure dans un tableau
    Line Input #fdStru, StrStru
    nbStru = STR_GetNbchamp(StrStru, ";")
    For ichp = 0 To nbStru
        str = STR_GetChamp(StrStru, ";", ichp)
        If str <> "" Then
            iUbound = 1
            On Error Resume Next
            iUbound = UBound(tb_Stru) + 1
            On Error GoTo 0
            ReDim Preserve tb_Stru(iUbound)
            tb_Stru(iUbound) = str
        End If
    Next ichp
    ' lire les autres lignes pour les données
    On Error GoTo err_diverses
    Do While True
        If w_lire_fich(fdStru, StrDon) = P_NON Then Exit Do
        str = left(StrDon, Len(StrDon) - 1)
        iUbound = 1
        On Error Resume Next
        iUbound = UBound(tb_Don) + 1
        On Error GoTo 0
        ReDim Preserve tb_Don(iUbound)
        tb_Don(iUbound) = str
    Loop
    '
    For I = 1 To wrk_doc.Sheets.Count
        If Not IsMissing(v_frm) Then
            If Not v_frm Is Nothing Then
                v_frm.visible = True
                v_frm.Caption = "Mise à jour du document Excel : Feuille " & I & " / " & wrk_doc.Sheets.Count
                v_frm.Refresh
            End If
        End If
        ' chercher quelle entete il faut mettre
        FeuilleEntete = 0
        If I = 1 Then
            FeuilleEntete = tb_page(PREMIER)
        ElseIf I = wrk_doc.Sheets.Count Then
            FeuilleEntete = tb_page(DERNIER)
        Else
            FeuilleEntete = tb_page(AUTRE)
        End If
        On Error GoTo err_diverses
        ' recopier le contenu des entêtes
        wrk_doc.Sheets(I).PageSetup.LeftHeader = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.LeftHeader, tb_Stru, tb_Don)
        wrk_doc.Sheets(I).PageSetup.CenterHeader = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.CenterHeader, tb_Stru, tb_Don)
        wrk_doc.Sheets(I).PageSetup.RightHeader = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.RightHeader, tb_Stru, tb_Don)
        wrk_doc.Sheets(I).PageSetup.LeftFooter = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.LeftFooter, tb_Stru, tb_Don)
        wrk_doc.Sheets(I).PageSetup.CenterFooter = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.CenterFooter, tb_Stru, tb_Don)
        wrk_doc.Sheets(I).PageSetup.RightFooter = Excel_Fct_Fusion(wrk_mod.Sheets(FeuilleEntete).PageSetup.RightFooter, tb_Stru, tb_Don)
        If Not v_garder_styles Then
            wrk_doc.Sheets(I).PageSetup.LeftMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.LeftMargin
            wrk_doc.Sheets(I).PageSetup.RightMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.RightMargin
            wrk_doc.Sheets(I).PageSetup.TopMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.TopMargin
            wrk_doc.Sheets(I).PageSetup.BottomMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.BottomMargin
            wrk_doc.Sheets(I).PageSetup.HeaderMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.HeaderMargin
            wrk_doc.Sheets(I).PageSetup.FooterMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.FooterMargin
            wrk_doc.Sheets(I).PageSetup.PrintHeadings = wrk_mod.Sheets(FeuilleEntete).PageSetup.PrintHeadings
            wrk_doc.Sheets(I).PageSetup.PrintGridlines = wrk_mod.Sheets(FeuilleEntete).PageSetup.PrintGridlines
            wrk_doc.Sheets(I).PageSetup.PrintComments = wrk_mod.Sheets(FeuilleEntete).PageSetup.PrintComments
            wrk_doc.Sheets(I).PageSetup.CenterHorizontally = wrk_mod.Sheets(FeuilleEntete).PageSetup.CenterHorizontally
            wrk_doc.Sheets(I).PageSetup.CenterVertically = wrk_mod.Sheets(FeuilleEntete).PageSetup.CenterVertically
            wrk_doc.Sheets(I).PageSetup.Orientation = wrk_mod.Sheets(FeuilleEntete).PageSetup.Orientation
            wrk_doc.Sheets(I).PageSetup.Draft = wrk_mod.Sheets(FeuilleEntete).PageSetup.Draft
            wrk_doc.Sheets(I).PageSetup.PaperSize = wrk_mod.Sheets(FeuilleEntete).PageSetup.PaperSize
            wrk_doc.Sheets(I).PageSetup.FirstPageNumber = wrk_mod.Sheets(FeuilleEntete).PageSetup.FirstPageNumber
            wrk_doc.Sheets(I).PageSetup.HeaderMargin = wrk_mod.Sheets(FeuilleEntete).PageSetup.HeaderMargin
        End If
    Next I
    ' fermer le modèle
    Call wrk_mod.Close(savechanges:=wdDoNotSaveChanges)
    Set wrk_mod = Nothing
    
    Call wrk_doc.Save
    Set wrk_doc = Nothing
    
    Exc_obj.Application.Quit
    Set Exc_obj = Nothing
Fin:
    Exit Function

err_sav_dest:
    MsgBox "Impossible de sauvegarder le fichier dans " & v_nomdest & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1

lab_fin_err1:
    Call wrk_mod.Close(savechanges:=wdDoNotSaveChanges)
    Set wrk_mod = Nothing
    Call wrk_doc.Close(savechanges:=wdDoNotSaveChanges)
    Set wrk_doc = Nothing
    Set Exc_obj = Nothing
    Exc_obj.visible = False
    Excel_Fusionner = P_ERREUR
    Exit Function

lab_fin_fusion:
    
    If v_word_mode = WORD_IMPRESSION Then
            Close #fdStru
            'Word_Obj.ActivePrinter = Printer.DeviceName
            Exc_obj.WordBasic.FilePrintSetup Printer:=Printer.DeviceName, DoNotSetAsSysDefault:=1
            Excel_Doc.PrintOut Copies:=v_nbex
    ElseIf v_word_mode = WORD_VISU Then
        Close #fdStru
        sval = wrk_doc.FullName
        Call wrk_mod.Close(savechanges:=wdDoNotSaveChanges)
        Set wrk_mod = Nothing
        Call wrk_doc.Close(savechanges:=wdSaveChanges)
        Set wrk_doc = Nothing
        Exc_obj.Documents.Open Filename:=sval, readonly:=True, passworddocument:=v_passwd
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_MODIF Then
        Close #fdStru
        wrk_doc.Saved = True
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_CREATE Then
        Close #fdStru
        GoTo lab_fin_create
    End If
    
lab_fin_create:
    Call FICH_EffacerFichier(v_nomdata, False)
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call wrk_mod.Close(savechanges:=wdDoNotSaveChanges)
        Set wrk_mod = Nothing
        Call wrk_doc.Close(savechanges:=wdSaveChanges)
        Set wrk_doc = Nothing
        Set Exc_obj = Nothing
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
        '
    End If
    Exc_obj.visible = False
    Excel_Fusionner = P_OK
    Exit Function

lab_fin:
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call wrk_mod.Close(savechanges:=wdDoNotSaveChanges)
        Set wrk_mod = Nothing
        Call wrk_doc.Close(savechanges:=wdDoNotSaveChanges)
        Set wrk_doc = Nothing
        Set Exc_obj = Nothing
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
'
    End If
    Call FICH_EffacerFichier(v_nomdata, False)
    Excel_Fusionner = P_OK
    Exit Function

lab_fin_visible:
    Call FICH_EffacerFichier(v_nomdata, False)
    Exc_obj.visible = True
    Excel_EstActif = False
    Excel_Fusionner = P_OK
    Exit Function

err_diverses:
    If Err = 1004 Then
        ' impossible de définir la propriété
        Resume Next
    Else
        MsgBox (Err & " " & Error$)
        Resume Fin
        Resume Next
    End If
End Function

Public Function Excel_Fct_Fusion(str As String, tb_Stru, tb_Don) As String
    
    Dim I As Integer
    Dim posdeb As Integer, posfin As Integer
    Dim strOut As String
    Dim car As String
    Dim s As String
    Dim Name As String
    
    Excel_Fct_Fusion = str
    posdeb = 0
    For I = 1 To Len(str)
        car = Mid(str, I, 1)
        If car = "[" Then
            posdeb = I
            Name = ""
        ElseIf car = "]" Then
            ' on peut fusionner
            strOut = strOut & Excel_Faire_Fusion(Name, tb_Stru, tb_Don)
            posdeb = 0
        Else
            If posdeb > 0 Then
                Name = Name & car
            Else
                strOut = strOut & car
            End If
        End If
    Next I
    Excel_Fct_Fusion = strOut
End Function

Public Function Excel_Faire_Fusion(Name As String, tb_Stru, tb_Don)
    Dim I As Integer
    
    Excel_Faire_Fusion = "<?>"
    For I = 1 To UBound(tb_Stru)
        If tb_Stru(I) = Name Then
            Excel_Faire_Fusion = tb_Don(I)
            Exit For
        End If
    Next I
End Function

Public Function Excel_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If Excel_Init() = P_ERREUR Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc) = P_ERREUR Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Exc_obj.visible = True
    If Exc_obj.WindowState <> xlMaximized Then
        Exc_obj.WindowState = xlMaximized
    End If
    If Exc_obj.ActiveWindow.WindowState <> xlMaximized Then
        Exc_obj.ActiveWindow.WindowState = xlMaximized
    End If
    
    encore = True
    Do
        Call SYS_Sleep(500)
        DoEvents
        On Error Resume Next
        If Not Exc_obj.visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Exc_doc = Nothing
    Set Exc_obj = Nothing
    
    Excel_AfficherDoc = P_OK

End Function

Public Sub Excel_Imprimer(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nomimp As String, _
                          ByVal v_nbex As Integer)

    If Excel_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Call Exc_doc.PrintOut(, , v_nbex, , v_nomimp)
    
    Call Exc_doc.Close(savechanges:=False)
    Set Exc_doc = Nothing
    Exc_obj.Application.Quit
    Set Exc_obj = Nothing
    
End Sub

Public Function Excel_Init()

    On Error GoTo err_create_obj
    Set Exc_obj = CreateObject("excel.application")
    On Error GoTo 0
    
    Excel_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet EXCEL." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Excel_Init = P_ERREUR
    Exit Function

End Function

Public Function Excel_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Excel.Workbook) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Exc_obj.Workbooks.Open(Filename:=v_nomdoc, _
                                        readonly:=False, _
                                        Password:=v_passwd)
    ' Pour ScruteConv
    'Set r_doc = Exc_obj.Workbooks.Open(Filename:=v_nomdoc, _
    '                                    readonly:=True, _
    '                                    Password:=v_passwd, _
    '                                    UpdateLinks:=False)
    On Error GoTo 0
    
    Excel_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Excel_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function

Public Function Excel_OuvrirDocW(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Excel.Workbook) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Exc_obj.Workbooks.Open(Filename:=v_nomdoc, _
                                        readonly:=False, _
                                        Password:=v_passwd)
    On Error GoTo 0
    
    Excel_OuvrirDocW = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Excel_OuvrirDocW = P_ERREUR
    Exit Function
    
End Function

Private Function w_lire_fich(ByVal v_fd As Integer, _
                             ByRef a_ligne As Variant) As Integer

    On Error GoTo fin_fichier
    Line Input #v_fd, a_ligne
    On Error GoTo 0
    w_lire_fich = P_OUI
    Exit Function

fin_fichier:
    On Error GoTo 0
    w_lire_fich = P_NON

End Function


