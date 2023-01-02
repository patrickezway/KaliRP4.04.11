Attribute VB_Name = "Mexcel"
Option Explicit

Rem Public Exc_obj As Excel.Application
Public Exc_doc As Excel.Workbook

Public Function Excel_Recup_Valeur(ByVal v_nomdoc As String, _
                                  ByVal v_feuille As String, _
                                  ByVal v_cellule As String) As String
    If Excel_Init() = P_ERREUR Then
        Excel_Recup_Valeur = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Excel_Recup_Valeur = P_ERREUR
        Exit Function
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       "", _
                       Exc_doc, _
                       True) = "" Then
        Excel_Recup_Valeur = P_ERREUR
        Exit Function
    End If
    
    Exc_obj.Sheets(v_feuille).Activate
    Excel_Recup_Valeur = Exc_obj.ActiveSheet.Range(v_cellule)
    Exc_doc.Close
End Function
Public Function Excel_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    Dim nb As Integer
    
    If Excel_Init() = P_ERREUR Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc, _
                       True) = "" Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Exc_obj.Visible = True
    Exc_obj.DisplayFullScreen = True
    Exc_obj.DisplayFullScreen = False
    If Exc_obj.WindowState <> xlMaximized Then
        Exc_obj.WindowState = xlMaximized
    End If
    If Exc_obj.ActiveWindow.WindowState <> xlMaximized Then
        Exc_obj.ActiveWindow.WindowState = xlMaximized
    End If
    
    Erase p_TbFenetres
    For nb = 1 To Exc_obj.Sheets.Count
        ReDim Preserve p_TbFenetres(nb)
        p_TbFenetres(nb) = Exc_obj.Sheets(nb).Name
    Next nb
    encore = True
    Do
        Call SYS_Sleep(500)
        DoEvents
        On Error Resume Next
        If Not Exc_obj.Visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Exc_obj = Nothing
    
    Excel_AfficherDoc = P_OK

End Function

Public Sub Excel_Imprimer(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nbex As Integer)

    If Excel_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc, False) = P_ERREUR Then
        Exit Sub
    End If
    
    Call Exc_doc.PrintOut(, , v_nbex, , Printer.DeviceName)
    
    Call Exc_doc.Close(savechanges:=False)
    Set Exc_doc = Nothing
    Exc_obj.Quit
    Set Exc_obj = Nothing
    
End Sub

Public Function Excel_Init() As Integer

    On Error GoTo err_create_obj
    
    Dim xlApp As Excel.Application

    Set Exc_obj = GetExcelApp()
        
    On Error GoTo 0
    
    Excel_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet EXCEL." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Excel_Init = P_ERREUR
    Exit Function

End Function



Private Function GetExcelApp() As Excel.Application
    
    On Error GoTo err_startExcel
    
    Set GetExcelApp = GetObject(, "Excel.Application")
    Exit Function
    
err_startExcel:
    If Err.Number = 429 Then 'No current instance of Excel start up Excel
        Set GetExcelApp = GetObject("", "Excel.Application")
        'MsgBox "Démarrage d'Excel"
        Resume Next
    Else
        Set GetExcelApp = GetObject("", "Excel.Application")
        Resume Next
    End If

End Function


Public Function Excel_MetVisible()
    Exc_obj.DisplayFullScreen = False
    If Exc_obj.WindowState <> xlMaximized Then
        Exc_obj.WindowState = xlMaximized
    End If
    On Error Resume Next
    If Exc_obj.ActiveWindow.WindowState <> xlMaximized Then
        Exc_obj.ActiveWindow.WindowState = xlMaximized
        Exc_obj.Visible = True
    End If
    Exc_obj.Visible = True
End Function

Public Function Excel_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Excel.Workbook, _
                               ByVal b_BoolDisplayAlerts As Boolean) As String

    On Error GoTo err_open_ficr
    v_nomdoc = Replace(v_nomdoc, "/", "\")
    FctTrace ("Début Excel_OuvrirDoc " & v_nomdoc)
    Exc_obj.DisplayAlerts = b_BoolDisplayAlerts

    FctTrace ("Excel_OuvrirDoc avant r_doc")
    Set r_doc = Exc_obj.Workbooks.Open(FileName:=v_nomdoc, _
                                        ReadOnly:=False, _
                                        password:=v_passwd)
    FctTrace ("Excel_OuvrirDoc avant DisplayAlerts")
    Exc_obj.DisplayAlerts = True
    
    On Error GoTo 0
    
    Excel_OuvrirDoc = r_doc.Name
    FctTrace ("Fin Excel_OuvrirDoc name=" & r_doc.Name)
    Exit Function
    
err_open_ficr:
    FctTrace ("Erreur Excel_OuvrirDoc err=" & Err)
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Ouvrir sous Excel"
    Excel_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function


