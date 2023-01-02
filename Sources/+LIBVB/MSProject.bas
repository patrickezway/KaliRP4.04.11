Attribute VB_Name = "MMSProject"
Option Explicit

Public MSProject_obj As MSProject.Application
'Public Project_doc As MSProject.Project

Public Function MSProject_AfficherDoc(ByVal v_nomdoc As String, _
                                    ByVal v_passwd As String, _
                                    ByVal v_fimprime As Boolean, _
                                    ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If MSProject_Init() = P_ERREUR Then
        MSProject_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        MSProject_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If MSProject_OuvrirDoc(v_nomdoc, v_passwd) = P_ERREUR Then
        MSProject_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    MSProject_obj.ActiveWindow.visible = True

    encore = True
    Do
        Call SYS_Sleep(10)
        On Error Resume Next
        If Not MSProject_obj.ActiveWindow.visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set MSProject_obj = Nothing
    
    MSProject_AfficherDoc = P_OK

End Function

Public Sub MSProject_Imprimer(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nbex As Integer)

    Dim dummy As String
    Dim I As Integer
    Dim doc_obj As Object
    
    If MSProject_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If MSProject_OuvrirDoc(v_nomdoc, v_passwd) = P_ERREUR Then
        Exit Sub
    End If
        
    'Project_obj.ActiveWindow.visible = False
    
    For I = 1 To v_nbex
        Call MSProject_obj.FilePrint(, , , , , , , , , , False)
    Next I
    
    MSProject_obj.FileClose pjDoNotSave
    MSProject_obj.Application.Quit
    Set MSProject_obj = Nothing
    
End Sub

Public Function MSProject_Init()

    On Error GoTo err_create_obj
    Set MSProject_obj = CreateObject("MSProject.Application")
    On Error GoTo 0
    
    MSProject_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet MSPRoject." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    MSProject_Init = P_ERREUR
    Exit Function

End Function

Public Function MSProject_OuvrirDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String) As Integer

    On Error GoTo err_open_ficr
    Call MSProject_obj.FileOpen(v_nomdoc, False, , , , , , , , , , , v_passwd)
    On Error GoTo 0
    
    MSProject_obj.ActiveWindow.visible = True
   
    MSProject_OuvrirDoc = 0
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    MSProject_OuvrirDoc = 1
    Exit Function
    
End Function




