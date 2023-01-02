Attribute VB_Name = "MPublisher"
Option Explicit

Public Pub_obj As Publisher.Application
Public Pub_doc As Publisher.Document

Public Function Publisher_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If Publisher_Init() = P_ERREUR Then
        Publisher_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Publisher_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Publisher_OuvrirDoc(v_nomdoc, v_passwd, Pub_doc) = P_ERREUR Then
        Publisher_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Pub_doc.ActiveWindow.visible = True
    
    encore = True
    Do
        Call SYS_Sleep(500)
        DoEvents
        On Error Resume Next
        If Not Pub_doc.ActiveWindow.visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Pub_doc = Nothing
    Set Pub_obj = Nothing
    
    Publisher_AfficherDoc = P_OK

End Function

Public Sub Publisher_Imprimer(ByVal v_nomdoc As String, _
                             ByVal v_passwd As String, _
                             ByVal v_nomimp As String, _
                             ByVal v_nbex As Integer)

    If Publisher_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If Publisher_OuvrirDoc(v_nomdoc, v_passwd, Pub_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Pub_doc.ActivePrinter = v_nomimp
    Call Pub_doc.PrintOut(Copies:=v_nbex)
    
    Call Pub_doc.Close
    Set Pub_doc = Nothing
    Pub_obj.Application.Quit
    Set Pub_obj = Nothing
    
End Sub

Public Function Publisher_Init()

    On Error GoTo err_create_obj
    Set Pub_obj = CreateObject("publisher.application")
    On Error GoTo 0
    
    Publisher_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet PUBLISHER." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Publisher_Init = P_ERREUR
    Exit Function

End Function

Public Function Publisher_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Publisher.Document) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Pub_obj.Open(Filename:=v_nomdoc)
    On Error GoTo 0
   
    Publisher_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Publisher_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function

