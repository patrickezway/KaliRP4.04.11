Attribute VB_Name = "MVisio"
Option Explicit

Public Visio_obj As Visio.Application
Public Visio_doc As Visio.Document

Public Function Visio_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If Visio_Init() = P_ERREUR Then
        Visio_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Visio_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Visio_OuvrirDoc(v_nomdoc, v_passwd, Visio_doc) = P_ERREUR Then
        Visio_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Visio_obj.ActiveWindow.visible = True
    
    encore = True
    Do
        Call SYS_Sleep(500)
        DoEvents
        On Error Resume Next
        If Not Visio_obj.ActiveWindow.visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Visio_doc = Nothing
    Set Visio_obj = Nothing
    
    Visio_AfficherDoc = P_OK

End Function

Public Sub Visio_ConvHTML(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nomhtml As String)

    Dim webSettings As VisWebPageSettings
    Dim saveAsWeb As VisSaveAsWeb
    
    If Visio_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    Visio_obj.visible = False
    
    If Visio_OuvrirDoc(v_nomdoc, v_passwd, Visio_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Set saveAsWeb = Visio_obj.SaveAsWebObject
    Set webSettings = saveAsWeb.WebPageSettings
    Call saveAsWeb.AttachToVisioDoc(Visio_doc)
    ' Configure our preferences.
    webSettings.StartPage = 1
    webSettings.EndPage = 2
    webSettings.LongFileNames = True
    webSettings.QuietMode = True
    webSettings.PageTitle = ""
    webSettings.TargetPath = v_nomhtml
    Call saveAsWeb.CreatePages
    
    Call Visio_doc.Close
    Set Visio_doc = Nothing
    Visio_obj.Application.Quit
    Set Visio_obj = Nothing
    
End Sub

Public Sub Visio_Imprimer(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nbex As Integer)

    Dim dummy As String
    Dim I As Integer
    Dim doc_obj As Object
    
    If Visio_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    Visio_obj.visible = False
    
    If Visio_OuvrirDoc(v_nomdoc, v_passwd, Visio_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Set doc_obj = Visio_doc
    For I = 1 To v_nbex
        dummy = doc_obj.Print
    Next I
    
    Call Visio_doc.Close
    Set Visio_doc = Nothing
    Visio_obj.Application.Quit
    Set Visio_obj = Nothing
    
End Sub

Public Function Visio_Init()

    On Error GoTo err_create_obj
    Set Visio_obj = CreateObject("Visio.application")
    On Error GoTo 0
    
    Visio_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet Visio." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Visio_Init = P_ERREUR
    Exit Function

End Function

Public Function Visio_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Visio.Document) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Visio_obj.Documents.Open(Filename:=v_nomdoc)
    On Error GoTo 0
   
    Visio_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbcr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Visio_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function
