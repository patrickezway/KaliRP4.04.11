VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ooff_listener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public lafin As Boolean

Private interfaces(0) As String


Public Property Get Bridge_ImplementedInterfaces() As Variant

    Bridge_ImplementedInterfaces = interfaces

End Property


Private Sub Class_Initialize()

    ' executé par : Ooff_Doc.addEventListener objEventlistener
    interfaces(0) = "com.sun.star.lang.XEventListener"
    lafin = False

End Sub


Private Sub Class_Terminate()

    ' Executé lorsque l'on ferme le programme appelant (end de l'application)

End Sub


Public Sub disposing(ByVal source As Object)

    ' exécuté lorsque l'on ferme le document OO
    lafin = True

End Sub



