Attribute VB_Name = "Module3"

Public Function hfhf()
    MsgBox "ici"
    Set Excel_Obj = CreateObject("Excel.Application")
    Excel_Obj.WindowState = xlMaximized   ' format plein �cran
    Excel_Obj.Visible = True              ' visible � l'�cran
    Excel_Obj.ShowWindowsInTaskbar = True ' visible dans la barre de t�ches
    Excel_Obj.DisplayFormulaBar = True    ' affichage de la barre de formule
    Excel_Obj.Caption = "Export Actions KaliDoc"
    Excel_Obj.Workbooks.Add               ' ajout d'un classeur Excel
    Excel_Obj.Worksheets(1).Name = "Feuille1"

End Function
Private Function Excel_Obj_onchange()

End Function

Private Sub xls_sheetactivate(ByVal sh As Object)

    ' essai pour trapper les events d'excel
    'Debug.Print Chr$(7) & "xls_sheetactivate"
    
End Sub

