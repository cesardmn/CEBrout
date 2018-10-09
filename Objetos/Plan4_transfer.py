'Chama macro ao atualiza célula'
    Private Sub Worksheet_Change(ByVal Target As Range)
        Application.ScreenUpdating = False
        Application.EnableEvents = False
                
                If Not Intersect(Target, Range("A1")) Is Nothing Then
                    Call TrataIn
                End If

                Range("A1:A5, D1, F1").Select
                Selection.ClearContents
                Range("A1").Select
                
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub
