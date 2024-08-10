Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim shouldHide As Boolean

    ' Définir la feuille sur laquelle vous travaillez (par exemple, "Feuil1")
    Set ws = ThisWorkbook.Sheets("Feuil1")
    
    ' Trouver la dernière ligne utilisée dans la feuille
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Boucle sur chaque ligne
    For Each rng In ws.Range("A1:A" & lastRow).Rows
        shouldHide = False
        
        ' Boucle sur chaque cellule de la ligne
        For Each cell In rng.Cells
            If cell.Value <> "" Then
                shouldHide = True
                Exit For
            End If
        Next cell

        ' Masquer la ligne si une cellule est remplie
        If shouldHide Then
            rng.EntireRow.Hidden = True
        End If
    Next rng
End Sub
