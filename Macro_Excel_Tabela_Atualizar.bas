Attribute VB_Name = "Módulo2"
Sub Atualizar_Tabelas_Dinamicas()

    Dim ws As Worksheet
    Dim pt As PivotTable

    Application.ScreenUpdating = False

    ' Atualiza todas as tabelas dinâmicas do arquivo
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Dashboard atualizado com sucesso!", vbInformation

End Sub

