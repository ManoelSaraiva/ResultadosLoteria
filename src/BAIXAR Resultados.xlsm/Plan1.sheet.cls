Option Explicit

Private Sub btnAtualizaResultados_Click()
    Dim wSorteio        As Worksheet
    Dim wMenu           As Worksheet
    
    Set wSorteio = Sheets("Sorteio")
    Set wMenu = Sheets("Menu")
    
    'Desativa atualização para ganho de velocidade
    Application.ScreenUpdating = False
    
    'Seleciona Planilha Sorteio
    wSorteio.Select
    
    'Limpa resultados anteriores
    wSorteio.UsedRange.ClearContents
    
    'Seleciona a celula A1 da planulha sorteio
    wSorteio.Range("A1").Select
    
    'Baixa e extrai os resultados
    Call BaixarSorteios
    
    Call LerArquivo
       
    OrdenaResultados
       
    'Reativa atualização
    Application.ScreenUpdating = True
    
    'Apaga os arquivos extraidos
    Kill sLocalPlan & sFileHTM
    Kill sLocalPlan & sFileZip
    Kill sLocalPlan & sFileGIF
    
    btnVerResultados_Click
    
End Sub

Private Sub btnVerResultados_Click()

    Dim wSorteio        As Worksheet
    Dim wMenu           As Worksheet
    
    Dim iUltimaLinha    As Long
    
    Set wSorteio = Sheets("Sorteio")
    Set wMenu = Sheets("Menu")
    
    wSorteio.Select
    
    iUltimaLinha = wSorteio.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    wMenu.Range("M7").Value = wSorteio.Cells(iUltimaLinha - 1, 2).Value
    
    wSorteio.Cells(iUltimaLinha, 1).Select
       
End Sub

Private Sub OrdenaResultados()
    Dim wSorteio        As Worksheet
    Dim iUltimaLinha    As Integer
    Dim iLinha          As Integer

    Set wSorteio = ActiveWorkbook.Sheets("Sorteio")
    iUltimaLinha = wSorteio.Cells(Rows.Count, 3).End(xlUp).Row

    For iLinha = 1 To iUltimaLinha
        
        DoEvents
        
        With wSorteio.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("C" & iLinha)
            .SetRange Range(Cells(iLinha, 3), Cells(iLinha, 17))
            .Orientation = xlLeftToRight
            .SortMethod = xlPinYin
            .Apply
        End With
    
    Next iLinha
    
End Sub