Attribute VB_Name = "modArquivo"
Option Explicit

Function LimpaTag(ByRef texto As Variant)

    Dim iInicio, iTamanho   As Integer

    iInicio = InStr(4, texto, ">", 1) + 1
    iTamanho = InStr(10, texto, "<", 1) - iInicio
    
    texto = Mid$(texto, iInicio, iTamanho)
    
End Function


Public Function LerArquivo()

    Dim wSorteio        As Worksheet
    Dim sLocalPlan      As String
    Dim sArquivo        As String
    Dim iLinha          As Integer
    Dim iColuna         As Integer
    Dim vLinhaArq       As Variant
    Dim vStringLimpa    As Variant
    
    Dim aResultados(30, 2000) As Variant
    
    Dim dTempo As Date
    
    dTempo = Now()
    
    Set wSorteio = Sheets("Sorteio")
    
    wSorteio.Select
    
    wSorteio.Columns("B").NumberFormat = "mm/dd/yyyy"
    
    wSorteio.Range("A2").Select
     
    iLinha = 0
    iColuna = 0
    
    sLocalPlan = Application.ThisWorkbook.Path
    
    sArquivo = sLocalPlan & sFileHTM

    Open sArquivo For Input As #1
    
    Do
    
        Line Input #1, vLinhaArq
        
        If InStr(1, vLinhaArq, "<td rowspan=") = 1 And Not InStr(10, vLinhaArq, "&nbsp") > 1 Then
                
            Call LimpaTag(vLinhaArq)
                 
            aResultados(iColuna, iLinha) = vLinhaArq
                
            'Debug.Print " Array " & iColuna; iLinha; aResultados(iColuna, iLinha)
                
            If iColuna <= 29 Then
                iColuna = iColuna + 1
            Else
                iLinha = iLinha + 1
                iColuna = 0
            End If
    
        End If
           
    Loop While EOF(1) = False
    
    Close #1
    
    wSorteio.Range("a1:AD2000").Value = Application.Transpose(aResultados)
    
    wSorteio.Range("R:AE").ClearContents
    
    dTempo = Now() - dTempo
    
    Debug.Print "Tempo de Arquivo " & dTempo
    
    
End Function
