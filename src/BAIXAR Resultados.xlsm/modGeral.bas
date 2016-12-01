Attribute VB_Name = "modGeral"
Option Explicit
Option Private Module

Public Sub BaixarSorteios()
    
    sLocalPlan = Application.ThisWorkbook.Path
    
    If BaixarArquivoDaNet(sUrlFileZip, sLocalPlan & sFileZip) = False Then
        Call DescompactarArquivo(sLocalPlan & sFileZip, sLocalPlan)
    Else
        MsgBox ("Falha ao baixar arquivo de resultados."), vbCritical, sMsgBoxPaulo
    End If
    
End Sub

Private Function BaixarArquivoDaNet(ByVal strArquivoWeb As String, ByVal strArquivoLocal As String) As Boolean
    
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte
    
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    oXMLHTTP.Open "GET", strArquivoWeb, False
    oXMLHTTP.Send
    
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop
    
    oResp = oXMLHTTP.responseBody
    vFF = FreeFile
    
    If Dir(strArquivoLocal) <> "" Then Kill strArquivoLocal
    
    Open strArquivoLocal For Binary As #vFF
    Put #vFF, , oResp
    Close #vFF
    Set oXMLHTTP = Nothing
    
End Function


Private Sub DescompactarArquivo(strOrigemArquivoCompactado As String, strDestinoArquivoDescompactado As String)
    
    With CreateObject("Shell.Application")
        .Namespace(Error$ & strDestinoArquivoDescompactado).CopyHere .Namespace(Error$ & strOrigemArquivoCompactado).Items
    End With
    
End Sub



