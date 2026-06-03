Attribute VB_Name = "CriarBotoes"
Sub CriarBotoesPainel()
    Dim ws As Worksheet
    Set ws = Sheets("PAINEL")
    
    ' Remover botões antigos se existirem
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    ' Limpar conteúdo das células de botão
    ws.Range("B29:C29").ClearContents
    ws.Range("D29:E29").ClearContents
    ws.Range("F29:G29").ClearContents
    
    ' Pegar posição e tamanho das células
    Dim rngNovo As Range
    Dim rngComp As Range
    Dim rngAtual As Range
    
    Set rngNovo = ws.Range("B29:C29")
    Set rngComp = ws.Range("D29:E29")
    Set rngAtual = ws.Range("F29:G29")
    
    ' Botão NOVO CRÉDITO
    Dim btn1 As Button
    Set btn1 = ws.Buttons.Add(rngNovo.Left, rngNovo.Top, rngNovo.Width, rngNovo.Height)
    With btn1
        .Caption = Chr(10) & Chr(10) & "  + NOVO CR" & Chr(201) & "DITO"
        .OnAction = "NovoCredito"
        .Font.Size = 10
        .Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
    End With
    ' Colorir fundo azul via shape
    btn1.ShapeRange.Fill.ForeColor.RGB = RGB(46, 117, 182)
    btn1.ShapeRange.Line.Visible = msoFalse
    
    ' Botão COMPLEMENTO
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(rngComp.Left, rngComp.Top, rngComp.Width, rngComp.Height)
    With btn2
        .Caption = "  COMPLEMENTO"
        .OnAction = "ComplementoCredito"
        .Font.Size = 10
        .Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
    End With
    btn2.ShapeRange.Fill.ForeColor.RGB = RGB(112, 173, 71)
    btn2.ShapeRange.Line.Visible = msoFalse
    
    ' Botão ATUALIZAR PAINEL
    Dim btn3 As Button
    Set btn3 = ws.Buttons.Add(rngAtual.Left, rngAtual.Top, rngAtual.Width, rngAtual.Height)
    With btn3
        .Caption = "  ATUALIZAR PAINEL"
        .OnAction = "AtualizarPainel"
        .Font.Size = 10
        .Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
    End With
    btn3.ShapeRange.Fill.ForeColor.RGB = RGB(68, 114, 196)
    btn3.ShapeRange.Line.Visible = msoFalse
    
    ' Botões de ABRIR para cada linha de credito (col H, linhas 13-27)
    Dim i As Integer
    For i = 13 To 27
        Dim rngAbrir As Range
        Set rngAbrir = ws.Range("H" & i)
        Dim btnAbrir As Button
        Set btnAbrir = ws.Buttons.Add(rngAbrir.Left, rngAbrir.Top, rngAbrir.Width, rngAbrir.Height)
        With btnAbrir
            .Caption = Chr(8594) & " Abrir"
            .OnAction = "AbrirLinha" & i
            .Font.Size = 9
            .Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
        End With
        btnAbrir.ShapeRange.Fill.ForeColor.RGB = RGB(31, 56, 100)
        btnAbrir.ShapeRange.Line.Visible = msoFalse
    Next i
    
    MsgBox "Botões criados com sucesso! Salve o arquivo.", vbInformation
End Sub

' Subs individuais para cada linha (necessario para botoes)
Sub AbrirLinha13(): AbrirLinhaPainel 13: End Sub
Sub AbrirLinha14(): AbrirLinhaPainel 14: End Sub
Sub AbrirLinha15(): AbrirLinhaPainel 15: End Sub
Sub AbrirLinha16(): AbrirLinhaPainel 16: End Sub
Sub AbrirLinha17(): AbrirLinhaPainel 17: End Sub
Sub AbrirLinha18(): AbrirLinhaPainel 18: End Sub
Sub AbrirLinha19(): AbrirLinhaPainel 19: End Sub
Sub AbrirLinha20(): AbrirLinhaPainel 20: End Sub
Sub AbrirLinha21(): AbrirLinhaPainel 21: End Sub
Sub AbrirLinha22(): AbrirLinhaPainel 22: End Sub
Sub AbrirLinha23(): AbrirLinhaPainel 23: End Sub
Sub AbrirLinha24(): AbrirLinhaPainel 24: End Sub
Sub AbrirLinha25(): AbrirLinhaPainel 25: End Sub
Sub AbrirLinha26(): AbrirLinhaPainel 26: End Sub
Sub AbrirLinha27(): AbrirLinhaPainel 27: End Sub

Sub AbrirLinhaPainel(r As Integer)
    Dim wsPainel As Worksheet
    Set wsPainel = Sheets("PAINEL")
    
    If wsPainel.Cells(r, 2).Value = "" Then
        MsgBox "Esta linha não contém um crédito.", vbInformation
        Exit Sub
    End If
    
    Dim tipoCred As String
    tipoCred = wsPainel.Cells(r, 2).Value
    Dim sufixo As String
    sufixo = SufixoCurto(tipoCred)
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " And InStr(1, ws.Name, sufixo, vbTextCompare) > 0 Then
            ws.Activate
            Exit Sub
        End If
    Next ws
    
    ' Tentar pelo numero da linha (complementos)
    Dim count As Integer
    count = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " Then
            count = count + 1
            If count = (r - 12) Then
                ws.Activate
                Exit Sub
            End If
        End If
    Next ws
    
    MsgBox "Aba não encontrada para: " & tipoCred, vbExclamation
End Sub
