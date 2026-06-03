
Attribute VB_Name = "ModControle"
Option Explicit

' ============================================================
' CONTROLE DE SALDO — Macro Principal
' Versão 1.0
' ============================================================

Public Const TIPOS_CREDITO As String = "CLÍNICAS MÉDICAS|EXCL. DIFAL ICMS BASE PIS/COFINS|EXCL. ICMS BASE PIS/COFINS|EXCL. ICMS ST BASE PIS/COFINS|EXCL. ISS BASE PIS/COFINS|EXCL. PIS/COFINS BASE PIS/COFINS|GROSS UP|ICMS AÇÃO|INSUMOS|IPI ATACADISTA|IPI CONTA GRÁFICA|LC 192 (COMBUSTÍVEIS)|PIS SOBRE FOLHA TERCEIRO SETOR|PREVIDENCIÁRIO TERCEIRO SETOR|REVENDA COM BENEFÍCIOS|SISTEMA S|VERBAS INDENIZATÓRIAS"

' ── Atualizar Painel ─────────────────────────────────────────
Sub AtualizarPainel()
    Dim wsPainel As Worksheet
    Dim ws As Worksheet
    Dim r As Integer
    
    Set wsPainel = Sheets("PAINEL")
    
    ' Limpar linhas de dados (13 a 27)
    Dim i As Integer
    For i = 13 To 27
        Dim c As Integer
        For c = 2 To 8
            wsPainel.Cells(i, c).ClearContents
            wsPainel.Cells(i, c).Interior.Color = IIf(i Mod 2 = 0, RGB(242, 242, 242), RGB(255, 255, 255))
            wsPainel.Cells(i, c).Font.Color = RGB(0, 0, 0)
            wsPainel.Cells(i, c).Font.Underline = xlUnderlineStyleNone
            wsPainel.Cells(i, c).Font.Bold = False
        Next c
    Next i
    
    r = 13

    Dim sufixo As String
    Dim tipoCred As String
    Dim regime As String
    Dim per1 As String
    Dim per2 As String
    Dim valorAp As Double
    Dim saldo As Double
    Dim status As String

    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " Then
            If r > 27 Then Exit For

            sufixo = Mid(ws.Name, 6)
            
            tipoCred = ws.Range("B13").Value
            regime = ws.Range("A1").Value
            per1 = ws.Range("G13").Value
            per2 = ws.Range("H13").Value
            If per1 = "" Then per1 = "?"
            If per2 = "" Then per2 = "?"
            valorAp = ws.Range("E15").Value
            
            ' Buscar saldo da Conta Corrente
            Dim ccName As String
            ccName = "CC - " & sufixo
            saldo = 0
            
            If SheetExists(ccName) Then
                Dim wsCC As Worksheet
                Set wsCC = Sheets(ccName)
                If regime = "CUMULATIVO" Then
                    Dim lastRow As Long
                    lastRow = wsCC.Cells(wsCC.Rows.Count, 8).End(xlUp).Row
                    If lastRow >= 8 Then
                        saldo = wsCC.Cells(lastRow, 8).Value
                    Else
                        saldo = wsCC.Range("E7").Value
                    End If
                Else
                    ' Não cumulativo: saldo restante em Q41
                    Dim vQ41 As Variant
                    vQ41 = wsCC.Range("Q41").Value
                    If IsError(vQ41) Or IsEmpty(vQ41) Then
                        saldo = wsCC.Range("E7").Value
                    Else
                        saldo = CDbl(vQ41)
                    End If
                End If
            End If
            
            ' Status
            If saldo <= 0 Then
                status = "ZERADO"
            ElseIf valorAp > 0 And saldo < valorAp * 0.15 Then
                status = "QUASE ZERADO"
            Else
                status = "EM USO"
            End If
            
            ' Preencher linha
            With wsPainel
                .Cells(r, 2).Value = tipoCred
                .Cells(r, 3).Value = regime
                .Cells(r, 4).Value = per1 & " a " & per2
                
                .Cells(r, 5).Value = valorAp
                .Cells(r, 5).NumberFormat = "R$ #,##0.00"
                
                .Cells(r, 6).Value = saldo
                .Cells(r, 6).NumberFormat = "R$ #,##0.00"
                
                .Cells(r, 7).Value = status
                Select Case status
                    Case "EM USO"
                        .Cells(r, 7).Interior.Color = RGB(198, 224, 180)
                        .Cells(r, 7).Font.Color = RGB(55, 86, 35)
                    Case "QUASE ZERADO"
                        .Cells(r, 7).Interior.Color = RGB(255, 235, 156)
                        .Cells(r, 7).Font.Color = RGB(124, 101, 0)
                    Case "ZERADO"
                        .Cells(r, 7).Interior.Color = RGB(255, 199, 206)
                        .Cells(r, 7).Font.Color = RGB(156, 0, 6)
                End Select
                .Cells(r, 7).Font.Bold = True
                
                ' Link "Abrir"
                .Cells(r, 8).Value = "→ Abrir"
                .Cells(r, 8).Font.Color = RGB(31, 56, 100)
                .Cells(r, 8).Font.Underline = xlUnderlineStyleSingle
                .Cells(r, 8).Font.Bold = True
            End With
            
            r = r + 1
        End If
    Next ws
    
    Sheets("PAINEL").Activate
    MsgBox "Painel atualizado com " & (r - 13) & " crédito(s).", vbInformation, "Atualizado"
End Sub

' ── Abrir crédito ao clicar na linha ─────────────────────────
Sub AbrirCreditoSelecionado()
    Dim wsPainel As Worksheet
    Dim r As Long
    
    Set wsPainel = Sheets("PAINEL")
    r = ActiveCell.Row
    
    If r < 13 Or r > 27 Then
        MsgBox "Clique em uma linha de crédito da tabela (linhas 13 a 27).", vbInformation
        Exit Sub
    End If
    
    If wsPainel.Cells(r, 2).Value = "" Then
        MsgBox "Esta linha não contém um crédito.", vbInformation
        Exit Sub
    End If
    
    Dim tipoCred As String
    tipoCred = wsPainel.Cells(r, 2).Value
    Dim sufixo As String
    sufixo = SufixoCurto(tipoCred)
    
    ' Tentar encontrar a aba QD correspondente
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " And InStr(ws.Name, sufixo) > 0 Then
            ws.Activate
            Exit Sub
        End If
    Next ws
    
    MsgBox "Aba não encontrada para: " & tipoCred, vbExclamation
End Sub

' ── Novo Crédito ─────────────────────────────────────────────
Sub NovoCredito()
    ' Verificar se dados do cliente foram preenchidos
    If Sheets("PAINEL").Range("C4").Value = "" Then
        MsgBox "Preencha a Razão Social do cliente antes de criar créditos.", vbExclamation
        Sheets("PAINEL").Range("C4").Select
        Exit Sub
    End If
    
    Dim tipoCred As String
    Dim regime As String
    Dim per1 As String
    Dim per2 As String
    Dim valorAp As Double
    
    ' Selecionar Tipo de Crédito
    tipoCred = EscolherTipoCred("Selecione o Tipo de Crédito:")
    If tipoCred = "" Then Exit Sub
    
    ' Selecionar Regime
    Dim resp As Integer
    resp = MsgBox("Qual o regime deste crédito?" & vbNewLine & vbNewLine & _
                  "SIM = CUMULATIVO" & vbNewLine & "NÃO = NÃO CUMULATIVO", _
                  vbYesNo + vbQuestion, "Regime Tributário")
    regime = IIf(resp = vbYes, "CUMULATIVO", "NÃO CUMULATIVO")
    
    ' Período início
    per1 = InputBox("Período de início do cálculo:" & vbNewLine & "(Ex: 01/2020)", "Período Inicial")
    If per1 = "" Then Exit Sub
    
    ' Período fim
    per2 = InputBox("Período de fim do cálculo:" & vbNewLine & "(Ex: 12/2022)", "Período Final")
    If per2 = "" Then Exit Sub
    
    ' Valor apurado
    Dim strValor As String
    strValor = InputBox("Valor total apurado (R$):" & vbNewLine & "(Ex: 80872.06)", "Valor Apurado")
    If strValor = "" Then Exit Sub
    strValor = Replace(strValor, ",", ".")
    If Not IsNumeric(strValor) Then
        MsgBox "Valor inválido.", vbExclamation: Exit Sub
    End If
    valorAp = CDbl(strValor)
    
    ' Confirmar
    Dim msg As String
    msg = "Criar crédito com os dados abaixo?" & vbNewLine & vbNewLine & _
          "Tipo: " & tipoCred & vbNewLine & _
          "Regime: " & regime & vbNewLine & _
          "Período: " & per1 & " a " & per2 & vbNewLine & _
          "Valor: R$ " & Format(valorAp, "#,##0.00")
    
    If MsgBox(msg, vbYesNo + vbQuestion, "Confirmar") = vbNo Then Exit Sub
    
    CriarAbas tipoCred, regime, per1, per2, valorAp, False
End Sub

' ── Complemento de Crédito ───────────────────────────────────
Sub ComplementoCredito()
    ' Listar créditos existentes
    Dim creditos() As String
    Dim count As Integer
    count = 0
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " Then count = count + 1
    Next ws
    
    If count = 0 Then
        MsgBox "Nenhum crédito encontrado. Crie um crédito primeiro.", vbInformation
        Exit Sub
    End If
    
    ReDim creditos(count - 1)
    Dim idx As Integer
    idx = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " Then
            creditos(idx) = Mid(ws.Name, 6)
            idx = idx + 1
        End If
    Next ws
    
    ' Mostrar lista para escolha
    Dim lista As String
    Dim n As Integer
    For n = 0 To count - 1
        lista = lista & (n + 1) & ". " & creditos(n) & vbNewLine
    Next n
    
    Dim escolha As String
    escolha = InputBox("Escolha o número do crédito para adicionar complemento:" & vbNewLine & vbNewLine & lista, "Complemento de Crédito")
    If escolha = "" Then Exit Sub
    If Not IsNumeric(escolha) Then MsgBox "Opção inválida.": Exit Sub
    
    Dim numEscolha As Integer
    numEscolha = CInt(escolha)
    If numEscolha < 1 Or numEscolha > count Then MsgBox "Opção inválida.": Exit Sub
    
    Dim baseNome As String
    baseNome = creditos(numEscolha - 1)
    
    ' Dados do complemento
    Dim per1 As String
    Dim per2 As String
    Dim valorComp As Double
    
    per1 = InputBox("Período de início do complemento (Ex: 01/2023):", "Período Inicial")
    If per1 = "" Then Exit Sub
    
    per2 = InputBox("Período de fim do complemento (Ex: 12/2023):", "Período Final")
    If per2 = "" Then Exit Sub
    
    Dim strValor As String
    strValor = InputBox("Valor do complemento (R$):", "Valor")
    If strValor = "" Then Exit Sub
    strValor = Replace(strValor, ",", ".")
    If Not IsNumeric(strValor) Then MsgBox "Valor inválido.": Exit Sub
    valorComp = CDbl(strValor)
    
    ' Buscar tipo e regime do crédito base
    Dim wsBase As Worksheet
    Set wsBase = Sheets("QD - " & baseNome)
    Dim tipoCred As String
    Dim regime As String
    tipoCred = wsBase.Range("B13").Value
    regime = wsBase.Range("A1").Value
    
    ' Gerar sufixo com número de complemento
    Dim compNum As Integer
    compNum = ContarAbasPorTipo(SufixoCurto(tipoCred)) + 1
    
    Dim sufixoComp As String
    sufixoComp = SufixoCurto(tipoCred) & " (" & compNum & ")"
    
    ' Confirmar
    If MsgBox("Criar complemento '" & sufixoComp & "'?" & vbNewLine & "Valor: R$ " & Format(valorComp, "#,##0.00"), _
              vbYesNo + vbQuestion, "Confirmar") = vbNo Then Exit Sub
    
    CriarAbasSufixo sufixoComp, tipoCred, regime, per1, per2, valorComp
End Sub

' ── Criar abas (entry point) ─────────────────────────────────
Sub CriarAbas(tipoCred As String, regime As String, per1 As String, per2 As String, valorAp As Double, isComp As Boolean)
    Dim sufixo As String
    sufixo = SufixoCurto(tipoCred)
    
    If Not isComp And SheetExists("QD - " & sufixo) Then
        Dim num As Integer
        num = 2
        Do While SheetExists("QD - " & sufixo & " (" & num & ")")
            num = num + 1
        Loop
        sufixo = sufixo & " (" & num & ")"
    End If
    
    CriarAbasSufixo sufixo, tipoCred, regime, per1, per2, valorAp
End Sub

Sub CriarAbasSufixo(sufixo As String, tipoCred As String, regime As String, per1 As String, per2 As String, valorAp As Double)
    Application.ScreenUpdating = False
    
    Dim wsPainel As Worksheet
    Set wsPainel = Sheets("PAINEL")
    Dim razao As String
    Dim nomeFant As String
    Dim cnpj As String
    razao = wsPainel.Range("C4").Value
    nomeFant = wsPainel.Range("C5").Value
    cnpj = wsPainel.Range("C6").Value
    
    ' Criar QD
    Dim wsQD As Worksheet
    Dim tplQD As String
    tplQD = IIf(regime = "CUMULATIVO", "_TPL_QD_CUM", "_TPL_QD_NCUM")
    Sheets(tplQD).Copy After:=Sheets(Sheets.Count)
    Set wsQD = ActiveSheet
    wsQD.Name = "QD - " & sufixo
    wsQD.Tab.Color = IIf(regime = "CUMULATIVO", RGB(68, 114, 196), RGB(112, 173, 71))
    wsQD.Visible = xlSheetVisible
    
    ' Preencher QD
    wsQD.Range("D5").Value = razao
    wsQD.Range("D6").Value = nomeFant
    wsQD.Range("D7").Value = cnpj
    wsQD.Range("B13").Value = tipoCred
    wsQD.Range("E13").Value = valorAp
    ' Período como texto
    wsQD.Range("G13").Value = per1
    wsQD.Range("H13").Value = per2
    ' Guardar regime em A1 (oculto)
    wsQD.Range("A1").Value = regime
    wsQD.Range("A1").Font.Color = RGB(255, 255, 255)
    
    ' Botão voltar
    AdicionarBotaoVoltar wsQD
    
    ' Criar CC
    Dim wsCC As Worksheet
    Dim tplCC As String
    tplCC = IIf(regime = "CUMULATIVO", "_TPL_CC_CUM", "_TPL_CC_NCUM")
    Sheets(tplCC).Copy After:=Sheets(Sheets.Count)
    Set wsCC = ActiveSheet
    wsCC.Name = "CC - " & sufixo
    wsCC.Tab.Color = IIf(regime = "CUMULATIVO", RGB(68, 114, 196), RGB(112, 173, 71))
    wsCC.Visible = xlSheetVisible
    wsCC.Range("E7").Value = valorAp
    wsCC.Range("E7").Formula = "='QD - " & sufixo & "'!E15"
    AdicionarBotaoVoltar wsCC
    
    ' Criar PERDCOMP e LG somente para cumulativo
    If regime = "CUMULATIVO" Then
        Sheets("_TPL_PC_CUM").Copy After:=Sheets(Sheets.Count)
        Dim wsPC As Worksheet
        Set wsPC = ActiveSheet
        wsPC.Name = "PC - " & sufixo
        wsPC.Tab.Color = RGB(68, 114, 196)
        wsPC.Visible = xlSheetVisible
        AdicionarBotaoVoltar wsPC
        
        Sheets("_TPL_LG").Copy After:=Sheets(Sheets.Count)
        Dim wsLG As Worksheet
        Set wsLG = ActiveSheet
        wsLG.Name = "LG - " & sufixo
        wsLG.Tab.Color = RGB(68, 114, 196)
        wsLG.Visible = xlSheetVisible
        AdicionarBotaoVoltar wsLG
    End If
    
    Application.ScreenUpdating = True
    
    ' Atualizar painel e navegar
    AtualizarPainelSilencioso
    wsQD.Activate
    
    MsgBox "Crédito '" & sufixo & "' criado com sucesso!" & vbNewLine & _
           "Regime: " & regime & vbNewLine & _
           "Valor: R$ " & Format(valorAp, "#,##0.00"), vbInformation, "Crédito Criado"
End Sub

' ── Adicionar botão Voltar ao Painel ─────────────────────────
Sub AdicionarBotaoVoltar(ws As Worksheet)
    On Error Resume Next
    Dim btn As Button
    Set btn = ws.Buttons.Add(5, 5, 90, 20)
    btn.Caption = "← PAINEL"
    btn.OnAction = "VoltarPainel"
    btn.Font.Size = 9
    btn.Font.Bold = True
    On Error GoTo 0
End Sub

' ── Voltar ao Painel ─────────────────────────────────────────
Sub VoltarPainel()
    AtualizarPainelSilencioso
    Sheets("PAINEL").Activate
End Sub

' ── Escolher tipo de crédito (lista numerada) ────────────────
Function EscolherTipoCred(titulo As String) As String
    Dim tipos() As String
    tipos = Split(TIPOS_CREDITO, "|")
    
    Dim lista As String
    Dim n As Integer
    For n = 0 To UBound(tipos)
        lista = lista & (n + 1) & ". " & tipos(n) & vbNewLine
    Next n
    
    Dim escolha As String
    escolha = InputBox(titulo & vbNewLine & vbNewLine & lista, "Tipo de Crédito")
    
    If escolha = "" Then
        EscolherTipoCred = ""
        Exit Function
    End If
    
    If Not IsNumeric(escolha) Then
        MsgBox "Digite apenas o número da opção.", vbExclamation
        EscolherTipoCred = ""
        Exit Function
    End If
    
    Dim num As Integer
    num = CInt(escolha)
    If num < 1 Or num > UBound(tipos) + 1 Then
        MsgBox "Opção inválida.", vbExclamation
        EscolherTipoCred = ""
        Exit Function
    End If
    
    EscolherTipoCred = tipos(num - 1)
End Function

' ── Atualizar painel sem mensagem ────────────────────────────
Sub AtualizarPainelSilencioso()
    Dim wsPainel As Worksheet
    Dim ws As Worksheet
    Dim r As Integer
    
    Set wsPainel = Sheets("PAINEL")
    
    Dim i As Integer
    For i = 13 To 27
        Dim c As Integer
        For c = 2 To 8
            wsPainel.Cells(i, c).ClearContents
            wsPainel.Cells(i, c).Interior.Color = IIf(i Mod 2 = 0, RGB(242, 242, 242), RGB(255, 255, 255))
            wsPainel.Cells(i, c).Font.Color = RGB(0, 0, 0)
            wsPainel.Cells(i, c).Font.Underline = xlUnderlineStyleNone
            wsPainel.Cells(i, c).Font.Bold = False
        Next c
    Next i
    
    r = 13
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " Then
            If r > 27 Then Exit For
            Dim sufixo As String
            sufixo = Mid(ws.Name, 6)
            Dim tipoCred As String
            Dim regime As String
            Dim valorAp As Double
            Dim saldo As Double
            Dim status As String
            
            tipoCred = ws.Range("B13").Value
            regime = ws.Range("A1").Value
            valorAp = ws.Range("E15").Value
            
            Dim per1 As String, per2 As String
            per1 = ws.Range("G13").Value
            per2 = ws.Range("H13").Value
            
            saldo = 0
            Dim ccName As String
            ccName = "CC - " & sufixo
            If SheetExists(ccName) Then
                Dim wsCC As Worksheet
                Set wsCC = Sheets(ccName)
                If regime = "CUMULATIVO" Then
                    Dim lastRow As Long
                    lastRow = wsCC.Cells(wsCC.Rows.Count, 8).End(xlUp).Row
                    If lastRow >= 8 Then saldo = wsCC.Cells(lastRow, 8).Value _
                    Else saldo = wsCC.Range("E7").Value
                Else
                    On Error Resume Next
                    saldo = wsCC.Range("Q41").Value
                    If saldo = 0 Then saldo = wsCC.Range("E7").Value
                    On Error GoTo 0
                End If
            End If
            
            If saldo <= 0 Then
                status = "ZERADO"
            ElseIf valorAp > 0 And saldo < valorAp * 0.15 Then
                status = "QUASE ZERADO"
            Else
                status = "EM USO"
            End If
            
            With wsPainel
                .Cells(r, 2).Value = tipoCred
                .Cells(r, 3).Value = regime
                .Cells(r, 4).Value = per1 & " a " & per2
                .Cells(r, 5).Value = valorAp
                .Cells(r, 5).NumberFormat = "R$ #,##0.00"
                .Cells(r, 6).Value = saldo
                .Cells(r, 6).NumberFormat = "R$ #,##0.00"
                .Cells(r, 7).Value = status
                Select Case status
                    Case "EM USO": .Cells(r, 7).Interior.Color = RGB(198, 224, 180): .Cells(r, 7).Font.Color = RGB(55, 86, 35)
                    Case "QUASE ZERADO": .Cells(r, 7).Interior.Color = RGB(255, 235, 156): .Cells(r, 7).Font.Color = RGB(124, 101, 0)
                    Case "ZERADO": .Cells(r, 7).Interior.Color = RGB(255, 199, 206): .Cells(r, 7).Font.Color = RGB(156, 0, 6)
                End Select
                .Cells(r, 7).Font.Bold = True
                .Cells(r, 8).Value = "→ Abrir"
                .Cells(r, 8).Font.Color = RGB(31, 56, 100)
                .Cells(r, 8).Font.Underline = xlUnderlineStyleSingle
                .Cells(r, 8).Font.Bold = True
            End With
            r = r + 1
        End If
    Next ws
End Sub

' ── Utilitários ──────────────────────────────────────────────
Function SheetExists(nome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(nome)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function SufixoCurto(tipoCred As String) As String
    Dim s As String
    s = tipoCred
    s = Replace(s, "EXCL. ", "")
    s = Replace(s, " BASE PIS/COFINS", "")
    s = Replace(s, "TERCEIRO SETOR", "3 SETOR")
    s = Replace(s, "COMBUSTÍVEIS", "COMB.")
    If Len(s) > 22 Then s = Left(s, 22)
    SufixoCurto = Trim(s)
End Function

Function ContarAbasPorTipo(sufixo As String) As Integer
    Dim ws As Worksheet
    Dim count As Integer
    count = 0
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 5) = "QD - " And InStr(1, ws.Name, sufixo, vbTextCompare) > 0 Then
            count = count + 1
        End If
    Next ws
    ContarAbasPorTipo = count
End Function
