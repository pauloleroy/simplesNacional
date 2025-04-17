Attribute VB_Name = "Módulo1"
Option Explicit

Sub ProcessarMultiplasNotasFiscaisParaTabela()
    Dim xmlDoc As Object
    Dim fileDialog As fileDialog
    Dim filePath As Variant
    Dim xmlNamespace As String
    Dim ws As Worksheet
    Dim ultimaLinha As ListRow
    Dim listaTabela As ListObject
    Dim numeroNF As String, dataEmissao As String, prestadora As String, dataEmissaoConfig As Date
    Dim cfop As String, issRetido As String, valorISS As String
    Dim valorTotal As String, valorDevolucao As String, cancelada As String
    
    ' Configura a planilha e a tabela
    Set ws = ThisWorkbook.Sheets("Lancamentos")
    Set listaTabela = ws.ListObjects("TabelaDados") ' Substitua "TabelaDados" pelo nome da sua tabela estruturada
    
    ' Configura o seletor de arquivos
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Selecione os arquivos XML"
        .Filters.Clear
        .Filters.Add "Arquivos XML", "*.xml"
        .AllowMultiSelect = True
        
        If .Show = False Then Exit Sub ' Se o usuário cancelar
    End With
    
    ' Processa cada arquivo selecionado
    For Each filePath In fileDialog.SelectedItems
        ' Carrega o XML
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        xmlDoc.Load (filePath)
        xmlDoc.SetProperty "SelectionNamespaces", "xmlns:nfe='http://www.portalfiscal.inf.br/nfe' xmlns:nfse='http://www.abrasf.org.br/nfse.xsd'"
        
        ' Identifica o tipo de nota
        If Not xmlDoc.SelectSingleNode("//nfe:infNFe") Is Nothing Then
            xmlNamespace = "nfe"
        ElseIf Not xmlDoc.SelectSingleNode("//nfse:InfNfse") Is Nothing Then
            xmlNamespace = "nfse"
        Else
            MsgBox "Tipo de nota fiscal não reconhecido no arquivo: " & filePath, vbExclamation
            GoTo ProximoArquivo
        End If
        
        ' Extrai dados comuns
        Select Case xmlNamespace
            Case "nfe"
                numeroNF = xmlDoc.SelectSingleNode("//nfe:ide/nfe:nNF").Text
                dataEmissao = xmlDoc.SelectSingleNode("//nfe:ide/nfe:dhEmi").Text
                dataEmissaoConfig = AjustarFormatoData(Mid(dataEmissao, 1, 10))  ' Formata a data
                prestadora = xmlDoc.SelectSingleNode("//nfe:emit/nfe:CNPJ").Text
                cfop = xmlDoc.SelectSingleNode("//nfe:det/nfe:prod/nfe:CFOP").Text
                valorTotal = xmlDoc.SelectSingleNode("//nfe:total/nfe:ICMSTot/nfe:vNF").Text
                valorISS = "0" ' NF-e geralmente não tem ISS Retido
                cancelada = "Não" ' NF-e não possui cancelamento nesta análise
                If cfop = "1202" Then
                    valorDevolucao = valorTotal
                    valorTotal = "" ' Zera o valor total porque é devolução
                Else
                    valorDevolucao = "" ' Zera devolução porque não é o caso
                End If
                
                ' Verifica se o CFOP é 1202 ou 5102 antes de adicionar na tabela
                If cfop = "1202" Or cfop = "5102" Or cfop = "6102" Or cfop = "5101" Or cfop = "6101" Then
                    ' Adiciona os dados na tabela
                    Set ultimaLinha = listaTabela.ListRows.Add
                    With ultimaLinha
                        .Range(1).Value = numeroNF
                        .Range(2).Value = prestadora
                        .Range(3).Value = dataEmissaoConfig
                        .Range(4).Value = valorTotal
                        .Range(5).Value = valorISS ' Valor do ISS Retido
                        .Range(6).Value = valorDevolucao
                        .Range(7).Value = cancelada ' Indicação de cancelamento
                    End With
                End If
                
            Case "nfse"
                numeroNF = Right(xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:Numero").Text, 4)
                dataEmissao = xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:DataEmissao").Text
                dataEmissaoConfig = AjustarFormatoData(Mid(dataEmissao, 1, 10))  ' Formata a data
                prestadora = xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:PrestadorServico/nfse:IdentificacaoPrestador/nfse:Cnpj").Text
                valorTotal = xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:Servico/nfse:Valores/nfse:ValorServicos").Text
                If xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:Servico/nfse:Valores/nfse:IssRetido").Text = "1" Then
                    valorISS = xmlDoc.SelectSingleNode("//nfse:InfNfse/nfse:Servico/nfse:Valores/nfse:ValorIss").Text
                Else
                    valorISS = "0"
                End If
                valorDevolucao = "" ' NFSe não lida com CFOP de devolução
                cancelada = IIf(Not xmlDoc.SelectSingleNode("//nfse:NfseCancelamento") Is Nothing, "Sim", "Não")
                
                ' Adiciona os dados na tabela
                Set ultimaLinha = listaTabela.ListRows.Add
                With ultimaLinha
                    .Range(1).Value = numeroNF
                    .Range(2).Value = prestadora
                    .Range(3).Value = dataEmissaoConfig
                    .Range(4).Value = valorTotal
                    .Range(5).Value = valorISS ' Valor do ISS Retido
                    .Range(6).Value = valorDevolucao
                    .Range(7).Value = cancelada ' Indicação de cancelamento
                End With
        End Select
        
ProximoArquivo:
    Next filePath
    
    MsgBox "Dados extraídos com sucesso!", vbInformation
End Sub


Function AjustarFormatoData(dataISO As String) As Date
    Dim partes() As String
    partes = Split(dataISO, "-")
    AjustarFormatoData = DateSerial(partes(0), partes(1), partes(2))
End Function


Sub ConsolidarDados()
    Dim wsOrigem As Worksheet, wsDestino As Worksheet, wsDestino2 As Worksheet
    Dim tabelaOrigem As ListObject, tabelaDestino As ListObject, tabelaDestino2 As ListObject
    Dim dictFilial As Object, dictMensal As Object
    Dim linha As ListRow, keyFilial As Variant, keyMensal As Variant
    Dim mesAno As String, cnpj As String
    Dim valoresFilial() As Variant, valoresMensal() As Variant
    Dim cancelada As String, nfsExistentes As Collection, nfFaltantes As String
    Dim nfAtual As Long, nfPrimeira As String, nfUltima As String
    Dim totalMesBruto As Double, totalMesLiquido As Double
    Dim nfEncontrada As Boolean
    
    ' Define as planilhas e tabelas
    Set wsOrigem = ThisWorkbook.Sheets("Lancamentos") ' Planilha com os dados originais
    Set wsDestino = ThisWorkbook.Sheets("Resumo") ' Planilha para o consolidado
    Set wsDestino2 = ThisWorkbook.Sheets("Calc_Simples") ' Planilha mensal consolidada
    Set tabelaOrigem = wsOrigem.ListObjects("TabelaDados") ' Tabela de origem
    Set tabelaDestino = wsDestino.ListObjects("TabelaConsolidada") ' Tabela destino por filial
    Set tabelaDestino2 = wsDestino2.ListObjects("TabelaMensal") ' Tabela mensal consolidada
    
    ' Remove todas as linhas das tabelas de destino
    On Error Resume Next
    Do While tabelaDestino.ListRows.Count > 0
        tabelaDestino.ListRows(1).Delete
    Loop
    Do While tabelaDestino2.ListRows.Count > 0
        tabelaDestino2.ListRows(1).Delete
    Loop
    On Error GoTo 0
    
    ' Configura os dicionários para consolidação
    Set dictFilial = CreateObject("Scripting.Dictionary")
    Set dictMensal = CreateObject("Scripting.Dictionary")
    
    ' Processa os dados da tabela original
    For Each linha In tabelaOrigem.ListRows
        ' Obtém os valores relevantes
        mesAno = Format(linha.Range(3).Value, "mm/yyyy")
        cnpj = linha.Range(2).Value
        keyFilial = mesAno & "|" & cnpj ' Chave única para mês e CNPJ
        keyMensal = mesAno ' Chave única apenas para o mês
        nfFaltantes = linha.Range(8).Value ' Coluna para NF_Faltantes
        nfAtual = linha.Range(9).Value ' Número da NF atual
        
        ' Inicializa os dados no dicionário por filial
        If Not dictFilial.exists(keyFilial) Then
            valoresFilial = Array(0, 0, 0, 0, 0, 0, 0, New Collection) ' COM_RETENCAO, SEM_RETENCAO, FAT_BRUTO, DEVOLUCAO, FAT_LIQUIDO, PRIMEIRA_NF, ULTIMA_NF, NFS_EXISTENTES
            dictFilial.Add keyFilial, valoresFilial
        End If
        
        ' Adiciona a NF à coleção de notas existentes
        On Error Resume Next
        valoresFilial(7).Add linha.Range(1).Value, CStr(linha.Range(1).Value)
        On Error GoTo 0
        
        ' Inicializa os dados no dicionário mensal
        If Not dictMensal.exists(keyMensal) Then
            valoresMensal = Array(0, 0, 0, 0) ' FAT_BRUTO, FAT_LIQUIDO, ISS_RET, DEVOLUCAO
            dictMensal.Add keyMensal, valoresMensal
        End If
        
        
        valoresFilial = dictFilial(keyFilial)
        valoresMensal = dictMensal(keyMensal)
        
        ' Determina a primeira e a última NF considerando todas as notas (incluindo canceladas)
        If valoresFilial(5) = 0 Or linha.Range(1).Value < valoresFilial(5) Then
            valoresFilial(5) = linha.Range(1).Value ' PRIMEIRA_NF
        End If
        If valoresFilial(6) = 0 Or linha.Range(1).Value > valoresFilial(6) Then
            valoresFilial(6) = linha.Range(1).Value ' ULTIMA_NF
        End If
        
        ' Ignora os cálculos financeiros se a nota for cancelada
        cancelada = linha.Range(7).Value
        If cancelada = "Sim" Then GoTo ProximoRegistro
        
        ' Atualiza os totais por filial
        If linha.Range(5).Value > 0 Then
            valoresFilial(0) = valoresFilial(0) + linha.Range(4).Value ' COM_RETENCAO
        Else
            valoresFilial(1) = valoresFilial(1) + linha.Range(4).Value ' SEM_RETENCAO
        End If
        valoresFilial(2) = valoresFilial(2) + linha.Range(4).Value ' FAT_BRUTO
        valoresFilial(3) = valoresFilial(3) + linha.Range(6).Value ' DEVOLUCAO
        valoresFilial(4) = valoresFilial(2) - valoresFilial(3) ' FAT_LIQUIDO
        
        
        ' Atualiza os totais mensais
        valoresMensal(0) = valoresMensal(0) + linha.Range(4).Value ' FAT_BRUTO
        valoresMensal(2) = valoresMensal(2) + linha.Range(5).Value ' ISS_RET
        valoresMensal(3) = valoresMensal(3) + linha.Range(6).Value ' Acumula devoluções
        valoresMensal(1) = valoresMensal(0) - valoresMensal(3) ' FAT_LIQUIDO = FAT_BRUTO - DEVOLUÇÕES
        
ProximoRegistro:
        dictFilial(keyFilial) = valoresFilial
        dictMensal(keyMensal) = valoresMensal
    Next linha
    
    ' Insere os dados consolidados por filial na tabela "TabelaConsolidada"
    For Each keyFilial In dictFilial.Keys
        valoresFilial = dictFilial(keyFilial)
        mesAno = Split(keyFilial, "|")(0)
        cnpj = Split(keyFilial, "|")(1)
        
        ' Verifica notas faltantes no intervalo
        Set nfsExistentes = valoresFilial(7)
        nfFaltantes = ""
        For nfAtual = valoresFilial(5) To valoresFilial(6)
            nfEncontrada = False
            On Error Resume Next
            nfEncontrada = Not IsEmpty(nfsExistentes(CStr(nfAtual)))
            On Error GoTo 0
            If Not nfEncontrada Then
                nfFaltantes = nfFaltantes & nfAtual & ", "
            End If
        Next nfAtual
        If nfFaltantes <> "" Then nfFaltantes = Left(nfFaltantes, Len(nfFaltantes) - 2)
        
        With tabelaDestino.ListRows.Add
            .Range(1).Value = mesAno
            .Range(2).Value = cnpj
            .Range(3).Value = valoresFilial(0) ' COM_RETENCAO
            .Range(4).Value = valoresFilial(1) ' SEM_RETENCAO
            .Range(5).Value = valoresFilial(2) ' FAT_BRUTO
            .Range(6).Value = valoresFilial(3) ' DEVOLUCAO
            .Range(7).Value = valoresFilial(4) ' FAT_LIQUIDO
            .Range(8).Value = "'" & valoresFilial(5) & " - " & valoresFilial(6) ' NF (como texto para evitar interpretação como data)
            .Range(9).Value = nfFaltantes ' NF_FALTANTES
        End With
    Next keyFilial
    
    ' Insere os dados consolidados mensais na tabela "TabelaMensal"
    For Each keyMensal In dictMensal.Keys
        valoresMensal = dictMensal(keyMensal)
        
        
        
         ' Busca o valor correspondente à TabelaGuia
        Dim valorGuia As Variant
        valorGuia = "" ' Inicializa o valor
        Dim wsValorGuia As Worksheet
        Set wsValorGuia = ThisWorkbook.Sheets("Valor_Guia")
        
        ' Localiza a linha da TabelaGuia correspondente ao mês (keyMensal)
        Dim tabelaGuia As ListObject
        Set tabelaGuia = wsValorGuia.ListObjects("TabelaGuia")
        
        ' Verifica se há uma correspondência com o mês de referência na TabelaGuia
        Dim linhaGuia As ListRow
        For Each linhaGuia In tabelaGuia.ListRows
            If Format(linhaGuia.Range(1).Value, "mm/yyyy") = keyMensal Then ' A coluna 1 da TabelaGuia é a data
                valorGuia = linhaGuia.Range(2).Value ' Coluna B da TabelaGuia
                Exit For
            End If
        Next linhaGuia

        With tabelaDestino2.ListRows.Add
            .Range(1).Value = keyMensal ' Mês/Ano
            .Range(2).Value = ThisWorkbook.Sheets("Empresa").Cells(3, 2)
            .Range(3).Value = valoresMensal(1) ' FAT_LIQUIDO
            .Range(5).Value = valoresMensal(2) ' ISS_RET
            .Range(9).Value = valorGuia ' Insere o valor da TabelaGuia na coluna I
        End With
    Next keyMensal
    
    Call CalcularRBT12
    
    Call CalcularSimples
    
    Call AtualizarAnexos
    
    Call OrdenarTabelas
    MsgBox "Tabela consolidada gerada com sucesso!", vbInformation
End Sub

Sub CalcularRBT12()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mesCalculo As Integer
    Dim anoCalculo As Integer
    Dim dataCalculo As Date
    Dim mesesAnteriores As Collection
    Dim mesesDisponiveis As Integer
    Dim somaFaturamento As Double
    Dim rbt12 As Double
    Dim j As Long
    Dim mesAnterior As Integer
    Dim anoAnterior As Integer
    Dim dataAnterior As Date
    Dim k As Long, l As Long
    Dim arrayMeses(11) As Date
    Dim nl As Long
    Dim validarMeses As Boolean, encontrou As Boolean
    
    ' Define a planilha e a última linha
    Set ws = ThisWorkbook.Sheets("Calc_Simples") ' Altere para o nome da sua planilha
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop para cada linha de dados (a partir da linha 2, assumindo que a linha 1 é o cabeçalho)
    For i = 2 To lastRow
        ' Obtém o mês e o ano do cálculo
        mesCalculo = Month(ws.Cells(i, 1).Value) ' Coluna MES
        anoCalculo = Year(ws.Cells(i, 1).Value) ' Coluna ANEXO (assumindo que é o ano)
        dataCalculo = DateSerial(anoCalculo, mesCalculo - 1, 1)
        
        ' Inicializa variáveis
        mesesDisponiveis = 0
        somaFaturamento = 0
        validarMeses = False
        
        
        For j = 0 To 11
            arrayMeses(j) = DateAdd("m", -j, dataCalculo)
        Next j
        For k = 11 To 0 Step -1
            For l = 2 To lastRow
                encontrou = False
                If Month(ws.Cells(l, 1)) = Month(arrayMeses(k)) And Year(ws.Cells(l, 1)) = Year(arrayMeses(k)) Then
                        somaFaturamento = somaFaturamento + ws.Cells(l, 3)
                        mesesDisponiveis = mesesDisponiveis + 1
                        validarMeses = True
                        encontrou = True
                        Exit For
                End If
            Next l
            If encontrou = False And validarMeses = True Then
                MsgBox "Mes " & arrayMeses(k) & " nao foi lancado na planilha."
                Exit Sub
            End If
        Next
        
        ' Calcula o RBT12 com base nos meses disponíveis
        If mesesDisponiveis = 0 Then
            rbt12 = 0 ' Se não houver meses anteriores
        ElseIf mesesDisponiveis < 12 Then
            rbt12 = somaFaturamento * (12 / mesesDisponiveis) ' Regra de três
        Else
            rbt12 = somaFaturamento ' Soma dos últimos 12 meses
        End If
        
        ' Preenche o valor do RBT12 na coluna correspondente
        ws.Cells(i, 4).Value = rbt12 ' Coluna RBT12
    Next i
    
End Sub

Sub CalcularSimples()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rbt12 As Double
    Dim fat As Double
    Dim anexo As String
    Dim aliqs As Variant
    
    Set ws = ThisWorkbook.Sheets("Calc_Simples")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
         rbt12 = ws.Cells(i, 4)
         fat = ws.Cells(i, 3)
         anexo = ws.Cells(i, 2)
         
         aliqs = CalcAliq(rbt12, anexo)
         
         
         ws.Cells(i, 6) = aliqs(1)
         ws.Cells(i, 7) = aliqs(2)
         If ThisWorkbook.Sheets("Empresa").Cells(5, 2) <> "Sim" Then
            ws.Cells(i, 8) = aliqs(1) * fat - ws.Cells(i, 5)
        Else
            ws.Cells(i, 8) = (aliqs(1) - aliqs(2)) * fat
        End If
    Next i
    
    
    
End Sub
Function CalcAliq(rbt12 As Double, anexo As String) As Variant
    Dim sheetName As String
    Dim ws As Worksheet
    Dim aliqs() As Double
    Dim ln As Integer
    
    If anexo = "" Then Exit Function
    
    sheetName = "ANEXO " & anexo
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ReDim aliqs(1 To 2)
    
    
    If rbt12 >= ws.Cells(4, 2) And rbt12 <= ws.Cells(4, 3) Then
        ln = 4
    ElseIf rbt12 >= ws.Cells(5, 2) And rbt12 <= ws.Cells(5, 3) Then
        ln = 5
    ElseIf rbt12 >= ws.Cells(6, 2) And rbt12 <= ws.Cells(6, 3) Then
        ln = 6
    ElseIf rbt12 >= ws.Cells(7, 2) And rbt12 <= ws.Cells(7, 3) Then
        ln = 7
    ElseIf rbt12 >= ws.Cells(8, 2) And rbt12 <= ws.Cells(8, 3) Then
        ln = 8
    ElseIf rbt12 >= ws.Cells(9, 2) And rbt12 <= ws.Cells(9, 3) Then
        ln = 9
    Else
        aliqs(1) = 0
        aliqs(2) = 0
        MsgBox "RBT12 excedeu valor máximo do Simples"
        Exit Function
    End If
    
    If rbt12 > ws.Cells(4, 3) Then
        aliqs(1) = (rbt12 * ws.Cells(ln, 4) - ws.Cells(ln, 5)) / rbt12
    Else
        aliqs(1) = ws.Cells(4, 4)
    End If
    
    If anexo = "III" Or anexo = "V" Then
        aliqs(2) = aliqs(1) * ws.Cells(ln, 11)
    ElseIf anexo = "IV" Then
        aliqs(2) = aliqs(1) * ws.Cells(ln, 10)
    Else
        aliqs(2) = 0
    End If
    
    CalcAliq = aliqs
    
End Function
Sub LançarDadosNaTabela()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    Dim linhaDados As Range
    Dim colunaC As Range
    Dim i As Integer
    Dim colunasObrigatorias As Variant
    Dim dadosValidos As Boolean
    
    ' Configurar variáveis
    Set ws = ThisWorkbook.Sheets(1) ' Alterar para o nome correto da planilha, se necessário
    Set tbl = ws.ListObjects("TabelaDados")
    Set linhaDados = ws.Range("A3:H3")
    Set colunaC = ws.Range("C3")
    
    ' Definir colunas obrigatórias
    colunasObrigatorias = Array("A3", "B3", "C3", "G3")
    dadosValidos = True
    
    ' Validar preenchimento das colunas obrigatórias
    For i = LBound(colunasObrigatorias) To UBound(colunasObrigatorias)
        If ws.Range(colunasObrigatorias(i)).Value = "" Then
            MsgBox "Erro: A célula " & colunasObrigatorias(i) & " deve estar preenchida.", vbExclamation
            dadosValidos = False
            Exit For
        End If
    Next i
    
    ' Verificar se C é uma data
    If dadosValidos And Not IsDate(colunaC.Value) Then
        MsgBox "Erro: A coluna C deve conter uma data válida.", vbExclamation
        dadosValidos = False
    End If
    
    ' Verificar relação entre D e F
    If dadosValidos Then
        Dim valorD As Variant, valorF As Variant
        valorD = linhaDados.Cells(1, 4).Value ' Coluna D
        valorF = linhaDados.Cells(1, 6).Value ' Coluna F
        
        If valorD = "" And valorF = "" Then
            MsgBox "Erro: Pelo menos uma das colunas D ou F deve estar preenchida.", vbExclamation
            dadosValidos = False
        ElseIf valorD <> "" And valorF <> "" Then
            MsgBox "Erro: As colunas D e F não podem estar preenchidas ao mesmo tempo.", vbExclamation
            dadosValidos = False
        End If
    End If
    
    ' Verificar se D, E e F são valores numéricos
    If dadosValidos Then
        For i = 4 To 6 ' Colunas D (4), E (5) e F (6)
            If linhaDados.Cells(1, i).Value <> "" And Not IsNumeric(linhaDados.Cells(1, i).Value) Then
                MsgBox "Erro: A coluna " & Chr(64 + i) & " deve conter um valor numérico.", vbExclamation
                dadosValidos = False
                Exit For
            End If
        Next i
    End If
    
    ' Se os dados forem válidos, adicionar na tabela
    If dadosValidos Then
        ' Adicionar nova linha na tabela
        Set novaLinha = tbl.ListRows.Add
        novaLinha.Range.Value = linhaDados.Value
        
        ' Limpar os dados da linha 3, exceto a coluna C
        For i = 1 To 8
            If i <> 3 Then ' Exceto coluna C
                linhaDados.Cells(1, i).ClearContents
            End If
        Next i
        
        ' Após o lançamento, configurar G3 como "Não"
        ws.Range("G3").Value = "Não"
        
        MsgBox "Dados lançados na tabela com sucesso!", vbInformation
    End If
End Sub

Sub AtualizarTabelaGuia()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim tblOrigem As ListObject
    Dim tblDestino As ListObject
    Dim rngOrigem As Range
    Dim rngDestino As Range
    Dim celOrigem As Range
    Dim celDestino As Range
    Dim mesReferencia As Date
    Dim valorAtualizado As Double
    Dim encontrada As Boolean
    Dim novaLinha As ListRow
    Dim ano As Integer, mes As Integer, dia As Integer
    
    ' Definir as planilhas
    Set wsOrigem = ThisWorkbook.Sheets("Calc_Simples")
    Set wsDestino = ThisWorkbook.Sheets("Valor_Guia")
    
    ' Verificar se as tabelas existem
    On Error Resume Next
    Set tblOrigem = wsOrigem.ListObjects("TabelaMensal")
    Set tblDestino = wsDestino.ListObjects("TabelaGuia")
    On Error GoTo 0
    
    ' Se alguma tabela não existir, mostrar mensagem de erro
    If tblOrigem Is Nothing Then
        MsgBox "Tabela 'TabelaMensal' não encontrada na planilha 'Calc_Simples'.", vbCritical
        Exit Sub
    End If
    If tblDestino Is Nothing Then
        MsgBox "Tabela 'TabelaGuia' não encontrada na planilha 'Valor_Guia'.", vbCritical
        Exit Sub
    End If
    
    ' Verificar se a Tabela de Origem tem dados
    If tblOrigem.ListRows.Count = 0 Then
        MsgBox "A Tabela 'TabelaMensal' está vazia.", vbCritical
        Exit Sub
    End If
    
    ' Acessar o intervalo de dados diretamente com DataBodyRange
    Set rngOrigem = tblOrigem.DataBodyRange ' Acessando apenas os dados da tabela
    
    ' Verificar se a Tabela de Destino tem dados
    If tblDestino.ListRows.Count > 0 Then
        Set rngDestino = tblDestino.ListColumns(1).DataBodyRange ' Coluna A da TabelaGuia
    Else
        ' Se não tiver dados, deixar rngDestino como Nothing
        Set rngDestino = Nothing
    End If
    
    ' Iterar pelas linhas da origem
    For Each celOrigem In rngOrigem.Columns(1).Cells ' A coluna 1 contém a data de referência
        If IsDate(celOrigem.Value) Then
            ' Separar o dia, mês e ano da data da origem
            ano = Year(celOrigem.Value)
            mes = Month(celOrigem.Value)
            dia = Day(celOrigem.Value)
            
            ' Criar a data corretamente com DateSerial
            mesReferencia = DateSerial(ano, mes, dia)
            valorAtualizado = celOrigem.Offset(0, 8).Value ' Coluna I está 8 colunas à direita de A
            encontrada = False
            
            ' Se rngDestino tiver dados, verificar na Tabela de Destino
            If Not rngDestino Is Nothing Then
                ' Iterar pela tabela de destino para verificar se o mês já existe
                For Each celDestino In rngDestino
                    If IsDate(celDestino.Value) Then
                        ' Separar a data de destino para comparação
                        ano = Year(celDestino.Value)
                        mes = Month(celDestino.Value)
                        dia = Day(celDestino.Value)
                        
                        ' Criar a data corretamente com DateSerial para a Tabela de Destino
                        If DateSerial(ano, mes, dia) = mesReferencia Then
                            ' Se o mês já existe, atualizar o valor correspondente
                            celDestino.Offset(0, 1).Value = valorAtualizado
                            encontrada = True
                            Exit For
                        End If
                    End If
                Next celDestino
            End If
            
            ' Se o mês não foi encontrado, adicionar uma nova linha na tabela de destino
            If Not encontrada Then
                ' Adicionar nova linha ao final da tabela
                Set novaLinha = tblDestino.ListRows.Add
                novaLinha.Range(1, 1).Value = mesReferencia
                novaLinha.Range(1, 2).Value = valorAtualizado
            End If
        End If
    Next celOrigem

    MsgBox "Atualização concluída!", vbInformation
End Sub


Sub AtualizarAnexos()
    Dim ws As Worksheet
    Dim maiorData As Date
    Dim mesAtual As Date, mesProx As Date
    Dim soma As Double
    Dim ultimaLinha As Long
    Dim i As Integer, j As Integer
    Dim mesEncontrado As Boolean
    Dim mesAnterior As Date
    Dim mesesEncontrados As Integer
    Dim mesAtualTabela As Date
    Dim listaMeses() As Date
    Dim indice As Integer

    ' Definir a planilha onde a TabelaMensal está localizada
    Set ws = ThisWorkbook.Sheets("Calc_Simples")

    ' Encontrar a maior data na TabelaMensal (coluna A)
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    maiorData = Application.WorksheetFunction.Max(ws.Range("A2:A" & ultimaLinha))

    ' Inicializar variáveis
    mesAtual = maiorData
    mesesEncontrados = 0
    soma = 0
    indice = 0

    ' Armazenar meses encontrados em uma lista
    ReDim listaMeses(11)

    ' Procurar até 12 meses anteriores
    For i = 0 To 11
        mesAtual = DateSerial(Year(maiorData), Month(maiorData) - i, 1)
        mesEncontrado = False

        ' Verificar se o mês está na tabela
        For j = 2 To ultimaLinha
            mesAtualTabela = ws.Cells(j, 1).Value
            If Format(mesAtualTabela, "mm/yyyy") = Format(mesAtual, "mm/yyyy") Then
                ' Adicionar soma da coluna C
                soma = soma + ws.Cells(j, 3).Value
                listaMeses(indice) = mesAtual
                indice = indice + 1
                mesEncontrado = True
                mesesEncontrados = mesesEncontrados + 1
                Exit For
            End If
        Next j

        ' Se não encontrou o mês, sair do loop (permitir quebra contínua)
        If Not mesEncontrado And mesesEncontrados > 0 Then
            Exit For
        End If
    Next i

    ' Validar continuidade dos meses
    For i = LBound(listaMeses) To UBound(listaMeses) - 1
        If listaMeses(i) <> 0 And listaMeses(i + 1) <> 0 Then
            If DateDiff("m", listaMeses(i + 1), listaMeses(i)) <> 1 Then
                MsgBox "Erro: Os meses anteriores não são contínuos.", vbCritical
                Exit Sub
            End If
        End If
    Next i
    
    mesProx = DateSerial(Year(maiorData), Month(maiorData) + 1, 1)
    
    Dim anexos() As Variant
    Dim wsAnx As Worksheet
    Dim x As Integer, y As Integer
    Dim aliqs As Variant
    Dim anexo As String
    anexos = Array("I", "II", "III", "IV", "V")
    
    For x = 0 To 4
        anexo = anexos(x)
        Set wsAnx = ThisWorkbook.Sheets("ANEXO " & anexo)
        wsAnx.Cells(12, 1) = Format(mesProx, "mmm/yy")
        wsAnx.Cells(12, 2) = soma
        aliqs = CalcAliqTodos(soma, anexo)
        For y = LBound(aliqs) To UBound(aliqs)
            wsAnx.Cells(12, 2 + y) = aliqs(y)
        Next y
    Next x
    
    
End Sub
Function CalcAliqTodos(rbt12 As Double, anexo As String) As Variant
    Dim sheetName As String
    Dim ws As Worksheet
    Dim aliqs() As Double
    Dim ln As Integer
    
    If anexo = "" Then Exit Function
    
    sheetName = "ANEXO " & anexo
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    If anexo = "V" Or anexo = "III" Or anexo = "I" Then
        ReDim aliqs(1 To 10)
    End If
    If anexo = "II" Then
        ReDim aliqs(1 To 11)
    End If
    If anexo = "IV" Then
        ReDim aliqs(1 To 9)
    End If
    
    
    If rbt12 >= ws.Cells(4, 2) And rbt12 <= ws.Cells(4, 3) Then
        ln = 4
    ElseIf rbt12 >= ws.Cells(5, 2) And rbt12 <= ws.Cells(5, 3) Then
        ln = 5
    ElseIf rbt12 >= ws.Cells(6, 2) And rbt12 <= ws.Cells(6, 3) Then
        ln = 6
    ElseIf rbt12 >= ws.Cells(7, 2) And rbt12 <= ws.Cells(7, 3) Then
        ln = 7
    ElseIf rbt12 >= ws.Cells(8, 2) And rbt12 <= ws.Cells(8, 3) Then
        ln = 8
    ElseIf rbt12 >= ws.Cells(9, 2) And rbt12 <= ws.Cells(9, 3) Then
        ln = 9
    Else
        aliqs(1) = 0
        aliqs(2) = 0
        MsgBox "RBT12 excedeu valor máximo do Simples"
        Exit Function
    End If
    
    If rbt12 > ws.Cells(4, 3) Then
        aliqs(4) = (rbt12 * ws.Cells(ln, 4) - ws.Cells(ln, 5)) / rbt12
    Else
        aliqs(4) = ws.Cells(4, 4)
    End If
    
    If anexo = "V" Or anexo = "III" Or anexo = "I" Then
        aliqs(1) = ws.Cells(ln, 4)
        aliqs(2) = rbt12 * ws.Cells(ln, 4)
        aliqs(3) = ws.Cells(ln, 5)
        aliqs(5) = ws.Cells(ln, 6) * aliqs(4)
        aliqs(6) = ws.Cells(ln, 7) * aliqs(4)
        aliqs(7) = ws.Cells(ln, 8) * aliqs(4)
        aliqs(8) = ws.Cells(ln, 9) * aliqs(4)
        aliqs(9) = ws.Cells(ln, 10) * aliqs(4)
        aliqs(10) = ws.Cells(ln, 11) * aliqs(4)
    End If
    If anexo = "II" Then
        aliqs(1) = ws.Cells(ln, 4)
        aliqs(2) = rbt12 * ws.Cells(ln, 4)
        aliqs(3) = ws.Cells(ln, 5)
        aliqs(5) = ws.Cells(ln, 6) * aliqs(4)
        aliqs(6) = ws.Cells(ln, 7) * aliqs(4)
        aliqs(7) = ws.Cells(ln, 8) * aliqs(4)
        aliqs(8) = ws.Cells(ln, 9) * aliqs(4)
        aliqs(9) = ws.Cells(ln, 10) * aliqs(4)
        aliqs(10) = ws.Cells(ln, 11) * aliqs(4)
        aliqs(11) = ws.Cells(ln, 12) * aliqs(4)
    End If
    If anexo = "IV" Then
        aliqs(1) = ws.Cells(ln, 4)
        aliqs(2) = rbt12 * ws.Cells(ln, 4)
        aliqs(3) = ws.Cells(ln, 5)
        aliqs(5) = ws.Cells(ln, 6) * aliqs(4)
        aliqs(6) = ws.Cells(ln, 7) * aliqs(4)
        aliqs(7) = ws.Cells(ln, 8) * aliqs(4)
        aliqs(8) = ws.Cells(ln, 9) * aliqs(4)
        aliqs(9) = ws.Cells(ln, 10) * aliqs(4)
    End If
    
    CalcAliqTodos = aliqs
    
End Function
Sub OrdenarTabelas()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIndex As Integer
    Dim tables As Variant
    Dim colunas As Variant
    Dim abas As Variant
    Dim i As Integer
    Dim colunaData As Range
    
    ' Definir as tabelas, as colunas e as planilhas correspondentes
    tables = Array("TabelaDados", "TabelaConsolidada", "TabelaMensal")
    colunas = Array(3, 1, 1) ' Índices das colunas para cada tabela
    abas = Array("Lancamentos", "Resumo", "Calc_Simples") ' Nomes das planilhas onde estão as tabelas

    ' Desativar atualizações para melhorar desempenho
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual


    ' Loop para percorrer todas as tabelas
    For i = LBound(tables) To UBound(tables)
        ' Ativar a planilha correta antes de acessar a tabela
        Set ws = ThisWorkbook.Sheets(abas(i))
        ws.Activate  ' Garantir que a planilha esteja ativa
        
        ' Buscar a tabela
        On Error Resume Next
        Set tbl = ws.ListObjects(tables(i))
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            colIndex = colunas(i) ' Obtém o índice da coluna correspondente
            
            ' Pegar a coluna correta dentro da tabela
            Set colunaData = tbl.ListColumns(colIndex).DataBodyRange
            
            ' Desativar filtros, se estiverem ativos
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            
            ' Ordenar a tabela pela coluna especificada (do mais recente para o mais antigo)
            With tbl.Sort
                .SortFields.Clear
                .SortFields.Add2 key:=colunaData, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .Header = xlYes
                .Apply
            End With
            
            ' Forçar a atualização da tela para refletir a mudança
            DoEvents
        End If
    Next i
    
    ' Restaurar configurações
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

