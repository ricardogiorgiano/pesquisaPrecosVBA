# pesquisaPrecosVBA
 Pesquisa de preços em vários martketplaces usando VBA e Excel.
 
 Sub buscaLinkseComparadorDePreco()
'Desenvolvido por Ricardo Giorgiano do Nascimento
'Para GETEZ

' Caixa de mensagem para começo da pesquisa
resp = MsgBox("Deseja iniciar o Comparador de preços? Isto pode levar alguns minutos!!", vbYesNo)
If resp <> 6 Then Exit Sub

'Limpa Barra de progresso
[E1] = ""

'Data e hora da pesuisa
Cells(1, 4).Value = Now

'Conta última linha da planilha
ultLin = Range("C1000000").End(xlUp).Row

'Verifica se apartir da linha 3 há algum dado para pesquisa
If ultLin = 3 Then MsgBox ("Favor inserir Produtos para pesquisa!!")
If ultLin = 3 Then Exit Sub

'Limpa os dados apartir da coluna F até a coluna P e todas as linhas
'após a linha 3
Range("F4:P4" & ultLin).ClearContents

'Criando objeto Internet Explorer
Set IE = CreateObject("InternetExplorer.Application")

'Percorrendo linhas apartir da linha 4
For lin = 4 To ultLin
    
'--------------------------------------------------------------------------
    ' PESQUISA DE ANÚNCIOS AMAZON
    linkAM = ""
    precoProdutoAM = ""
    UrlAM = "https://www.amazon.com.br/s?k="
    
    'concatenação do site de busca com a coluna C e linha percorrida
    IE.Navigate UrlAM & Range("C" & lin)
    
    'while serve para esperar a leitura do navegador
    Do While IE.ReadyState <> 4: Loop
    
    'Removendo a aspas simples é possível ver o navegador trabalhando
    'IE.Visible = True
    
    'tratamento de erro caso a página não retorne nenhum resultado
    On Error GoTo VazioAM
    
    'criação do link
    linkAM = IE.document.getElementsByTagName("h2")(0).getElementsByTagName("A")(0).href
    Do While IE.ReadyState <> 4: Loop
    On Error GoTo VazioAM
    
    'gravando link na planilha
    Range("L" & lin) = linkAM
    
    On Error GoTo VazioAM
    
VazioAM:
    
    If linkAM = "" Then Range("L" & lin) = "Produto não encontrado!"
    
    Resume PassarErroVazioAM
PassarErroVazioAM:
    If linkAM = "" Then GoTo VazioAMp
On Error GoTo 0
    
    'PESQUISA DE PREÇO DO ANÚNCIO AMAZON
    'abrindo o link pesquisado
    IE.Navigate linkAM
    
    On Error GoTo VazioAMp
    Do While IE.ReadyState <> 4: Loop
    'IE.Visible = True
    On Error GoTo VazioAMp
    
    'buscando o elemento preço dentro do anúncio
    precoProdutoAM = CDbl(Replace(IE.document.getElementsByClassName("a-offscreen")(0).innerText, "R$", ""))
    On Error GoTo VazioAMp
    Do While IE.ReadyState <> 4: Loop
    precoProduto = precoProdutoAM
    
    'gravando na planilha o preço
    Cells(lin, 9).Value = precoProdutoAM 'Preço Amazon
    
VazioAMp:
    If linkAM = "" Then Cells(lin, 9).Value = "ND"
    If linkAM = "" Then GoTo PassarErroVazioAMp
    Resume PassarErroVazioAMp
PassarErroVazioAMp:
On Error GoTo 0

'------------------------------------------------------------------
    ' PESQUISA DE ANÚNCIOS LOJAS AMERICANAS
    linkLA = ""
    precoProdutoLA = ""
    UrlLA = "https://www.americanas.com.br/busca/"
    
    'concatenação do site de busca com a coluna C e linha percorrida
    IE.Navigate UrlLA & Range("C" & lin)
    
    'while serve para esperar a leitura do navegador
    Do While IE.ReadyState <> 4: Loop
    
    'IE.Visible = True
    
    'tratamento de erro caso a página não retorne nenhum resultado
    On Error GoTo VazioLA
    Do While IE.ReadyState <> 4: Loop
    
    'criação do link
    linkLA = IE.document.getElementsByClassName("inStockCard__Wrapper-sc-1ngt5zo-0 iRvjrG")(0).getElementsByTagName("A")(0).href
    On Error GoTo VazioLA
    
    'gravando link na planilha
    Range("M" & lin) = linkLA
    
    On Error GoTo VazioLA
VazioLA:
    If linkLA = "" Then Range("M" & lin) = "Pruduto não encontrado!"
    If linkLA = "" Then Cells(lin, 10).Value = "ND"
    Resume PassarErroVazioLA
PassarErroVazioLA:
On Error GoTo 0

    'PESQUISA DE PREÇO DO ANÚNCIO
    'abrindo o link pesquisado
    If linkLA = "" Then GoTo VazioLAp
 
    IE.Navigate linkLA
    Do While IE.ReadyState <> 4: Loop
    
    'IE.Visible = True
    On Error GoTo VazioLAp
    
    'buscando o elemento preço dentro do anúncio
    precoProdutoLA = CDbl(Replace(IE.document.getElementsByClassName("styles__PriceText-sc-x06r9i-0 dUTOlD priceSales")(0).innerText, "R$ ", ""))
    Do While IE.ReadyState <> 4: Loop
    precoProduto = precoProdutoLA
    On Error GoTo VazioLAp
    
    'gravando na planilha o preço
    Cells(lin, 10).Value = precoProdutoLA 'Preço Lojas Americanas

    
VazioLAp:
    If linkLA = "" Then Cells(lin, 10).Value = "ND"
    If linkLA = "" Then GoTo PassarErroVazioLAp
    Resume PassarErroVazioLAp
PassarErroVazioLAp:
On Error GoTo 0

'------------------------------------------------------------------
    ' PESQUISA DE ANÚNCIOS MAGALU
    linkMA = ""
    precoProdutoMA = ""
    UrlMA = "https://www.magazineluiza.com.br/busca/"
    
    'concatenação do site de busca com a coluna C e linha percorrida
    IE.Navigate UrlMA & Range("E" & lin)
    Do While IE.ReadyState <> 4: Loop
    
    'IE.Visible = True
    On Error GoTo VazioMA
    
    'criação do link
    linkMA = IE.document.getElementsByClassName("sc-ePIFMk jjJcaw")(0).getElementsByTagName("A")(0).href
    Do While IE.ReadyState <> 4: Loop
    On Error GoTo VazioMA
    
    'gravando na planilha
    Range("N" & lin) = linkMA
    
    On Error GoTo VazioMA
VazioMA:
    If linkMA = "" Then Range("N" & lin) = "Pruduto não encontrado!"
    If linkMA = "" Then Cells(lin, 11).Value = "ND"
    Resume passarErroVazioMA
passarErroVazioMA:

On Error GoTo 0
    
    'PESQUISA DE PREÇO DO ANÚNCIO
    'abrindo o link pesquisado
    If linkMA = "" Then GoTo VazioMAp
    
    IE.Navigate linkMA
    Do While IE.ReadyState <> 4: Loop
    On Error GoTo VazioMAp
    
    'IE.Visible = True
    
    'buscando o elemento preço dentro do anúncio
    precoProdutoMA = CDbl(IE.document.getElementsByClassName("price-template__text")(0).innerText)
    Do While IE.ReadyState <> 4: Loop
    On Error GoTo VazioMAp
    precoProduto = precoProdutoMA
    
    'gravando na planilha o preço
    Cells(lin, 11).Value = precoProdutoMA 'Preço Magalu

VazioMAp:
        If linkMA = "" Then Cells(lin, 11).Value = "ND"
        If linkMA = "" Then GoTo PassarErroVazioMAp
        Resume PassarErroVazioMAp
        
PassarErroVazioMAp:
    On Error GoTo 0


'------------------------------------------------------------------
    ' PESQUISA DE ANÚNCIOS MAGALU
    
    linkGO = ""
    precoProdutoGO = ""
    UrlGO = "https://www.google.com.br/search?tbm=shop&hl=pt-BR&psb=1&ved=2ahUKEwiW_9WUsOv2AhWRUEgAHUAeAQ0Qu-kFegQIABAK&q="
    
    'concatenação do site de busca com a coluna C e linha percorrida
    IE.Navigate UrlGO & Range("C" & lin)
    
    Do While IE.ReadyState <> 4: Loop
    'IE.Visible = True
    On Error GoTo VazioGO
    
    'criação do link
    linkGO = IE.document.getElementsByClassName("rgHvZc")(0).getElementsByTagName("A")(0).href
    
    'buscando o elemento preço dentro do anúncio
    precoProdutoGO = IE.document.getElementsByClassName("HRLxBb")(0).innerText
    
    'gravando na planilha
    Range("O" & lin) = precoProdutoGO
    Range("P" & lin) = linkGO
    On Error GoTo VazioGO
    
VazioGO:
    If linkGO = "" Then Range("P" & lin) = "Pruduto não encontrado!"
    If precoProdutoGO = "" Then Range("O" & lin) = "ND"
    Resume PassarErroVazioGO
PassarErroVazioGO:
On Error GoTo 0
'---------------------------------------------------------------------

'---------COMPARAÇÃO DE MENOR PREÇO-----------------------------------

menorPreco = ""
precoProduto = ""
        
    For col = 12 To 14
        
        'Percorrendo colunas dos preços
        If col = 12 Then
            precoProduto = precoProdutoAM
        ElseIf col = 13 Then
            precoProduto = precoProdutoLA
        Else
            precoProduto = precoProdutoMA
        End If
                     
        
        'Identificando menor preço e atribuindo loja menor preço
        If menorPreco = "" Then
          menorPreco = precoProduto
          lojaMenorPreco = Cells(3, col).Value
        ElseIf precoProduto < menorPreco Then
          menorPreco = precoProduto
          lojaMenorPreco = Cells(3, col).Value
         End If
              
         If lojaMenorPreco = "Amazon" Then
            linkMenorPreco = Cells(lin, 12).Text
         ElseIf lojaMenorPreco = "Lojas Americanas" Then
            linkMenorPreco = Cells(lin, 13).Text
         ElseIf lojaMenorPreco = "Magalu" Then
            linkMenorPreco = Cells(lin, 14).Text
        End If
    Next col
    
  Cells(lin, 6).Value = menorPreco 'Definindo o melhor preço
  Cells(lin, 7).Value = lojaMenorPreco 'Definindo loja com menor preço
  Cells(lin, 8).Value = linkMenorPreco 'Link do menor preço
  
  [E1] = lin / ultLin 'Barra de progresso

Next lin
        
  
'Fecha intenet explorer
IE.Quit

Set IE = Nothing

'Mensagem de conclusão da pesquisa
MsgBox ("Comparador de preços Concluído com Sucesso!!!")


Exit Sub

'Desenvolvido por Ricardo Giorgiano do Nascimento
'Para GETEZ

End Sub


