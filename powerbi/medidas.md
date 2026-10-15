# Medidas DAX

Todas as medidas ficam na tabela `_Medidas`. O projeto usa **dados 100% fictícios**.

## Convenções

- **AA** = mesmo período do ano anterior (`SAMEPERIODLASTYEAR`).
- **p.p.** = pontos percentuais: variação de um indicador que já é percentual (subtração, não divisão).
- **Entregas válidas** = `fEntregas[valido_kpi] = 1`: pedido com status "Entregue", datas coerentes e quantidade entregue informada. Pedidos "Em trânsito" e entregas com dado inválido ficam fora de OTIF, prazo e lead time.
- Cores em hexadecimal: verde `#2E7D32` (melhor que a meta ou alta), vermelho `#C0392B` (pior que a meta ou queda), âmbar `#E0A100` (atenção), azul-acinzentado `#5B8E9B` e cinza `#9AA5B1` (neutro).
- Para "melhor" no lead time, a lógica é invertida: **menor é melhor**.

**Colunas criadas no Power Query e usadas aqui:** `fEntregas[no_prazo]`, `[completo]`, `[otif]`, `[lead_time_dias]`, `[valido_kpi]`; `fEstoque[dias_sem_movimento]`, `[flag_estoque_invalido]`; `dProdutos[custo]` (com custo imputado nos SKUs sem custo válido).

---

## 1. Vendas

| Medida | O que faz | Expressão |
|---|---|---|
| Faturamento | Soma, linha a linha, de quantidade × preço × (1 − desconto). Usa `SUMX` (contexto de linha) em vez de somar colunas separadas. | `SUMX ( fVendas, fVendas[quantidade] * fVendas[preco_unitario] * ( 1 - fVendas[desconto_pct] ) )` |
| Custo Total | Custo das unidades vendidas. `RELATED` traz o custo unitário de `dProdutos` para cada linha de venda. | `SUMX ( fVendas, fVendas[quantidade] * RELATED ( dProdutos[custo] ) )` |
| Margem R$ | Lucro bruto em reais: faturamento menos custo total. | `[Faturamento] - [Custo Total]` |
| Margem % | Margem sobre o faturamento. Usa `DIVIDE` para evitar erro de divisão por zero. | `DIVIDE ( [Margem R$], [Faturamento] )` |
| Pedidos | Quantidade de pedidos distintos (um pedido tem vários itens). | `DISTINCTCOUNT ( fVendas[id_pedido] )` |
| Ticket Médio | Faturamento médio por pedido. | `DIVIDE ( [Faturamento], [Pedidos] )` |

## 2. Comparação com o período anterior

| Medida | O que faz | Expressão |
|---|---|---|
| Faturamento AA | Faturamento do mesmo período do ano anterior. `CALCULATE` altera o contexto de filtro de data. Fica em branco em 2024, que é o primeiro ano da base. | `CALCULATE ( [Faturamento], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. Faturamento vs AA % | Variação percentual do faturamento sobre o ano anterior. | `DIVIDE ( [Faturamento] - [Faturamento AA], [Faturamento AA] )` |
| Faturamento Mês Anterior | Faturamento do mês imediatamente anterior ao do contexto. | `CALCULATE ( [Faturamento], DATEADD ( dCalendario[Date], -1, MONTH ) )` |
| Faturamento YTD | Faturamento acumulado do ano até a data do contexto (*year to date*). | `TOTALYTD ( [Faturamento], dCalendario[Date] )` |
| Margem % AA | Margem percentual do mesmo período do ano anterior. | `CALCULATE ( [Margem %], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. Margem p.p. | Diferença da margem em pontos percentuais (subtração, pois margem já é %). | `[Margem %] - [Margem % AA]` |
| Pedidos AA | Número de pedidos no mesmo período do ano anterior. | `CALCULATE ( [Pedidos], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. Pedidos % | Variação percentual do número de pedidos sobre o ano anterior. | `DIVIDE ( [Pedidos] - [Pedidos AA], [Pedidos AA] )` |
| Ticket Médio AA | Ticket médio do ano anterior, calculado dentro do período deslocado (razão do ano anterior, não mistura anos). | `CALCULATE ( [Ticket Médio], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. Ticket Médio % | Variação percentual do ticket médio sobre o ano anterior. | `DIVIDE ( [Ticket Médio] - [Ticket Médio AA], [Ticket Médio AA] )` |

## 3. Logística

| Medida | O que faz | Expressão |
|---|---|---|
| OTIF % | *On Time In Full*: percentual de entregas válidas feitas no prazo **e** completas. | `DIVIDE ( CALCULATE ( SUM ( fEntregas[otif] ), fEntregas[valido_kpi] = 1 ), CALCULATE ( COUNTROWS ( fEntregas ), fEntregas[valido_kpi] = 1 ) )` |
| % No Prazo | Percentual de entregas válidas entregues até a data prometida. | `DIVIDE ( CALCULATE ( SUM ( fEntregas[no_prazo] ), fEntregas[valido_kpi] = 1 ), CALCULATE ( COUNTROWS ( fEntregas ), fEntregas[valido_kpi] = 1 ) )` |
| % Completo | Percentual de entregas válidas com quantidade entregue igual ou maior que a pedida. | `DIVIDE ( CALCULATE ( SUM ( fEntregas[completo] ), fEntregas[valido_kpi] = 1 ), CALCULATE ( COUNTROWS ( fEntregas ), fEntregas[valido_kpi] = 1 ) )` |
| % Atraso | Percentual de entregas válidas fora do prazo (o complemento de % No Prazo). | `1 - DIVIDE ( CALCULATE ( SUM ( fEntregas[no_prazo] ), fEntregas[valido_kpi] = 1 ), CALCULATE ( COUNTROWS ( fEntregas ), fEntregas[valido_kpi] = 1 ) )` |
| Lead Time Médio | Média de dias entre o pedido e a entrega. | `CALCULATE ( AVERAGE ( fEntregas[lead_time_dias] ) )` |
| OTIF % AA | OTIF do mesmo período do ano anterior. | `CALCULATE ( [OTIF %], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. OTIF p.p. | Diferença do OTIF em pontos percentuais. | `[OTIF %] - [OTIF % AA]` |
| Lead Time AA | Lead time médio do ano anterior. | `CALCULATE ( [Lead Time Médio], SAMEPERIODLASTYEAR ( dCalendario[Date] ) )` |
| Var. Lead Time | Diferença do lead time em dias (negativo = melhorou). | `[Lead Time Médio] - [Lead Time AA]` |

## 4. Estoque

| Medida | O que faz | Expressão |
|---|---|---|
| Data Posição | Data da "foto" do estoque, usada como referência de consumo e de dias parados. | `DATE ( 2025, 12, 31 )` |
| Consumo Diário (90d) | Venda média diária em unidades nos 90 dias até a data da foto. `REMOVEFILTERS` ignora as segmentações de data, então o consumo não depende do ano escolhido. | `VAR DataRef = [Data Posição] RETURN DIVIDE ( CALCULATE ( SUM ( fVendas[quantidade] ), REMOVEFILTERS ( dCalendario ), dCalendario[Date] > DataRef - 90, dCalendario[Date] <= DataRef ), 90 )` |
| Qtd em Estoque | Unidades em estoque, ignorando linhas com estoque inválido (negativo). | `CALCULATE ( SUM ( fEstoque[qtd_atual] ), fEstoque[flag_estoque_invalido] = 0 )` |
| Cobertura (dias) | Por quantos dias o estoque atual aguenta no ritmo de consumo recente. Faz sentido por SKU, não no total (unidades diferentes). | `DIVIDE ( [Qtd em Estoque], [Consumo Diário (90d)] )` |
| Cobertura Mediana (dias) | Mediana da cobertura entre os SKUs. Preferida à média, que é puxada pelos itens em excesso. | `MEDIANX ( VALUES ( dProdutos[sku] ), [Cobertura (dias)] )` |
| Dias Sem Movimento | Dias desde a última movimentação do SKU (máximo; pensado para uso por SKU). | `MAX ( fEstoque[dias_sem_movimento] )` |
| Valor em Estoque | Valor do estoque ao custo: quantidade × custo unitário de cada SKU. | `SUMX ( fEstoque, fEstoque[qtd_atual] * RELATED ( dProdutos[custo] ) )` |
| Valor Parado | Valor em estoque dos itens sem movimentação há mais dias que a janela definida. | `VAR Janela = [Janela Parado] RETURN CALCULATE ( [Valor em Estoque], fEstoque[dias_sem_movimento] > Janela )` |
| Valor Parado % | Participação do capital parado no valor total em estoque. | `DIVIDE ( [Valor Parado], [Valor em Estoque] )` |
| Itens em Risco | Número de SKUs com cobertura abaixo do mínimo. Percorre os SKUs (contexto de linha) e usa `NOT ISBLANK` porque o DAX trata vazio como zero. | `COUNTROWS ( FILTER ( VALUES ( dProdutos[sku] ), NOT ISBLANK ( [Cobertura (dias)] ) && [Cobertura (dias)] < [Cobertura Mínima] ) )` |
| Itens em Excesso | Número de SKUs com cobertura acima do máximo. | `COUNTROWS ( FILTER ( VALUES ( dProdutos[sku] ), NOT ISBLANK ( [Cobertura (dias)] ) && [Cobertura (dias)] > [Cobertura Máxima] ) )` |
| Itens Parados | Número de SKUs com estoque positivo e sem movimentação há mais dias que a janela. | `COUNTROWS ( FILTER ( VALUES ( dProdutos[sku] ), [Qtd em Estoque] > 0 && [Dias Sem Movimento] > [Janela Parado] ) )` |
| Status SKU | Classifica cada SKU em Sem estoque, Parado, Sem consumo, Ruptura, Excesso ou OK. A ordem dos testes importa: o `SWITCH` para no primeiro verdadeiro. | `VAR Dias = [Dias Sem Movimento] VAR Cob = [Cobertura (dias)] VAR Qtd = [Qtd em Estoque] RETURN SWITCH ( TRUE (), ISBLANK ( Qtd ), BLANK (), Qtd <= 0 && ISBLANK ( Cob ), "Sem estoque", Dias > [Janela Parado], "Parado", ISBLANK ( Cob ), "Sem consumo", Cob < [Cobertura Mínima], "Ruptura", Cob > [Cobertura Máxima], "Excesso", "OK" )` |

## 5. Metas e parâmetros (fictícios)

| Medida | O que faz | Expressão |
|---|---|---|
| Meta Margem % | Meta de margem: 28%. | `0.28` |
| Meta OTIF | Meta de OTIF: 92%. | `0.92` |
| Meta Lead Time | Meta de lead time: 7 dias. | `7` |
| Cobertura Mínima | Abaixo disso, o SKU está em risco de ruptura (dias). | `15` |
| Cobertura Máxima | Acima disso, o SKU está em excesso (dias). | `45` |
| Janela Parado | Dias sem movimentação a partir dos quais um item é considerado parado. | `90` |

## 6. Cores para formatação condicional

Usadas em **Formatar visual → fx → Valor do campo**. Retornam códigos hexadecimais.

| Medida | O que faz | Expressão |
|---|---|---|
| Cor Margem | Verde se a margem está na meta ou acima, vermelho se abaixo. | `IF ( [Margem %] >= [Meta Margem %], "#2E7D32", "#C0392B" )` |
| Cor OTIF | Verde se o OTIF atinge a meta, vermelho se não. | `IF ( [OTIF %] >= [Meta OTIF], "#2E7D32", "#C0392B" )` |
| Cor Lead Time | Lógica invertida: verde se o lead time está dentro da meta (menor é melhor). | `IF ( [Lead Time Médio] <= [Meta Lead Time], "#2E7D32", "#C0392B" )` |
| Cor Status | Cor de cada status de SKU: vermelho (Ruptura), âmbar (Parado), azul-acinzentado (Excesso), verde (OK), cinza (demais). | `SWITCH ( [Status SKU], "Ruptura", "#C0392B", "Parado", "#E0A100", "Excesso", "#5B8E9B", "Sem estoque", "#9AA5B1", "OK", "#2E7D32", "#9AA5B1" )` |
| Var Fat Cor | Verde se o faturamento cresceu sobre o ano anterior, vermelho se caiu. | `IF ( [Var. Faturamento vs AA %] >= 0, "#2E7D32", "#C0392B" )` |
| Var Margem Cor | Verde se a margem subiu (ou ficou igual) sobre o ano anterior. | `IF ( [Var. Margem p.p.] >= 0, "#2E7D32", "#C0392B" )` |
| Var OTIF Cor | Verde se o OTIF subiu (ou ficou igual) sobre o ano anterior. | `IF ( [Var. OTIF p.p.] >= 0, "#2E7D32", "#C0392B" )` |
| Var Lead Cor | Verde se o lead time diminuiu ou ficou igual (menor é melhor). | `IF ( [Var. Lead Time] <= 0, "#2E7D32", "#C0392B" )` |
| Var Pedidos Cor | Verde se o número de pedidos subiu sobre o ano anterior. | `IF ( [Var. Pedidos %] >= 0, "#2E7D32", "#C0392B" )` |
| Var Ticket Cor | Verde se o ticket médio subiu sobre o ano anterior. | `IF ( [Var. Ticket Médio %] >= 0, "#2E7D32", "#C0392B" )` |

## 7. Textos dinâmicos dos cartões

Montam a linha de variação dos cartões (seta, variação e base de comparação). Mostram "sem comparação" quando não há ano anterior (por exemplo, em 2024).

| Medida | O que faz | Expressão |
|---|---|---|
| Var Fat Texto | Texto da variação do faturamento: seta e % sobre o ano anterior. | `VAR v = [Var. Faturamento vs AA %] RETURN IF ( ISBLANK ( v ), "sem comparação", IF ( v >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( v ), "0.0%" ) & " vs ano anterior" )` |
| Base Fat AA Texto | Texto com o faturamento do ano anterior em R$ milhões. | `VAR aa = [Faturamento AA] RETURN IF ( ISBLANK ( aa ), "", "vs R$ " & FORMAT ( aa / 1000000, "0.0" ) & " mi no ano anterior" )` |
| Var Margem Texto | Texto da variação da margem em p.p., com a margem do ano anterior. | `VAR aa = [Margem % AA] VAR v = [Var. Margem p.p.] RETURN IF ( ISBLANK ( aa ), "sem comparação", IF ( v >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( v ) * 100, "0.0" ) & " p.p. vs " & FORMAT ( aa, "0.0%" ) & " no ano anterior" )` |
| Var Pedidos Texto | Texto da variação de pedidos: % e diferença em unidades. | `VAR aa = [Pedidos AA] VAR v = [Var. Pedidos %] RETURN IF ( ISBLANK ( aa ), "sem comparação", IF ( v >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( v ), "0.0%" ) & " (" & IF ( v >= 0, "+", "-" ) & FORMAT ( ABS ( [Pedidos] - aa ), "#,0" ) & " pedidos)" )` |
| Var Ticket Texto | Texto da variação do ticket médio: % e diferença em R$. | `VAR aa = [Ticket Médio AA] VAR v = [Var. Ticket Médio %] RETURN IF ( ISBLANK ( aa ), "sem comparação", IF ( v >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( v ), "0.0%" ) & " (" & IF ( v >= 0, "+", "-" ) & "R$ " & FORMAT ( ABS ( [Ticket Médio] - aa ), "#,0" ) & ")" )` |
| Base Ticket AA Texto | Texto com o ticket médio do ano anterior. | `VAR aa = [Ticket Médio AA] RETURN IF ( ISBLANK ( aa ), "", "vs R$ " & FORMAT ( aa, "#,0" ) & " no ano anterior" )` |
| Var OTIF Texto | Texto da variação do OTIF em p.p., com o OTIF do ano anterior. | `VAR aa = [OTIF % AA] VAR v = [Var. OTIF p.p.] RETURN IF ( ISBLANK ( aa ), "sem comparação", IF ( v >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( v ) * 100, "0.0" ) & " p.p. vs " & FORMAT ( aa, "0.0%" ) & " no ano anterior" )` |
| Var Lead Texto | Texto da variação do lead time em dias. A seta aponta para baixo quando o prazo diminui. | `VAR aa = [Lead Time AA] VAR v = [Var. Lead Time] RETURN IF ( ISBLANK ( aa ), "sem comparação", IF ( v <= 0, "▼ ", "▲ " ) & FORMAT ( ABS ( v ), "0.0" ) & " dias vs " & FORMAT ( aa, "0.0" ) & " no ano anterior" )` |
| Título Região | Título dinâmico do gráfico de região: as duas regiões líderes e sua participação. `ALLSELECTED` respeita as segmentações mas ignora o recorte da própria barra. | `VAR Tot = CALCULATE ( [Faturamento], ALLSELECTED ( dClientes[regiao] ) ) VAR Top2 = TOPN ( 2, ALLSELECTED ( dClientes[regiao] ), [Faturamento], DESC ) RETURN CONCATENATEX ( Top2, dClientes[regiao], " e ", [Faturamento], DESC ) & " somam " & FORMAT ( DIVIDE ( SUMX ( Top2, [Faturamento] ), Tot ), "0%" ) & " das vendas"` |

## 8. Textos fixos (rótulos dos cartões de Estoque)

| Medida | O que faz | Expressão |
|---|---|---|
| Risco Texto | Legenda do cartão "Itens em risco". | `"Cobertura abaixo de 15 dias"` |
| Itens Parados Texto | Legenda do cartão "Itens parados". | `"Sem movimentação há mais de 90 dias"` |
| Itens em Excesso Texto | Legenda do cartão "Itens em excesso". | `"Cobertura acima de 45 dias"` |
| Valor Parado Texto | Legenda do cartão "Valor parado %". | `"% do valor total em estoque"` |
| Cobertura Texto | Legenda do cartão de cobertura mediana. | `"Meta: 15 a 45 dias"` |

> Os textos fixos repetem os valores das metas. Se as metas mudarem, as legendas também devem ser atualizadas.