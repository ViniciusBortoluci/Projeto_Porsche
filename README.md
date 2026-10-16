# Projeto: Dashboard de Vendas Porsche

O projeto utiliza a base `Vendas Porsche.xlsx` para analisar vendas de veículos Porsche e alimentar um dashboard desenvolvido em HTML. A planilha `Tabela_Porsche_Sanitizada.xlsx` apresenta os campos com nomes padronizados em inglês e colunas adicionais com os valores tratados.

## Mapeamento das colunas

| Coluna em `Vendas Porsche.xlsx` | Coluna correspondente na sanitizada | O que mudou |
| --- | --- | --- |
| `ID_Vendas` | `sale_id` | Nome traduzido/padronizado. Identificador mantido como número inteiro. |
| `Data de Venda` | `sale_date` e `SaleDateSanitized` | `sale_date` guarda o valor de entrada; a coluna sanitizada deve conter a data validada em `AAAA-MM-DD`. |
| `Nome Customizador` | `customer_name` | Nome traduzido/padronizado; texto mantido. |
| `Modelos` | `porsche_model` e `PorscheModelSanitized` | O modelo original é mantido e uma versão com grafia/capitalização padronizada é adicionada. |
| `Ano do Modelo` | `model_year` e `ModelYearSanitized` | O valor de entrada é mantido e outra coluna converte o ano para número inteiro validado. |
| `Vendas ($)` | `sale_price` e `SalesPriceSanitized` | O preço de entrada é mantido e outra coluna o converte para valor numérico em dólares. |
| `Vendas (R$)` | Sem coluna correspondente | O valor em reais não aparece como campo separado na planilha sanitizada. |
| `Autonomia em Milhas` e `Autonomia Kilometros` | `vehicle_mileage` e `VehicleMileageSanitized` | As duas medidas aparecem consolidadas em uma coluna de quilometragem; a versão tratada usa milhas inteiras. |
| `Tipo de Pagamento` | `payment_method` e `PayMethodSanitized` | O texto original é mantido e uma versão com forma de pagamento padronizada é adicionada. |
| `Cidade` | `city` e `CitySanitized` | O texto original é mantido e a cidade tratada recebe capitalização consistente. |
| `Estado` | `state` e `StateSanitized` | O texto original é mantido e o estado tratado é representado pela sigla oficial em maiúsculas. |
| `Vendedor` | `salesperson` | Nome traduzido/padronizado; texto mantido, sem coluna sanitizada adicional. |
| `Status da Entrega` | `delivery_status` e `DeliveryStatusSanitized` | O texto original é mantido e o status tratado é padronizado. |

Os nomes em inglês são os nomes usados na estrutura da tabela sanitizada; isso não significa que os cabeçalhos do arquivo de origem tenham sido alterados. As colunas terminadas em `Sanitized` são campos adicionais. `sale_id`, `customer_name` e `salesperson` não têm uma coluna sanitizada correspondente. `Vendas (R$)` também não foi preservada como coluna separada.


### Dicionário e tipos de dados

Os tipos abaixo descrevem o valor esperado nas colunas da tabela sanitizada. Em colunas numéricas, o texto `INVALID` pode aparecer no lugar do número quando um valor não passa pela validação; nesses casos, a coluna contém valores de tipos diferentes no arquivo Excel.

| Coluna | Tipo esperado | Descrição e tratamento |
| --- | --- | --- |
| `sale_id` | Inteiro | Identificador do registro de venda. |
| `sale_date` | Data ou texto original | Data como recebida na tabela de entrada. Consulte `SaleDateSanitized` para o valor validado. |
| `SaleDateSanitized` | Texto | Data válida no padrão `AAAA-MM-DD`; datas ausentes ou inválidas recebem `INVALID`. |
| `customer_name` | Texto | Nome associado à venda, mantido sem tratamento adicional. |
| `porsche_model` | Texto | Modelo como recebido na tabela de entrada. |
| `PorscheModelSanitized` | Texto | Modelo com grafia/capitalização padronizada para facilitar agrupamentos. |
| `model_year` | Número ou texto original | Ano-modelo como recebido na tabela de entrada. |
| `ModelYearSanitized` | Inteiro ou `INVALID` | Ano convertido para número inteiro e validado no intervalo de 1990 a 2035. |
| `sale_price` | Número ou texto original | Preço como recebido na tabela de entrada, podendo conter moeda e separadores. |
| `SalesPriceSanitized` | Número decimal ou `INVALID` | Preço em dólares, convertido para número com precisão de duas casas decimais. Valores ausentes, inválidos ou não positivos recebem `INVALID`. |
| `vehicle_mileage` | Número ou texto original | Quilometragem/odometria como recebida, podendo trazer unidade ou texto. |
| `VehicleMileageSanitized` | Inteiro ou `INVALID` | Quilometragem em milhas inteiras. Valores identificados em quilômetros são convertidos para milhas. |
| `payment_method` | Texto | Forma de pagamento como recebida na tabela de entrada. |
| `PayMethodSanitized` | Texto ou `INVALID` | Forma de pagamento com rótulo padronizado. |
| `city` | Texto | Cidade como recebida na tabela de entrada. |
| `CitySanitized` | Texto ou `INVALID` | Espaços ajustados e iniciais das palavras em maiúsculas. |
| `state` | Texto | Estado como recebido na tabela de entrada. |
| `StateSanitized` | Texto ou `INVALID` | Sigla oficial do estado em duas letras maiúsculas. |
| `salesperson` | Texto | Nome ou identificação do vendedor, mantido como recebido. |
| `delivery_status` | Texto | Status de entrega como recebido na tabela de entrada. |
| `DeliveryStatusSanitized` | Texto ou `INVALID` | Status padronizado segundo os valores aceitos pelo schema. |

## Tratamentos realizados

- Criação de um agente para fazer a Sanitização nas colunas;
- **Datas:** interpretação de formatos diferentes, validação do calendário e conversão para `AAAA-MM-DD`;
- **Ano-modelo:** conversão de números e anos escritos em formatos alternativos para inteiro; valores fora do intervalo aceito são marcados como `INVALID`;
- **Preços:** remoção de símbolos/moedas e interpretação de separadores decimais e de milhar; o resultado é um valor numérico em dólares;
- **Quilometragem:** interpretação de valores com unidades ou texto, conversão de quilômetros para milhas quando a unidade é identificada e arredondamento para milha inteira;
- **Modelo, pagamento e status:** padronização de grafia, capitalização e rótulos para permitir agrupamentos consistentes;
- **Cidade e estado:** capitalização consistente para cidades e conversão dos estados para siglas oficiais em maiúsculas;
- **Valores inválidos:** uso de `INVALID` quando o valor está ausente, não pode ser interpretado ou não passa pela validação da coluna;

## Perguntas de negócio

1. **Quais modelos têm maior destaque em cada cidade?** A análise relaciona os modelos mais vendidos aos locais de melhor desempenho para revelar preferências regionais.
2. **Como o volume de vendas evolui ao longo do tempo?** A visualização mensal, trimestral e anual ajuda a identificar períodos de maior ou menor atividade.
3. **Quais estados concentram mais vendas e receita?** A comparação entre quantidade de vendas e faturamento ajuda a avaliar o desempenho dos mercados estaduais.

## Prompts usados na criação e no refinamento

A dashboard foi criada e refinada com auxílio do ChatGPT e do recurso Canvas. A sequência abaixo registra o pedido inicial e os ajustes solicitados durante o desenvolvimento.

### Prompt 1: dashboard inicial

> Utilizando o recurso Canvas, renderize um dashboard em HTML ao lado.
>
> **Filtros**
> - Modelos
> - Ano do Modelo
> - Cidade
> - Tipo de Pagamento
>
> **Perguntas de negócio e KPIs**
> - Quais são os modelos mais vendidos por cidade?
> - Qual modelo é mais vendido por período?

### Prompt 2: perguntas de negócio e identidade visual

> Faça algumas adaptações. Em relação às perguntas de negócio e aos KPIs, quero insights sobre os modelos mais populares em cada cidade. Para a interface e a experiência de uso (UI/UX), use como referência o site oficial da Porsche Brasil: https://www.porsche.com/brazil/pt/. Mantenha um visual elegante e refinado.

### Prompt 3: carro campeão de vendas

> Crie uma seção que mostre o carro campeão de vendas, sua descrição, preço e o local onde ele foi mais vendido. Inclua uma foto do modelo. A seção deve ser dinâmica: se os dados da tabela mudarem, as informações exibidas no site também devem ser atualizadas automaticamente.

### Prompt 4: vendas por estado

> Crie uma seção com um gráfico de barras que mostre a quantidade de vendas por estado e os respectivos valores, usando a coluna Estado.

### Prompt 5: receita por estado

> Adicione ao gráfico de barras uma variável com a receita por estado.

O primeiro prompt definiu a estrutura inicial, os filtros e as perguntas de negócio. Nos refinamentos seguintes, foram detalhados os insights por cidade e a referência visual, incluído o destaque dinâmico do carro campeão e ampliada a comparação estadual para considerar também a receita.

Este projeto foi uma oportunidade de praticar a análise e o tratamento de dados, desenvolver um agente em Python para padronizar campos e criar um dashboard interativo em HTML para apresentar os resultados. A comparação entre as planilhas também revelou divergências que precisam ser revisadas antes de considerar todos os dados sanitizados equivalentes à base original.

