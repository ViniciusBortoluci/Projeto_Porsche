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

## Dicionário e tipos de dados

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

## Criação do HTML

Com o auxilio do ChatGPT, foi passado um promtp para criação de um site que:

- Se parecesse com o Site da POrsche BR;
- Indicadores de KPIs;
- Métrica para avaliação de resultados;
- Gráficos de barras indicando as vendas por periodos;
- Gráfico de barras horizontais mostrando os 15 estado com maior venda;
- Os modelos com maior vendas e as cidades;
- Uma tabela de destaque mostrando os modelos com maiores vendas descrescente;


Este projeto foi uma oportunidade de praticar a análise e o tratamento de dados, desenvolver um agente em Python para padronizar campos e criar um dashboard interativo em HTML para apresentar os resultados. A comparação entre as planilhas também revelou algumas divergências que precisam ser revisadas antes de considerar todos os dados sanitizados equivalentes à base original.

