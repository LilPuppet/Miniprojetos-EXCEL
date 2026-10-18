# Miniprojetos EXCEL

Este repositório reúne projetos práticos desenvolvidos no Excel como atividades da DIO para organizar informações, automatizar cálculos e apoiar tarefas do dia a dia.

## Projetos

## 1. Simulador de Investimentos em Fundos Imobiliários (FIIs)

Planilha educacional em Excel para simular aportes mensais, estimar patrimônio futuro e visualizar uma divisão ilustrativa do aporte entre seis categorias de fundos imobiliários.

**Arquivo:** `SimuladorFII.xlsx` — contém a tela principal (`APP`) e a tabela auxiliar (`Planilha2`).

### O que a ferramenta responde

Na aba **APP**, preencha as células destacadas:

- **Salário:** usado para calcular a sugestão de investimento.
- **Rendimento da carteira:** taxa mensal hipotética usada para estimar os dividendos.
- **Aporte mensal:** valor investido a cada mês.
- **Prazo em anos:** horizonte da simulação.
- **Taxa de rendimento mensal:** taxa hipotética usada na projeção do patrimônio.
- **Perfil:** Conservador, Moderado ou Agressivo.

Os resultados incluem patrimônio acumulado, dividendos mensais estimados, cenários para 2, 5, 10, 20 e 30 anos e divisão do aporte entre seis tipos de FII, com gráfico de pizza.

### Como os cálculos funcionam

#### Função VF

A fórmula principal usa a função de valor futuro:

```excel
=VF(taxa_mensal; qtd_anos*12; aporte*-1)
```

No formato interno do arquivo `.xlsx`, a função pode aparecer como `FV`:

```excel
=FV(taxa_mensal,qtd_anos*12,aporte*-1)
```

- `taxa_mensal`: taxa de retorno mensal hipotética.
- `qtd_anos*12`: quantidade de períodos mensais.
- `aporte*-1`: aporte periódico tratado como saída de caixa, para retornar um valor futuro positivo.

A fórmula considera aportes regulares no fim de cada período e uma taxa constante. Não considera impostos, inflação, variação no preço das cotas, vacância, taxas ou mudanças nos rendimentos distribuídos. É uma simplificação didática, não uma previsão ou garantia de retorno.

Os dividendos mensais estimados são calculados como:

```excel
=patrimonio*rendimento_carteira
```

#### PROCV e chave composta

Na aba **Planilha2**, a coluna `CHAVE` concatena o perfil e o tipo de FII, por exemplo:

```excel
=B3&"-"&C3
```

Na aba `APP`, a fórmula procura a chave correspondente ao perfil selecionado e à categoria da linha:

```excel
=VLOOKUP($C$32&"-"&B36,Planilha2!$A:$D,4,FALSE)
```

O `PROCV` (chamado `VLOOKUP` no formato interno do `.xlsx`) procura a chave na primeira coluna do intervalo e retorna o percentual da quarta coluna, exigindo correspondência exata (`FALSE`). O valor mensal de cada categoria é o percentual multiplicado pelo aporte total.

### Intervalos nomeados

| Nome | Referência na aba APP | Finalidade |
|---|---|---|
| `aporte` | `APP!$D$17` | Aporte mensal informado |
| `qtd_anos` | `APP!$D$18` | Prazo em anos |
| `taxa_mensal` | `APP!$D$19` | Taxa mensal da projeção |
| `patrimonio` | `APP!$D$20` | Patrimônio futuro calculado |
| `salario` | `APP!$D$12` | Salário usado na sugestão de aporte |
| `rendimento_carteira` | `APP!$D$13` | Taxa mensal hipotética para dividendos |
| `sugestao_investimento` | `APP!$D$14` | Sugestão calculada como 30% do salário |

### Perfis e percentuais ilustrativos

Os percentuais ficam na aba `Planilha2` e são exemplos didáticos, não recomendações financeiras.

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---:|---:|---:|
| PAPEL | 30% | 32% | 50% |
| TIJOLO | 50% | 35% | 10% |
| HÍBRIDOS | 10% | 8% | 5% |
| FOFs | 10% | 5% | 5% |
| DESENVOLVIMENTO | 0% | 10% | 20% |
| HOTELARIAS | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

Os nomes dos perfis servem para demonstrar validação de dados e busca na tabela. Não significam que a alocação seja adequada para qualquer pessoa.

---

## 2. Organizador de Imposto de Renda

**Arquivo:** `ImpostoRenda.xlsx`.

Planilha de apoio para organizar dados relacionados à preparação do Imposto de Renda, com abas para informações do titular, informes de rendimentos bancários, notas/entradas e uma lista auxiliar de bancos.

### Estrutura da planilha

- **Titular:** área para registrar dados cadastrais do contribuinte.
- **Informes:** área para organizar informes de rendimentos bancários, com banco, valor e referência de anexo; inclui um campo de totalização.
- **Notas:** área para registrar entradas com data, categoria e valor.
- **Bancos:** lista auxiliar de instituições financeiras usada como referência para preenchimento.

### Como utilizar

1. Abra `ImpostoRenda.xlsx` no Excel ou em um aplicativo compatível.
2. Revise os campos de cada aba e preencha apenas com dados corretos e atualizados.
3. Confira os totais calculados e compare-os com os documentos originais.
4. Antes de usar as informações em uma declaração real, valide os dados e as regras aplicáveis em fontes oficiais ou com um profissional de contabilidade.

**Privacidade:** antes de publicar o arquivo no GitHub, substitua todos os dados pessoais, números de documentos, nomes, endereços, valores e anexos por exemplos fictícios. Não publique informações reais de contribuintes.

**Aviso:** a planilha é uma ferramenta de organização e não substitui as orientações oficiais da Receita Federal nem a avaliação de um profissional de contabilidade.

---

## 3. Dashboard de Vendas XBOX

**Arquivo:** `DashboardVendasXBOX.xlsx`.

Dashboard desenvolvido no Excel para consolidar e visualizar informações de vendas relacionadas ao XBOX, facilitando a análise dos dados por meio de indicadores, gráficos e filtros interativos.

### Funcionalidades

- **Visão geral:** apresentação consolidada das informações disponíveis na base de vendas.
- **Filtro por tipo de inscrição:** permite selecionar uma categoria de inscrição para visualizar os dados correspondentes.
- **Indicadores dinâmicos:** atualização das informações apresentadas de acordo com o filtro selecionado.
- **Gráficos interativos:** visualização dos dados agregados para facilitar a identificação de padrões e a comparação entre categorias.
- **Análise consolidada:** possibilidade de retornar à visão geral para considerar todos os registros da base.

### Como utilizar

1. Abra `DashboardVendasXBOX.xlsx` no Excel ou em um aplicativo compatível.
2. Acesse a aba **Dashboard Final**.
3. Utilize o filtro **Tipo de inscrição** para selecionar a categoria desejada.
4. Analise os indicadores e gráficos atualizados conforme a seleção.
5. Retorne à opção **(Todos)** para visualizar os dados consolidados, quando disponível.

O projeto demonstra a aplicação do Excel na construção de dashboards, na consolidação de dados e na apresentação visual de informações para apoiar análises e tomadas de decisão.

---

## Arquivos do repositório

- `SimuladorFII.xlsx` — simulador educacional de investimentos em fundos imobiliários.
- `ImpostoRenda.xlsx` — planilha de organização relacionada ao Imposto de Renda.
- `DashboardVendasXBOX.xlsx` — dashboard interativo para análise consolidada de vendas relacionadas ao XBOX.
- `README.md` — documentação dos projetos.

## Aviso geral

Os projetos têm finalidade educacional e informativa. As taxas e alocações do simulador de FIIs são hipóteses ilustrativas, não recomendações de investimento nem garantia de resultados. A planilha de Imposto de Renda não substitui a conferência dos documentos nem as orientações oficiais. Os dashboards e demais ferramentas devem ser utilizados considerando a qualidade, a atualização e as limitações dos dados de origem.
