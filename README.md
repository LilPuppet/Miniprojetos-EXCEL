## Simulador de Investimentos em Fundos Imobiliários (FIIs)

Planilha educacional em Excel para simular aportes mensais, estimar patrimônio futuro e visualizar uma divisão ilustrativa do aporte entre seis categorias de fundos imobiliários.

### Arquivos

- `simulador_fii.xlsx`: planilha com a tela principal (`APP`) e a tabela auxiliar (`Planilha2`).

### O que a ferramenta responde

Na aba **APP**, preencha as células amarelas/azul-claro (campos editáveis):

- **Salário:** usado para calcular a sugestão de investimento.
- **Rendimento da carteira:** taxa mensal hipotética usada para estimar os dividendos.
- **Aporte mensal:** valor que será investido a cada mês.
- **Prazo em anos:** horizonte da simulação.
- **Taxa de rendimento mensal:** taxa hipotética usada na projeção do patrimônio.
- **Perfil:** Conservador, Moderado ou Agressivo.

Os resultados calculados aparecem em células cinzas, com destaque para:

- **Patrimônio acumulado:** valor futuro estimado ao fim do prazo informado.
- **Dividendos mensais estimados:** patrimônio multiplicado pela taxa mensal de rendimento da carteira.
- **Cenários:** estimativas para 2, 5, 10, 20 e 30 anos.
- **Divisão do aporte:** percentual e valor mensal direcionado a cada um dos seis tipos de FII, além de um gráfico de pizza.

### Como os cálculos funcionam

#### Função VF

A fórmula principal usa a função de valor futuro:

```excel
=VF(taxa_mensal; qtd_anos*12; aporte*-1)
```

No arquivo Excel, as funções são armazenadas em inglês (`FV`) para compatibilidade com o formato `.xlsx`:

```excel
=FV(taxa_mensal,qtd_anos*12,aporte*-1)
```

- `taxa_mensal`: taxa de retorno mensal hipotética.
- `qtd_anos*12`: quantidade de períodos mensais.
- `aporte*-1`: aporte periódico informado como saída de caixa, para retornar um valor futuro positivo.

A fórmula considera aportes regulares no fim de cada período e taxa constante. Não considera impostos, inflação, variação no preço das cotas, vacância, taxas, nem mudanças nos rendimentos distribuídos. Portanto, é uma simplificação didática e não uma previsão ou garantia de retorno.

Os dividendos mensais estimados são calculados como:

```excel
=patrimonio*rendimento_carteira
```

#### PROCV e chave composta

Na aba **Planilha2**, a coluna `CHAVE` concatena o perfil e o tipo de FII, por exemplo:

```excel
=B3&"-"&C3
```

Na aba `APP`, a fórmula busca a chave correspondente ao perfil selecionado e à categoria da linha:

```excel
=VLOOKUP($C$32&"-"&B36,Planilha2!$A:$D,4,FALSE)
```

O `PROCV` (nome em português; `VLOOKUP` no arquivo `.xlsx`) procura a chave na primeira coluna do intervalo e retorna o percentual da quarta coluna, exigindo correspondência exata (`FALSE`). O valor mensal de cada categoria é o percentual multiplicado pelo aporte total.

### Intervalos nomeados

O arquivo usa estes nomes para tornar as fórmulas mais legíveis:

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

Os percentuais ficam na aba `Planilha2`. Eles foram mantidos como **exemplos didáticos da planilha de referência**, não como recomendação financeira.

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---:|---:|---:|
| PAPEL | 30% | 32% | 50% |
| TIJOLO | 50% | 35% | 10% |
| HÍBRIDOS | 10% | 8% | 5% |
| FOFs | 10% | 5% | 5% |
| DESENVOLVIMENTO | 0% | 10% | 20% |
| HOTELARIAS | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

Os rótulos dos perfis servem apenas para demonstrar a validação de dados e a busca da tabela. Não significam que a alocação seja adequada para qualquer pessoa.

### Aviso

Este projeto é exclusivamente educacional. Taxas e percentuais são hipóteses ilustrativas, não constituem recomendação de investimento e não garantem resultados futuros. Fundos imobiliários envolvem riscos, oscilações de preço e rendimentos variáveis.
