# Especificação Técnica — Liga Padel Vale do Iguaçu

## Objetivo

Este documento consolida a arquitetura, regras de negócio, dependências, fórmulas, triggers, validações e casos de borda da planilha oficial da Liga, com foco em permitir sua reconstrução como software fora do Excel.

A referência funcional é a versão final utilizada nos testes:

`Backup_LIGA_OFICIAL_INICIO3_PRONTOPARA4_25-08-2026(6).xlsm`

> Observação importante: alguns textos explicativos da aba `INÍCIO` são legados e não refletem totalmente a lógica final. Para uma reconstrução em software, a fonte de verdade deve ser: regras testadas + fórmulas atuais + VBA utilizado na versão final.

---

# 1. Inventário de abas / camadas

A planilha final possui **15 abas**, todas visíveis.

| Aba | Tipo | Propósito |
|---|---|---|
| `INÍCIO` | Documentação | Explica o uso geral, regras dos jogos e conceitos; contém alguns textos legados/desatualizados. |
| `CADASTRO DE ATLETAS` | Fonte primária + motor auxiliar | Cadastro mestre dos atletas, status, grupo, ordem de entrada e controle dos novos grupos/fila de espera. |
| `CONTROLE DE CORINGAS` | Fonte primária + auxiliar | Registra até 3 ausências/coringas por grupo e ciclo. |
| `GRUPOS` | Derivada | Exibe os grupos atuais e grupos de espera a partir do Cadastro. |
| `LANÇAR RESULTADOS` | Fonte primária parcial + derivada | Gera automaticamente os confrontos; usuário informa essencialmente `Games A` e `Games B`. |
| `CLASSIFICAÇÃO` | Derivada + input de desempate | Calcula classificação individual dentro de cada grupo; permite TB manual. |
| `RANKING PESSOAL` | Derivada | Ranking acumulado usando histórico + ciclo atual. |
| `SUBIDAS E DESCIDAS` | Derivada + override manual | Calcula movimentação de grupo e permite decisão manual diferente da automática. |
| `FIM DE CICLO` | Derivada + confirmação manual | Valida se o ciclo pode ser fechado e contém confirmação manual do fechamento. |
| `PRÓXIMA RODADA` | Derivada + controle manual de cascata | Monta os grupos seguintes e aplica o efeito cascata quando atletas ficam inativos. |
| `HISTÓRICO DOS CICLOS` | Persistência / arquivo histórico | Armazena resultados encerrados de ciclos anteriores e alimenta o ranking acumulado. |
| `HISTÓRICO CICLO 3` | Placeholder | Existe, mas na versão analisada está vazia. |
| `RANKING APÓS CICLO 3` | Placeholder | Existe, mas na versão analisada está vazia. |
| `PREPARAÇÃO CICLO 4` | Derivada + input de continuidade | Prepara a distribuição do ciclo 4 e remove inativos pela cascata. |
| `PRÉVIA CICLO 4` | Derivada + confirmação | Exibe os grupos pré-ciclo 4 e verifica se há grupos incompletos antes da virada. |

## Ordem principal de dependência

```text
CADASTRO DE ATLETAS
        ↓
      GRUPOS
        ↓
LANÇAR RESULTADOS ← CONTROLE DE CORINGAS
        ↓
   CLASSIFICAÇÃO
      ↙       ↘
RANKING       SUBIDAS E DESCIDAS
  ↑                 ↓
HISTÓRICO      FIM DE CICLO
                    ↓
          PRÓXIMA RODADA
                    ↓
          PREPARAÇÃO CICLO 4
                    ↓
             PRÉVIA CICLO 4
```

Dependências detectadas diretamente pelas fórmulas:

- `CONTROLE DE CORINGAS` ← `GRUPOS`, `HISTÓRICO DOS CICLOS`
- `GRUPOS` ← `CADASTRO DE ATLETAS`
- `LANÇAR RESULTADOS` ← `CADASTRO DE ATLETAS`, `CONTROLE DE CORINGAS`, `GRUPOS`, `HISTÓRICO DOS CICLOS`
- `CLASSIFICAÇÃO` ← `CADASTRO DE ATLETAS`, `LANÇAR RESULTADOS`
- `RANKING PESSOAL` ← `CADASTRO DE ATLETAS`, `CLASSIFICAÇÃO`, `HISTÓRICO DOS CICLOS`
- `SUBIDAS E DESCIDAS` ← `CADASTRO DE ATLETAS`, `CLASSIFICAÇÃO`
- `FIM DE CICLO` ← `CADASTRO DE ATLETAS`, `LANÇAR RESULTADOS`, `RANKING PESSOAL`, `SUBIDAS E DESCIDAS`
- `PRÓXIMA RODADA` ← `CADASTRO DE ATLETAS`, `FIM DE CICLO`, `SUBIDAS E DESCIDAS`
- `PREPARAÇÃO CICLO 4` ← `CADASTRO DE ATLETAS`, `FIM DE CICLO`, `SUBIDAS E DESCIDAS`
- `PRÉVIA CICLO 4` ← `PREPARAÇÃO CICLO 4`

---

# 2. Schema de cada aba

## 2.1 `CADASTRO DE ATLETAS`

Esta é a principal tabela mestre do sistema.

| Coluna | Campo | Tipo | Entrada |
|---|---|---|---|
| A | ID ÚNICO | Texto (`ATL-n`) | Manual / macro |
| B | NOME DO ATLETA | Texto | Manual |
| C | CATEGORIA | Texto | Manual |
| D | GRUPO ATUAL | Fórmula | Calculado |
| E | STATUS | Texto | Manual |
| F | DATA DE CADASTRO | Data | Manual |
| G | OBSERVAÇÕES | Texto | Manual |
| H | GRUPO AUTOMÁTICO | Fórmula | Calculado |
| I | GRUPO MANUAL (opcional) | Texto | Manual / override |
| J | — | vazio | Separador |
| K | VAGA (aux.) | Número/fórmula | Calculado |
| L | rótulo `PRÓXIMO ID SUGERIDO` | Texto | Fixo |
| M | Próximo ID sugerido | Fórmula | Calculado |
| N | CHAVE (aux.) | Fórmula | Calculado |
| O | ORDEM DE ENTRADA | Número fixo | Preenchida pela macro |
| P:T | Controle dos novos grupos | Misto | Fórmulas + decisões manuais |
| U | — | sem papel principal confirmado | — |
| V | CATEGORIA da fila | Fórmula | Auxiliar |
| W | GRUPO DE ESPERA | Fórmula | Auxiliar |
| X | VAGA `1/4...4/4` | Fórmula | Auxiliar |
| Y | ATLETA aguardando | Fórmula | Auxiliar |
| Z | CHAVE ESPERA | Fórmula | Auxiliar |
| AA | ID ESPERA | Fórmula | Auxiliar |

Valores atuais de `CATEGORIA` usados pelo motor:

```text
Masculino
Feminino
```

### Controle P:T

Há dois blocos:

- Masculino: candidatos posteriores aos grupos-base, atualmente Q–Z.
- Feminino: candidatos posteriores aos grupos-base, atualmente D–L.

Campos conceituais:

```text
Grupo candidato
Atletas no lote
Situação
INCLUIR NO CICLO ATUAL?  → SIM / NÃO
```

Também existe `Q3` como data limite de entrada, com status de janela calculado.

---

## 2.2 `CONTROLE DE CORINGAS`

| Coluna | Campo | Tipo | Entrada |
|---|---|---|---|
| A | Categoria | Texto | Estrutural |
| B | Grupo | Texto | Estrutural |
| C | Atleta ausente 1 | Texto/dropdown | Manual |
| D | Coringa 1 | Texto | Manual |
| E | Ciclo | Fórmula | Calculado |
| F | Situação | Fórmula | Calculado |
| G | Atleta ausente 2 | Texto/dropdown | Manual |
| H | Coringa 2 | Texto | Manual |
| I | Atleta ausente 3 | Texto/dropdown | Manual |
| J | Coringa 3 | Texto | Manual |
| K | Separador | — | — |
| L:AW | Listas dos atletas de cada grupo | Fórmulas | Auxiliares ocultas |

Regra de negócio confirmada: máximo de **3 substituições/coringas por grupo**.

O coringa participa do jogo, mas:

- não entra na classificação;
- não entra no ranking;
- o atleta ausente não recebe estatísticas dos jogos em que foi substituído.

---

## 2.3 `GRUPOS`

Não é uma tabela única. É formada por blocos repetidos.

Schema lógico de cada grupo:

| Campo | Tipo | Entrada |
|---|---|---|
| Vaga | 1–4 | Estrutural |
| ID | Fórmula | Calculado |
| Atleta | Fórmula | Calculado |
| Status | Fórmula | Calculado |
| Observação | Fórmula | Calculado |

Não deve ser editada pelo operador.

Há grupos oficiais e grupos futuros/de espera.

---

## 2.4 `LANÇAR RESULTADOS`

Schema por partida:

| Coluna | Campo | Tipo | Entrada |
|---|---|---|---|
| A | Categoria | Texto/fórmula | Calculado |
| B | Grupo | Texto/fórmula | Calculado |
| C | Rodada | Número 1–3 | Estrutural |
| D | Dupla A – atleta 1 | Fórmula | Calculado |
| E | Dupla A – atleta 2 | Fórmula | Calculado |
| F | Games A | Número | Manual |
| G | Games B | Número | Manual |
| H | Dupla B – atleta 1 | Fórmula | Calculado |
| I | Dupla B – atleta 2 | Fórmula | Calculado |
| J | Situação | Fórmula | Calculado |
| K | TB (opcional) | Dropdown | Manual |
| L | Separador | — | — |
| M:P | Participantes auxiliares | Fórmula | Auxiliar |

O usuário normal deve alterar essencialmente **F e G**.

---

## 2.5 `CLASSIFICAÇÃO`

Schema visível por atleta:

| Campo | Tipo |
|---|---|
| Posição | Fórmula |
| Atleta | Fórmula |
| Jogos | Fórmula |
| Vitórias | Fórmula |
| Derrotas | Fórmula |
| Games Pró | Fórmula |
| Games Contra | Fórmula |
| Saldo | Fórmula |
| Pontos | Fórmula |

`Pontos`, na implementação atual, é essencialmente o número de vitórias.

Há um segundo bloco manual:

```text
TB
1º TB
2º TB
3º TB
```

Nas colunas auxiliares `AB:AI` ficam os dados usados pelo motor de ordenação:

```text
AB Atleta
AC Jogos
AD Vitórias
AE Derrotas
AF Games Pró
AG Games Contra
AH Saldo
AI Score/Chave de ordenação
```

---

## 2.6 `RANKING PESSOAL`

Campos exibidos:

```text
Posição
Atleta
Categoria
Grupo
Jogos
Vitórias
Derrotas
Games Pró
Games Contra
Saldo
Aproveitamento
```

Auxiliar `AA:AM`:

| Coluna | Campo |
|---|---|
| AA | Origem / ID |
| AB | Atleta |
| AC | Categoria |
| AD | Grupo |
| AE | Jogos acumulados |
| AF | Vitórias acumuladas |
| AG | Derrotas acumuladas |
| AH | Games Pró acumulados |
| AI | Games Contra acumulados |
| AJ | Saldo |
| AK | Aproveitamento |
| AL | Posição calculada |
| AM | Chave |

---

## 2.7 `SUBIDAS E DESCIDAS`

| Coluna | Campo | Tipo | Entrada |
|---|---|---|---|
| A | ID ÚNICO | Fórmula | Calculado |
| B | Atleta | Fórmula | Calculado |
| C | Categoria atual | Texto/fórmula | Calculado |
| D | Grupo atual | Texto/fórmula | Calculado |
| E | Posição | Fórmula | Calculado |
| F | Sugestão automática | Fórmula | Calculado |
| G | Decisão manual | Dropdown/texto | Manual opcional |
| H | Nova categoria | Fórmula | Calculado |
| I | Novo grupo final | Fórmula | Calculado |
| J | Observações | Texto | Manual |
| K | Chave próxima rodada | Fórmula | Calculado |

Na versão atual, `Nova Categoria` simplesmente mantém a categoria atual.

---

## 2.8 `FIM DE CICLO`

| Campo | Tipo |
|---|---|
| Verificação | Texto |
| Resultado | Fórmula/número |
| Status | Fórmula `OK/PENDENTE` |
| Orientação | Fórmula/texto |

Verifica:

1. atletas ativos no ciclo;
2. atletas presentes em movimentações;
3. IDs preenchidos;
4. novo grupo final;
5. nova categoria;
6. ranking;
7. partidas pendentes.

`B18` é a confirmação manual `SIM`.

---

## 2.9 `PRÓXIMA RODADA`

Blocos de grupos:

```text
VAGA
ATLETA
CATEGORIA
GRUPO
STATUS
```

Depois há confrontos possíveis R1/R2/R3.

Auxiliar `S:X`:

| Coluna | Campo |
|---|---|
| S | Ordem |
| T | Categoria |
| U | Atleta após subidas/descidas |
| V | Próximo ciclo (`CONTINUA` / `INATIVO`) |
| W | Resultado da cascata |
| X | Ordem válida |

`T2` controla `APLICAR CASCATA? SIM/NÃO`.

---

## 2.10 `HISTÓRICO DOS CICLOS`

| Coluna | Campo |
|---|---|
| A | Ciclo |
| B | Data de encerramento |
| C | Categoria |
| D | Grupo |
| E | Posição |
| F | Atleta |
| G | Jogos |
| H | Vitórias |
| I | Derrotas |
| J | Games Pró |
| K | Games Contra |
| L | Saldo |
| M | Novo Grupo |
| N | ID ÚNICO |

É um **registro persistente append-only**. Os ciclos anteriores não devem ser apagados.

Há ainda:

- `B5`: número do ciclo usado pelo sistema;
- `B6`: data de encerramento.

Nota de migração: os dados históricos antigos contêm categorias como `Avançado Masculino`, enquanto o cadastro atual usa `Masculino/Feminino`. Isso deve ser normalizado ao migrar para banco de dados.

---

## 2.11 `HISTÓRICO CICLO 3`

Na versão final analisada: **vazia**.

Não há schema operacional ativo confirmado.

---

## 2.12 `RANKING APÓS CICLO 3`

Na versão final analisada: **vazia**.

Não há lógica ativa confirmada.

---

## 2.13 `PREPARAÇÃO CICLO 4`

| Campo | Tipo |
|---|---|
| Ordem Base | Número |
| Grupo Base | Texto |
| Atleta | Fórmula |
| ID | Fórmula |
| Próx. Ciclo | Manual (`CONTINUA/INATIVO`) |
| Ordem Válida | Fórmula |
| Ordem Final | Fórmula |
| Atleta após cascata | Fórmula |
| Grupo 4º Ciclo | Fórmula |

Controle lateral:

```text
Ciclo atual
Próximo ciclo
Status
Cascata
```

---

## 2.14 `PRÉVIA CICLO 4`

Possui:

- contagem masculina;
- contagem feminina;
- quantidade de grupos incompletos;
- status geral;
- confirmação de início;
- blocos de visualização por grupo com:
  - vaga;
  - atleta;
  - ID;
  - `VAZIO / COMPLETO / INCOMPLETO`.

---

## 2.15 `INÍCIO`

Não contém dados do domínio. É documentação.

Para um software novo, não deve virar tabela; deve virar documentação/help.

---

## Named ranges

**Não foram encontrados Named Ranges / Defined Names ativos.**

As ligações atuais são feitas por referências diretas, por exemplo:

```text
'CADASTRO DE ATLETAS'!$A$3:$A$300
CLASSIFICAÇÃO!$B$6:$B$200
```

---

# 3. Fórmulas-chave e sua lógica

## Grupo atual

`CADASTRO DE ATLETAS!D3`

```excel
=IF($B3="","",IF($I3<>"",$I3,$H3))
```

Lógica:

```text
se não tem atleta → vazio
senão, se existe Grupo Manual → usa Grupo Manual
senão → usa Grupo Automático
```

---

## Grupo automático + ordem fixa

`CADASTRO DE ATLETAS!H3`

```excel
=IF($B3="","",
 IF($E3<>"ATIVO","AGUARDANDO",
 IF($O3="","AGUARDANDO",
 IF(ROUNDUP($O3/4,0)<=IF($C3="Masculino",$Q$6,$Q$7),
 "GRUPO "&CHAR(64+ROUNDUP($O3/4,0)),
 IF($C3="Masculino",
 IFERROR(IF(INDEX($T$10:$T$19,ROUNDUP($O3/4,0)-$Q$6)="SIM",
 "GRUPO "&CHAR(64+ROUNDUP($O3/4,0)),"AGUARDANDO"),"AGUARDANDO"),
 IFERROR(IF(INDEX($T$24:$T$32,ROUNDUP($O3/4,0)-$Q$7)="SIM",
 "GRUPO "&CHAR(64+ROUNDUP($O3/4,0)),"AGUARDANDO"),"AGUARDANDO")
 )))))
```

Regra:

```text
ordem 1–4   → grupo A
5–8         → B
9–12        → C
...
```

Para grupo além dos grupos-base, o atleta permanece `AGUARDANDO` até o respectivo controle `SIM`.

A **ordem O**, e não a posição física da linha no Excel, é a verdade da fila.

---

## Chave de atleta dentro do grupo

`N3`

```excel
=IF($K3="","",$C3&"|"&$D3&"|"&$K3)
```

Exemplo:

```text
Masculino|GRUPO A|3
```

---

## Grupo de espera

`W5`

```excel
=IF($Y5="","","GRUPO "&CHAR(64+ROUNDUP($O3/4,0)))
```

`X5`

```excel
=IF($Y5="","",RESTO($O3-1,4)+1&"/4")
```

`Y5`

```excel
=IF(AND($B3<>"",$E3="ATIVO",$O3<>"",$D3="AGUARDANDO"),$B3,"")
```

Essas fórmulas transformam a ordem fixa em:

```text
Grupo E
1/4
Nome do atleta
```

---

## Status de um grupo futuro

Exemplo feminino `R24`:

```excel
=IF(Q24="","",
 IF(T24="SIM",
 COUNTIFS($C$3:$C$300,"Feminino",$D$3:$D$300,Q24),
 COUNTIFS($V$5:$V$302,"Feminino",$W$5:$W$302,Q24,$Y$5:$Y$302,"<>")))
```

Antes de `SIM`, conta a fila.

Depois de `SIM`, conta os atletas efetivamente pertencentes ao grupo.

`S24`:

```excel
=IF(Q24="","",
 IF(R24=4,
 IF(OR($Q$3="",TODAY()<=$Q$3),
 IF(T24="SIM","INCLUÍDO NO CICLO","GRUPO COMPLETO — DECIDIR SIM/NÃO"),
 IF(T24="SIM","INCLUÍDO NO CICLO","GRUPO COMPLETO — PRÓXIMO CICLO")),
 "AGUARDANDO "&R24&"/4"))
```

Estados possíveis:

```text
AGUARDANDO 0/4
AGUARDANDO 1/4
...
GRUPO COMPLETO — DECIDIR SIM/NÃO
INCLUÍDO NO CICLO
GRUPO COMPLETO — PRÓXIMO CICLO
```

---

## Busca do atleta de um grupo

Exemplo `GRUPOS!B8`:

```excel
=IFERROR(
 INDEX('CADASTRO DE ATLETAS'!$A:$A,
 MATCH("Masculino|GRUPO A|"&1,
 'CADASTRO DE ATLETAS'!$N:$N,0)),
"")
```

Usa chave composta, e não posição física da linha.

---

## Coringa

`CONTROLE DE CORINGAS!E5`

```excel
=IF(AND(C5="",G5="",I5=""),"",
'HISTÓRICO DOS CICLOS'!$B$5)
```

Situação:

```excel
=IF(AND(C5="",G5="",I5=""),"SEM CORINGA",
 IF(E5<>'HISTÓRICO DOS CICLOS'!$B$5,"CICLO ANTERIOR",
 IF(OR(
 AND(C5<>"",D5=""),
 AND(G5<>"",H5=""),
 AND(I5<>"",J5="")
 ),"PREENCHER CORINGA","ATIVO NESTE CICLO")))
```

---

## Situação de uma partida

`LANÇAR RESULTADOS!J6`

```excel
=IF(
 COUNTIF(D6:E6,"?*")+COUNTIF(H6:I6,"?*")=0,
 "",
 IF(COUNT(F6:G6)=2,"CONCLUÍDO","PENDENTE")
)
```

Uma partida só é concluída quando os dois placares existem.

---

## Classificação do grupo

Critérios reais da classificação:

```text
1. Vitórias
2. Games Pró
3. Saldo de Games
4. TB manual
5. Vaga original como fallback técnico
```

Score auxiliar em `AI`:

```excel
=IF($AB2="","",
 $AD2*100000000
 +$AF2*100000
 +($AH2+100)*100
 +IF($AB2=$K$4,30,
   IF($AB2=$L$4,20,
     IF($AB2=$M$4,10,0)))
 +(5-1))
```

Posição final:

```excel
=IFERROR(
 INDEX($AB$2:$AB$5,
 MATCH(LARGE($AI$2:$AI$5,1),
 $AI$2:$AI$5,0)),
"")
```

---

## Ranking geral

As estatísticas acumuladas usam:

```text
Histórico dos ciclos
+
Classificação do ciclo atual
```

Exemplo de jogos:

```excel
=IF($AA2="","",
 SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA2,
       'HISTÓRICO DOS CICLOS'!$G$9:$G$10000)
 +
 SUMIF(CLASSIFICAÇÃO!$B$6:$B$200,$AB2,
       CLASSIFICAÇÃO!$C$6:$C$200))
```

Posição geral em `AL2`:

```excel
=IF($AA2="","",
1
+COUNTIF($AF$2:$AF$299,">"&$AF2)
+COUNTIFS($AF$2:$AF$299,$AF2,$AH$2:$AH$299,">"&$AH2)
+COUNTIFS($AF$2:$AF$299,$AF2,$AH$2:$AH$299,$AH2,$AI$2:$AI$299,"<"&$AI2)
+COUNTIFS($AF$2:$AF$299,$AF2,$AH$2:$AH$299,$AH2,$AI$2:$AI$299,$AI2,$AJ$2:$AJ$299,">"&$AJ2)
+COUNTIFS($AF$2:$AF$299,$AF2,$AH$2:$AH$299,$AH2,$AI$2:$AI$299,$AI2,$AJ$2:$AJ$299,$AJ2,$AB$2:$AB$299,"<"&$AB2))
```

Critérios:

```text
1. Mais vitórias
2. Mais Games Pró
3. Menos Games Contra
4. Maior saldo
5. Nome em ordem alfabética
```

---

## Subidas e descidas

Regra-base:

```text
1º → sobe 2 grupos
2º → sobe 1 grupo
3º → desce 1 grupo
4º → desce 2 grupos
```

Exceções de borda confirmadas:

```text
Grupo A:
 posição 1 → mantém
 posição 2 → mantém

Grupo B:
 posição 2 → mantém

Penúltimo grupo:
 posição 3 → mantém

Último grupo:
 posição 3 → mantém
 posição 4 → mantém
```

`I6` respeita override manual primeiro:

```excel
=IF(B6="","",
 IF(G6<>"",G6,
 ...grupo calculado automaticamente...))
```

`K6` produz chave da próxima rodada:

```excel
=IF(I6="","",
 H6&"|"&I6&"|"&
 COUNTIFS($H$6:H6,H6,$I$6:I6,I6))
```

---

## Fechamento

Jogos pendentes:

```excel
=COUNTIFS('LANÇAR RESULTADOS'!$A$5:$A$268,"Masculino",
          'LANÇAR RESULTADOS'!$J$5:$J$268,"PENDENTE")
+
 COUNTIFS('LANÇAR RESULTADOS'!$A$5:$A$268,"Feminino",
          'LANÇAR RESULTADOS'!$J$5:$J$268,"PENDENTE")
```

Status geral:

```excel
=IF(COUNTIF(C5:C11,"PENDENTE")=0,
"PRONTO PARA GERAR A PRÓXIMA RODADA",
"EXISTEM PENDÊNCIAS — REVISE ANTES DE CONTINUAR")
```

Confirmação:

```excel
=IF(UPPER(B18)="SIM",
 IF(A14="PRONTO PARA GERAR A PRÓXIMA RODADA",
 "CONFIRMADO — CONSULTE A ABA PRÓXIMA RODADA",
 "NÃO FOI POSSÍVEL CONFIRMAR — EXISTEM PENDÊNCIAS"),
"MANTENHA NÃO OU DIGITE SIM APÓS REVISAR O CICLO")
```

---

# 4. Triggers — o que dispara o quê

## Trigger VBA automático confirmado

Na planilha `Planilha12 (CADASTRO DE ATLETAS)` existe:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
```

Código utilizado na versão final:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range
    Dim cel As Range
    Dim linha As Long
    Dim i As Long
    Dim categoria As String
    Dim maiorOrdem As Long

    On Error GoTo Saida

    Set rng = Intersect(Target, _
        Union(Me.Range("B3:C300"), Me.Range("E3:E300")))

    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False

    For Each cel In rng.Cells

        linha = cel.Row

        If Trim(Me.Cells(linha, "B").Value) <> "" _
        And UCase(Trim(Me.Cells(linha, "E").Value)) = "ATIVO" _
        And Me.Cells(linha, "O").Value = "" Then

            categoria = Trim(Me.Cells(linha, "C").Value)

            If categoria = "Masculino" Or categoria = "Feminino" Then

                maiorOrdem = 0

                For i = 3 To 300

                    If Trim(Me.Cells(i, "C").Value) = categoria Then

                        If IsNumeric(Me.Cells(i, "O").Value) Then

                            If CLng(Me.Cells(i, "O").Value) > maiorOrdem Then
                                maiorOrdem = CLng(Me.Cells(i, "O").Value)
                            End If

                        End If

                    End If

                Next i

                Me.Cells(linha, "O").Value = maiorOrdem + 1

            End If

        End If

    Next cel

Saida:
    Application.EnableEvents = True

End Sub
```

### Trigger

Mudança em:

```text
B3:C300
ou
E3:E300
```

Se:

```text
Nome preenchido
Status = ATIVO
Ordem O vazia
Categoria = exatamente Masculino ou Feminino
```

então:

```text
O = maior ordem já existente nessa categoria + 1
```

O valor gravado é fixo.

---

## Cadeia automática de fórmulas

### Ativar atleta

```text
STATUS → ATIVO
       ↓
macro grava ORDEM DE ENTRADA
       ↓
H calcula grupo/AGUARDANDO
       ↓
V:AA montam fila
       ↓
R/S atualizam 1/4...4/4
       ↓
GRUPOS atualiza
```

### Marcar grupo futuro como SIM

```text
T = SIM
 ↓
H deixa de retornar AGUARDANDO
 ↓
D vira GRUPO X
 ↓
GRUPOS mostra grupo oficial
 ↓
LANÇAR RESULTADOS libera confrontos
```

### Registrar coringa

```text
CONTROLE DE CORINGAS
 ↓
LANÇAR RESULTADOS substitui nome nos jogos
 ↓
CLASSIFICAÇÃO continua usando roster oficial
```

### Preencher Games A e Games B

```text
Resultado
 ↓
Situação = CONCLUÍDO
 ↓
CLASSIFICAÇÃO recalcula
 ↓
RANKING recalcula
 ↓
SUBIDAS E DESCIDAS recalcula
 ↓
FIM DE CICLO recalcula pendências
```

### Alterar TB manual

```text
K/L/M da CLASSIFICAÇÃO
 ↓
AI (score) muda
 ↓
posição do grupo muda
 ↓
SUBIDAS/DESCIDAS muda
```

### Override de movimentação

```text
SUBIDAS!G = grupo manual
 ↓
I = novo grupo final manual
 ↓
K = nova chave
 ↓
PRÓXIMA RODADA muda
```

---

## Outros módulos VBA presentes

Módulos/procedures confirmados no projeto:

```text
CicloLiga
HistoricoLiga
IDSeguroManual
ReiniciarLigaOficial
ReiniciarLigaOficialV2
```

Rotinas identificadas:

```text
EncerrarCicloEIniciarProximo
ArquivarCicloNoHistorico
GerarIDsDosAtletas
ReiniciarParaLigaOficial
ReiniciarParaLigaOficialV2
Worksheet_Change
```

Não foi extraído com segurança o código-fonte integral dos cinco módulos padrão. Portanto, seus detalhes internos não devem ser inventados.

O que está confirmado:

- há rotina de encerramento/início de ciclo;
- há rotina de arquivamento no histórico;
- há rotina de geração de IDs;
- há rotinas de reinicialização;
- o `Worksheet_Change` acima é o preenchimento automático da ordem.

Também não foi confirmado com 100% de precisão qual botão chama qual Sub.

---

# 5. Lógica de posicionamento

## Classificação dentro do grupo

Ordem real:

```text
1. Vitórias
2. Games Pró
3. Saldo de Games
4. TB manual
5. Vaga original como fallback técnico
```

Há campos:

```text
1º TB
2º TB
3º TB
```

para resolver empates restantes.

## Ranking geral acumulado

Critérios:

```text
1. Vitórias
2. Games Pró
3. Menos Games Contra
4. Saldo
5. Nome em ordem alfabética
```

Não há TB manual confirmado no ranking geral.

## Efeito da posição

A posição **dentro do grupo** determina a movimentação:

```text
1º → +2 grupos para cima
2º → +1
3º → -1
4º → -2
```

respeitando exceções de borda.

O ranking geral acumulado é uma classificação separada e não é utilizado diretamente para formar os grupos do próximo ciclo.

---

# 6. Lógica de chaveamento / confrontos

A planilha não possui chave eliminatória/bracket.

O “chaveamento” é um rodízio dentro de um grupo de 4.

Com vagas:

```text
1
2
3
4
```

os jogos são:

```text
Rodada 1:
1 + 4  x  2 + 3

Rodada 2:
1 + 3  x  2 + 4

Rodada 3:
1 + 2  x  3 + 4
```

Todos jogam 3 partidas por grupo.

Não existe regra confirmada para:

- oitavas;
- quartas;
- semifinal;
- final;
- bye;
- cabeças de chave eliminatória;
- avanço de vencedor;
- clube;
- evitar adversário repetido;
- impedir repetição de dupla entre ciclos.

Essas regras não devem ser inventadas no software novo.

---

# 7. Validações e restrições

## `CADASTRO DE ATLETAS`

Dropdown de Status:

```text
ATIVO
INATIVO
AFASTADO
```

Limitação atual: a validação está aplicada somente em `E3:E54`, enquanto a macro trabalha até linha 300.

Controles de novos grupos:

```text
T10:T19
T24:T32
```

Dropdown:

```text
NÃO
SIM
```

### Categoria

Não há validação confirmada em C para `Masculino/Feminino`.

Fragilidade atual: a macro exige o texto exatamente:

```text
Masculino
Feminino
```

---

## Coringas

Campos de atleta ausente têm dropdown baseado nos quatro atletas do grupo.

As listas auxiliares `L:AW` estão ocultas.

Não foi confirmada validação equivalente para o nome do coringa/substituto.

---

## Resultados

`K` possui dropdown a partir dos participantes do jogo.

Não foi detectada validação numérica forte em `Games A / Games B`.

O Excel atual não impede, por exemplo:

```text
99 x 50
```

No sistema novo, placar deve possuir validação de domínio.

---

## Classificação

Os campos de TB possuem dropdown com os quatro atletas do grupo.

---

## Subidas e Descidas

Há dropdown em `G6:G69`, mas ele contém apenas:

```text
GRUPO A
GRUPO B
...
GRUPO G
```

Isso não cobre todos os grupos masculinos atuais/futuros.

É uma limitação real da versão Excel.

---

## Próxima Rodada

`T2`:

```text
SIM
NÃO
```

`V6:V157`:

```text
CONTINUA
INATIVO
```

---

## Preparação Ciclo 4

`E5:E68` e `E72:E83`:

```text
CONTINUA
INATIVO
```

---

## Prévia Ciclo 4

Existe validação `SIM/NÃO` associada a `H5`.

Na versão analisada, o valor de confirmação aparece em `H6`.

Isso deve ser tratado como inconsistência da planilha e não copiado literalmente.

---

## Formatação condicional

### Cadastro

- `ATIVO` → verde
- `INATIVO` → vermelho
- `AFASTADO` → amarelo

### Grupos

- ativo → verde
- inativo → vermelho

### Resultados

- `CONCLUÍDO` → verde
- `PENDENTE` → vermelho/alaranjado

### Fim de Ciclo

- `OK` → verde
- `PENDENTE` → vermelho
- status final alterna verde/vermelho

### Ranking

- posição 1 recebe destaque amarelo/dourado;
- coluna de Saldo tem escala de cores.

---

# 8. Casos de borda conhecidos

## 8.1 Atleta antigo ativado depois tomava vaga de atleta novo

Problema original: a lógica usava a posição física da linha no cadastro.

Solução final:

```text
ORDEM DE ENTRADA fixa + macro
```

Exemplo:

```text
D = 13,14,15,16
depois:
nova/reativada = 17
```

Mesmo que a atleta 17 esteja numa linha antiga do Excel, ela não toma uma vaga 13–16.

---

## 8.2 Inativar alguém não reaproveita automaticamente seu número

Isso é proposital.

Se:

```text
Nanda = ordem 13
```

e fica INATIVA, a ordem 13 não é reutilizada por outra atleta.

A sequência segue em frente.

---

## 8.3 Macro depende da categoria escrita exatamente

Valores diferentes como:

```text
masculino
Masculino 
MASC
```

podem falhar.

No software novo: usar enum/FK.

---

## 8.4 Coluna O pode ser alterada manualmente

No Excel ela não está protegida.

Alterar ou duplicar uma ordem pode quebrar a fila.

No software novo, `entryOrder` deve ser imutável para usuário comum.

---

## 8.5 IDs não têm constraint real de unicidade

Há lógica para IDs e validação de prefixo `ATL-`, mas isso não garante unicidade real.

No sistema novo:

```text
athlete_id UNIQUE NOT NULL
```

---

## 8.6 Uso de nome em alguns relacionamentos

Algumas fórmulas ainda usam nome para localizar registros.

Dois atletas com exatamente o mesmo nome podem gerar comportamento incorreto.

No software novo, todos os relacionamentos devem usar ID.

---

## 8.7 Resultado incompleto

Se apenas um dos placares for preenchido:

```text
Situação = PENDENTE
```

e o fechamento do ciclo fica bloqueado.

---

## 8.8 Coringa incompleto

Atleta ausente informado sem substituto:

```text
PREENCHER CORINGA
```

Registro de ciclo anterior:

```text
CICLO ANTERIOR
```

e não vale no ciclo atual.

---

## 8.9 Grupo futuro incompleto

1/4, 2/4 ou 3/4 permanece em espera.

Com 4/4:

```text
DECIDIR SIM/NÃO
```

Somente `SIM` inclui no ciclo atual.

---

## 8.10 Grupo formado não deve ser desmontado por entrada posterior

Regra de domínio confirmada e testada.

---

## 8.11 Problemas históricos já corrigidos

Durante o desenvolvimento ocorreram:

- Grupo Q/coringa não refletindo corretamente na classificação;
- fila avançando Q→R→S incorretamente;
- feminino não aparecendo antes de completar 4;
- quinto atleta indo para grupo errado;
- grupos futuros sem nome visível;
- atleta substituído recebendo pontuação indevida;
- fórmulas experimentais criando referência circular;
- arquivos experimentais gerados que faziam o Excel pedir recuperação.

Esses casos devem virar testes de regressão no software novo.

---

## 8.12 Limites hard-coded inconsistentes

Há ranges fixos diferentes:

```text
Cadastro / nova macro: até linha 300
FIM DE CICLO: cadastro até linha 200
PREPARAÇÃO CICLO 4: cadastro até linha 200
SUBIDAS: estrutura limitada por linhas fixas
Histórico: até linha 10000
Ranking visível: aproximadamente 120 posições
```

Limites de grupos atuais:

```text
Masculino → até Z
Feminino → até L
```

No software novo, não reproduzir esses limites como regra estrutural.

---

## 8.13 `PREPARAÇÃO CICLO 4` está dimensionada para a base antiga

Base explícita:

```text
Masculino: 64 posições / grupos A–P
Feminino: 12 posições / grupos A–C
```

Se novos grupos forem adicionados, essa aba não é verdadeiramente dinâmica para qualquer quantidade de atletas.

No sistema novo, a preparação deve ser calculada a partir do estado real do banco.

---

# 9. Frequência e fluxo de atualização

## Frequência

Os ciclos foram pensados em torno de **15 dias**, mas essa duração não está codificada como regra rígida.

A operação é orientada a eventos:

```text
Cadastro / ativação → quando ocorre
Coringa → antes do jogo afetado
Resultados → conforme os jogos são realizados
Classificação → imediatamente por fórmula
Ranking → imediatamente por fórmula
Fechamento → uma vez ao terminar o ciclo
Cascata → entre um ciclo e o próximo
```

Não existe processo diário automático ou agendamento.

---

## Quem edita

O modelo operacional atual é de **um organizador principal operando a planilha**.

Não foi identificado no Excel:

- controle de usuários;
- permissões;
- lock por atleta/grupo;
- trilha de auditoria;
- histórico de quem alterou o quê;
- edição concorrente estruturada.

Se o sistema novo tiver múltiplos administradores, isso será uma funcionalidade nova.

---

# 10. Modelo de domínio sugerido para reconstrução

```text
Athlete
  id
  name
  category
  status
  registrationDate
  entryOrder

Cycle
  id
  startDate
  endDate
  status

Group
  id
  cycleId
  category
  groupCode
  status

GroupMember
  groupId
  athleteId
  slot

Match
  id
  groupId
  round
  sideAPlayer1Id
  sideAPlayer2Id
  sideBPlayer1Id
  sideBPlayer2Id
  gamesA
  gamesB
  status

Substitution
  id
  cycleId
  groupId
  absentAthleteId
  substituteAthleteIdOrName
  slot

GroupStanding
  athleteId
  cycleId
  groupId
  games
  wins
  losses
  gamesFor
  gamesAgainst
  differential
  position

GlobalRanking
  athleteId
  cumulativeGames
  cumulativeWins
  cumulativeLosses
  cumulativeGamesFor
  cumulativeGamesAgainst
  differential
  position

Movement
  athleteId
  cycleId
  fromGroup
  groupPosition
  suggestedGroup
  manualOverride
  finalGroup

WaitlistEntry
  athleteId
  entryOrder
  targetGroup
  slot

CycleHistory
  immutable snapshot
```

---

# 11. Regras essenciais a preservar

1. Grupos têm 4 atletas.
2. Ordem de entrada é imutável.
3. Grupos já liberados não são reorganizados por entradas posteriores.
4. Grupo futuro só entra com 4/4 + confirmação.
5. Cada grupo possui 3 confrontos fixos.
6. Coringa joga, mas não pontua.
7. Atleta ausente não pontua os jogos substituídos.
8. Classificação do grupo e ranking acumulado são coisas diferentes.
9. Posição no grupo gera movimentação.
10. Movimentação automática aceita override administrativo.
11. Fechamento é bloqueado por pendências.
12. Histórico encerrado deve ser imutável.
13. Relacionamentos no novo sistema devem usar IDs, nunca nomes.
14. Limites fixos de linhas/grupos do Excel não devem ser reproduzidos como limites do sistema.

---

# 12. Pontos não confirmados — não inventar na reconstrução

Não há evidência suficiente para afirmar ou implementar automaticamente:

- chave eliminatória/bracket;
- oitavas/quartas/semi/final;
- byes;
- cabeça de chave eliminatória;
- restrição de clube;
- regra para evitar repetir adversários;
- regra para impedir dupla repetida entre ciclos;
- detalhes internos completos dos módulos VBA:
  - `CicloLiga`
  - `HistoricoLiga`
  - `IDSeguroManual`
  - `ReiniciarLigaOficial`
  - `ReiniciarLigaOficialV2`
- associação exata de cada botão gráfico a cada macro.

Esses pontos devem ser tratados como requisitos futuros, não como comportamento confirmado da planilha atual.

---

# 13. Princípio de migração

A reconstrução não deve ser uma tradução literal de células para código.

Deve preservar as **regras de negócio**, eliminando fragilidades do Excel:

- substituir ranges fixos por tabelas dinâmicas;
- substituir nomes por IDs;
- substituir texto livre por enums;
- tornar `entryOrder` imutável;
- garantir unicidade de IDs;
- registrar auditoria;
- separar estados de ciclo;
- tratar fechamento como transação;
- persistir histórico como snapshot imutável;
- implementar testes de regressão para todos os casos de borda conhecidos.

