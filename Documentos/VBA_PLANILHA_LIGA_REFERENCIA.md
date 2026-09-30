# Referência VBA — Liga Padel Vale do Iguaçu

## Finalidade

Este documento registra o código VBA extraído diretamente da planilha:

`Anexar-Claude.xlsm`

SHA-256 do arquivo analisado:

`bf84dae87a387b56d3727058b30ff6a07f3cfb72ef722af21adf4e924fffe1ab`

O objetivo é servir como **referência de migração para o novo sistema web/app** e complementar o arquivo `ESPECIFICACAO_TECNICA_LIGA_PADEL.md`.

> **Importante:** presença de uma macro no arquivo não significa que sua lógica ainda representa a versão final do negócio. Alguns módulos permanecem no workbook como legado de fases anteriores da planilha. Este documento separa explicitamente o que é atual, o que é útil como referência e o que não deve ser portado literalmente.

---

# 1. Resumo executivo dos módulos

| Módulo | Rotina principal | Status para migração | Recomendação |
|---|---|---|---|
| `Planilha12` | `Worksheet_Change` | **ATUAL / VALIDADO** | Preservar a regra de negócio da ordem de entrada automática. |
| `IDSeguroManual` | `GerarIDsDosAtletas` | **ÚTIL / PARCIALMENTE ATUAL** | Preservar conceito de ID permanente e unicidade; remover limite de linha 200. |
| `CicloLiga` | `EncerrarCicloEIniciarProximo` | **LEGADO / NÃO PORTAR LITERALMENTE** | Aproveitar validações/transação/rollback como referência; reimplementar sobre o modelo atual. |
| `HistoricoLiga` | `ArquivarCicloNoHistorico` | **LEGADO** | Não usar mapa hard-coded de categorias/grupos; preservar conceito de snapshot histórico. |
| `ReiniciarLigaOficial` | `ReiniciarParaLigaOficial` | **UTILITÁRIO LEGADO E DESTRUTIVO** | Não é regra de negócio do produto; no máximo referência administrativa. |
| `ReiniciarLigaOficialV2` | `ReiniciarParaLigaOficialV2` | **UTILITÁRIO LEGADO, V2 MAIS SEGURO** | Não portar como fluxo normal; se houver reset administrativo, exigir backup/confirmação. |

---

# 2. Descoberta crítica: macros antigas ainda existem no arquivo

A estrutura final da Liga trabalha com categorias gerais:

```text
Masculino
Feminino
```

e grupos masculinos/femininos independentes.

Entretanto, `CicloLiga` e `HistoricoLiga` ainda possuem mapas hard-coded da estrutura anterior, como:

```text
Avançado Masculino
Intermediário Masculino
Iniciante Masculino
Intermediário Feminino
```

Exemplo real encontrado em `CicloLiga`:

```vb
mapa.Add MontarChave("Avançado Masculino", "GRUPO A"), ...
mapa.Add MontarChave("Intermediário Masculino", "GRUPO A"), ...
mapa.Add MontarChave("Iniciante Masculino", "GRUPO A"), ...
mapa.Add MontarChave("Intermediário Feminino", "GRUPO A"), ...
```

Isso demonstra que esses módulos são **remanescente de uma arquitetura anterior**.

### Consequência para o Claude Code

**NÃO** portar literalmente `CicloLiga` ou `HistoricoLiga` para o sistema novo.

O que deve ser aproveitado deles:

- validações antes do fechamento;
- bloqueio de duplicidade;
- ideia de fechamento atômico/transacional;
- snapshot histórico;
- rollback em caso de erro;
- limpeza de placares ao iniciar novo ciclo;
- preservação do ranking acumulado;
- confirmação explícita antes de ações destrutivas.

O que deve ser descartado/reprojetado:

- listas fixas de categorias;
- mapas fixos de grupo → linha;
- limites de linha (`200`, `57`, `93`, `80`);
- dependência de posições físicas de células;
- reconstrução de grupos por endereços fixos.

---

# 3. `Planilha12` — ORDEM DE ENTRADA AUTOMÁTICA

## Status

**ATUAL E VALIDADO NOS TESTES FINAIS.**

`Planilha12` é o codename da aba `CADASTRO DE ATLETAS`.

## Trigger

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
```

Ele observa alterações em:

```text
B3:C300 → Nome / Categoria
E3:E300 → Status
```

## Regra exata

Quando uma linha satisfaz simultaneamente:

```text
Nome preenchido
Status = ATIVO
Coluna O (ORDEM DE ENTRADA) vazia
Categoria = Masculino ou Feminino
```

a macro:

1. percorre todos os atletas da mesma categoria;
2. identifica a maior ordem de entrada existente;
3. grava `maiorOrdem + 1` na coluna O;
4. grava como valor fixo, não fórmula;
5. não altera ordens já existentes.

## Regra de domínio que deve existir no novo sistema

`entryOrder` deve ser:

- atribuído automaticamente quando o atleta entra como ativo;
- sequencial dentro da categoria;
- persistente;
- imutável por mudanças posteriores de status;
- não reutilizado automaticamente quando alguém fica inativo;
- usado para ordenar a fila de novos grupos.

### Testes finais realizados na planilha

Foram validados os cenários:

```text
Feminino:
ordens existentes → sequência continuou automaticamente

Masculino:
64 ordens existentes
novo ativo → recebeu 65 automaticamente
```

Também foi validado que atleta reativado posteriormente **não desloca atletas de grupo já confirmado**.

---

# 4. `IDSeguroManual` — IDs permanentes

## Rotina

```vb
Public Sub GerarIDsDosAtletas()
```

## Comportamento confirmado

Antes de gerar IDs:

1. exige arquivo `.xlsm`;
2. lê IDs já existentes no cadastro;
3. verifica se obedecem ao padrão `ATL-n`;
4. bloqueia ID duplicado dentro do cadastro;
5. lê IDs já utilizados no histórico;
6. descobre o maior número já utilizado;
7. conta atletas com nome e sem ID;
8. pede confirmação ao operador.

Na geração:

```text
novo ID = ATL-(maior número já utilizado + 1)
```

Ele não reutiliza IDs históricos.

Se a data de cadastro estiver vazia, também grava a data atual.

## Funções auxiliares

```vb
IDValido(textoId)
NumeroDoID(textoId)
```

Padrão aceito:

```text
ATL-1
ATL-2
ATL-3
...
```

## O que preservar no novo sistema

- ID permanente;
- unicidade;
- IDs antigos nunca reutilizados;
- histórico também participa da verificação;
- IDs existentes nunca são renumerados.

## O que melhorar

O VBA só lê:

```text
Cadastro: linhas 3–200
Histórico: linhas 9–10000
```

Esses limites não devem existir no novo sistema.

Ideal:

```text
athlete.id UUID ou chave interna
athlete.publicCode = ATL-123
UNIQUE(publicCode)
```

---

# 5. `CicloLiga` — fechamento/início de ciclo

## Rotina principal

```vb
Public Sub EncerrarCicloEIniciarProximo()
```

## Status

**LEGADO em relação à estrutura atual de grupos/categorias.**

Apesar disso, é o módulo com a melhor referência para entender as intenções de segurança do fechamento.

## Pré-condições verificadas

A macro exige:

### Confirmação

```text
FIM DE CICLO!B18 = SIM
```

### Status geral

```text
FIM DE CICLO!A14 =
PRONTO PARA GERAR A PRÓXIMA RODADA
```

### Ciclo válido

```text
HISTÓRICO DOS CICLOS!B5
deve ser numérico e >= 1
```

### IDs

- detecta ID duplicado no cadastro;
- exige prefixo `ATL-`;
- exige que ID da movimentação exista no cadastro;
- não permite o mesmo atleta duas vezes na movimentação.

### Movimentação

Cada atleta movimentado precisa ter:

```text
nova categoria
novo grupo
```

Também valida:

- grupo de destino configurado;
- no máximo 4 atletas por grupo;
- número de atletas ativos = número de atletas movimentados;
- todo atleta ativo aparece na movimentação.

### Resultados

A helper `ValidarResultadosConcluidos` bloqueia o fechamento se houver uma partida com quatro participantes e placar incompleto.

---

## Operação de fechamento

Em alto nível:

```text
validar
↓
fazer backup em memória
↓
arquivar ciclo, se ainda não arquivado
↓
atualizar cadastro com nova categoria/grupo
↓
reconstruir blocos da aba GRUPOS
↓
limpar placares
↓
limpar decisões manuais/observações
↓
reconfigurar ranking acumulado
↓
incrementar número do ciclo
↓
resetar confirmação de fechamento
↓
recalcular tudo
↓
salvar workbook
```

## Histórico

O módulo grava:

```text
Ciclo
Data
Categoria atual
Grupo atual
Posição
Atleta
Jogos
Vitórias
Derrotas
Games Pró
Games Contra
Saldo
Novo Grupo
ID ÚNICO
```

Ele evita duplicar um ciclo já existente.

---

## Estratégia de rollback

Antes de alterar dados, guarda cópias de:

```text
CADASTRO C3:D200
RESULTADOS F1:G80
SUBIDAS G6:G57
SUBIDAS J6:J57
RANKING AA2:AM53
RANKING A5:W36
Número do ciclo
Data
Confirmação de fechamento
Blocos da aba GRUPOS
```

Se ocorrer erro:

1. restaura esses backups;
2. remove linhas de histórico inseridas parcialmente;
3. reativa cálculo/eventos/tela;
4. força recálculo completo.

### Equivalente no sistema novo

Isto deve virar **transação de banco de dados**:

```text
BEGIN TRANSACTION

validar fechamento
persistir snapshot
criar memberships do próximo ciclo
atualizar estado do ciclo
limpar/encerrar estado transitório

COMMIT
```

Se qualquer etapa falhar:

```text
ROLLBACK
```

---

## Por que NÃO portar a macro literalmente

`CicloLiga` contém hard-coding antigo, incluindo:

```text
ULTIMA_LINHA_CADASTRO = 200
ULTIMA_LINHA_MOVIMENTACAO = 57
ULTIMA_LINHA_CLASSIFICACAO = 93
ULTIMA_LINHA_RESULTADOS = 80
```

e mapeia explicitamente categorias/grupos da arquitetura anterior.

Portar a intenção; **não portar o desenho físico do Excel**.

---

# 6. `HistoricoLiga` — arquivamento antigo

## Rotina

```vb
Public Sub ArquivarCicloNoHistorico()
```

## Comportamento

- exige `FIM DE CICLO = SIM`;
- exige ciclo válido;
- usa data informada ou data atual;
- bloqueia arquivamento de ciclo já registrado;
- percorre blocos específicos da CLASSIFICAÇÃO;
- copia estatísticas para `HISTÓRICO DOS CICLOS`;
- busca `Novo Grupo` na aba `SUBIDAS E DESCIDAS`;
- se falhar no meio, apaga registros parciais.

## Limitação crítica

O mapa é totalmente antigo e explícito:

```text
Avançado Masculino A/B
Intermediário Masculino A–G
Iniciante Masculino A/B
Intermediário Feminino A/B
```

Além disso, essa rotina usa o **nome do atleta** para localizar movimentação:

```vb
Application.Match(atleta, wsMov.Range("B6:B57"), 0)
```

Isso não deve ser reproduzido.

## No sistema novo

O histórico deve ser gerado dinamicamente pelo ciclo e usar exclusivamente:

```text
athleteId
cycleId
groupId
```

Nunca nome como chave de relação.

---

# 7. `ReiniciarLigaOficial`

## Rotina

```vb
Public Sub ReiniciarParaLigaOficial()
```

## Natureza

É uma **ferramenta administrativa destrutiva** criada durante a construção da planilha para limpar dados de teste e preparar uma base oficial.

Não representa fluxo normal de operação da Liga.

## Ações

Entre outras coisas:

- exige que a planilha esteja salva;
- exige formato com macro;
- conta conteúdo existente;
- mostra alerta;
- exige confirmação textual:
  `REINICIAR LIGA OFICIAL`;
- cria backup;
- limpa cadastro de teste;
- restaura fórmulas dos grupos;
- limpa placares;
- limpa campos manuais de movimentação;
- apaga histórico;
- redefine ciclo para 1;
- redefine confirmação do fim de ciclo.

## Cuidado

A macro contém configuração antiga com:

```text
13 grupos x 4 = 52 atletas
```

Portanto, é legada.

---

# 8. `ReiniciarLigaOficialV2`

## Rotina

```vb
Public Sub ReiniciarParaLigaOficialV2()
```

## Diferença relevante para V1

A V2 foi criada principalmente para tornar o backup pré-reset mais seguro em ambientes OneDrive.

Em vez de confiar diretamente em `ThisWorkbook.Path`, ela:

1. sugere Downloads;
2. abre `GetSaveAsFilename`;
3. deixa o usuário escolher onde salvar;
4. usa `SaveCopyAs`;
5. verifica se o arquivo de backup realmente existe;
6. só depois executa a primeira alteração destrutiva.

Comentário presente no próprio código:

```text
Quando a planilha está aberta pelo OneDrive,
ThisWorkbook.Path pode retornar um endereço
https://d.docs.live.net.
SaveCopyAs não consegue gravar diretamente nesse endereço.
```

## Para o novo sistema

Se houver uma função de reset administrativo:

- não deverá existir no fluxo comum;
- exigir papel de administrador;
- confirmação forte;
- backup/snapshot prévio;
- auditoria;
- idealmente soft-delete ou ambiente de teste em vez de limpeza destrutiva.

---

# 9. O que o Claude deve considerar fonte de verdade

## Manter como regra atual

### Ordem automática

`Planilha12.Worksheet_Change`

A regra de negócio por trás dela é atual e foi testada.

### IDs

Conceitos de `IDSeguroManual`:

- ID permanente;
- único;
- não reutilizar histórico;
- não renumerar registros existentes.

### Segurança de fechamento

Conceitos de `CicloLiga`:

- pré-validação;
- nenhuma pendência;
- confirmação explícita;
- snapshot histórico;
- atomicidade;
- rollback;
- grupos com no máximo 4.

---

## Usar apenas como referência histórica

```text
CicloLiga
HistoricoLiga
ReiniciarLigaOficial
ReiniciarLigaOficialV2
```

quando o assunto for:

```text
quais categorias existem
quantos grupos existem
onde cada grupo está
qual linha representa cada grupo
quantos atletas o sistema suporta
```

Esses valores estão hard-coded e ficaram defasados durante a evolução da planilha.

---

# 10. Triggers encontrados no VBA

## Automático

Apenas um trigger de edição relevante foi encontrado:

```vb
Worksheet_Change
```

na aba `CADASTRO DE ATLETAS`.

Ele dispara automaticamente em edição de:

```text
Nome
Categoria
Status
```

e pode gerar `ORDEM DE ENTRADA`.

---

## Manuais

As seguintes rotinas são `Public Sub` e precisam ser chamadas manualmente por macro/botão/Alt+F8 ou outra chamada VBA:

```text
EncerrarCicloEIniciarProximo
ArquivarCicloNoHistorico
GerarIDsDosAtletas
ReiniciarParaLigaOficial
ReiniciarParaLigaOficialV2
```

A extração do código não permite afirmar com segurança qual objeto gráfico da planilha está associado a cada uma. Não inventar essa associação.

---

# 11. Riscos técnicos identificados no VBA

## 11.1 Hard-coded ranges

Exemplos:

```text
A3:G200
G6:G57
B6:B93
F1:G80
AA2:AM53
A9:N10000
```

No sistema novo tudo deve ser orientado a registros, não linhas.

## 11.2 Categorias legadas

Há coexistência de:

```text
Masculino / Feminino
```

com:

```text
Avançado Masculino
Intermediário Masculino
Iniciante Masculino
Intermediário Feminino
```

Não modelar categoria a partir dessas strings antigas.

## 11.3 Match por nome

Algumas rotinas antigas fazem lookup por nome.

No sistema novo: somente ID.

## 11.4 Macros destrutivas

As duas rotinas `ReiniciarLigaOficial*` limpam dados.

Não portar para produção sem:

- RBAC/permissão;
- confirmação;
- snapshot;
- auditoria;
- rollback.

## 11.5 Estado distribuído entre células

No Excel o ciclo depende de células como:

```text
HISTÓRICO DOS CICLOS!B5
HISTÓRICO DOS CICLOS!B6
FIM DE CICLO!B18
```

No novo sistema devem ser campos explícitos da entidade `Cycle`.

---

# 12. Equivalência recomendada no backend

| VBA/Excel | Serviço no sistema novo |
|---|---|
| `Worksheet_Change` | `activateAthlete()` / `assignEntryOrder()` |
| `GerarIDsDosAtletas` | criação de `Athlete` + geração de public code |
| `ArquivarCicloNoHistorico` | `createCycleSnapshot()` |
| `EncerrarCicloEIniciarProximo` | `closeCycleAndCreateNext()` transacional |
| `ReiniciarLigaOficial*` | ferramenta administrativa separada, não fluxo comum |
| fórmulas de grupo | `GroupAssignmentService` |
| fórmulas de ranking | `RankingService` |
| subidas/descidas | `MovementService` |
| coringas | `SubstitutionService` |

---

# 13. Testes de regressão derivados do VBA

O novo sistema deve testar no mínimo:

### IDs

- não criar ID duplicado;
- não reutilizar ID existente no histórico;
- não alterar ID permanente;
- atleta sem ID recebe próximo código.

### Ordem de entrada

- atleta ativo novo recebe próximo número da categoria;
- masculino e feminino têm sequências independentes;
- inativação não libera/reutiliza ordem;
- reativação com ordem existente não gera nova ordem;
- grupo já confirmado não muda quando atleta antigo é reativado.

### Fechamento

- bloqueia com partida pendente;
- bloqueia sem confirmação;
- bloqueia atleta ativo sem movimentação;
- bloqueia destino inexistente;
- bloqueia grupo com mais de 4;
- não duplica histórico do mesmo ciclo;
- fechamento completo incrementa ciclo;
- erro no meio não pode produzir estado parcial.

---

# 14. Código-fonte completo extraído

Abaixo estão os módulos VBA com lógica relevante presentes no arquivo analisado.

## 14.1 `Planilha12`

**Nota:** Evento da aba `CADASTRO DE ATLETAS`; lógica atual da ordem automática.

SHA-256 do código extraído:

`63540a81a734a141eff2a62c06e38adca130708f4f7c4464745e68f7b76d939c`

```vb
Attribute VB_Name = "Planilha12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
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

## 14.2 `IDSeguroManual`

**Nota:** Gerador de IDs permanentes; conceito útil, porém com limite físico até linha 200.

SHA-256 do código extraído:

`8171c01c06404d9abadd1ba841397aa1288de513a98623ef2fd595b5c8660acb`

```vb
Attribute VB_Name = "IDSeguroManual"
Option Explicit

Public Sub GerarIDsDosAtletas()

    Const PRIMEIRA_LINHA_CADASTRO As Long = 3
    Const ULTIMA_LINHA_CADASTRO As Long = 200
    Const PRIMEIRA_LINHA_HISTORICO As Long = 9
    Const ULTIMA_LINHA_HISTORICO As Long = 10000

    Dim wsCadastro As Worksheet
    Dim wsHistorico As Worksheet
    Dim idsUsados As Object

    Dim linha As Long
    Dim idAtual As String
    Dim nomeAtleta As String
    Dim novoId As String
    Dim proximoNumero As Long
    Dim quantidadeGerada As Long
    Dim resposta As VbMsgBoxResult

    Dim calculoAnterior As XlCalculation
    Dim eventosAnteriores As Boolean
    Dim telaAnterior As Boolean
    Dim mensagemErro As String

    On Error GoTo TratarErro

    Set wsCadastro = ThisWorkbook.Worksheets("CADASTRO DE ATLETAS")
    Set wsHistorico = ThisWorkbook.Worksheets("HISTÓRICO DOS CICLOS")

    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        MsgBox "Esta função deve ser usada na planilha oficial .xlsm.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    Set idsUsados = CreateObject("Scripting.Dictionary")
    idsUsados.CompareMode = 1

    For linha = PRIMEIRA_LINHA_CADASTRO To ULTIMA_LINHA_CADASTRO

        idAtual = Trim$(CStr(wsCadastro.Cells(linha, 1).Value))

        If idAtual <> "" Then

            If Not IDValido(idAtual) Then
                MsgBox "Foi encontrado um ID fora do padrão na linha " & linha & ":" & vbCrLf & _
                       idAtual & vbCrLf & vbCrLf & _
                       "O padrão correto é ATL-1, ATL-2, ATL-3...", _
                       vbExclamation, "Liga Padel"
                Exit Sub
            End If

            If idsUsados.Exists(UCase$(idAtual)) Then
                MsgBox "Existe ID repetido no cadastro:" & vbCrLf & _
                       idAtual & vbCrLf & vbCrLf & _
                       "Corrija antes de gerar novos IDs.", _
                       vbCritical, "Liga Padel"
                Exit Sub
            End If

            idsUsados.Add UCase$(idAtual), True

            If NumeroDoID(idAtual) > proximoNumero Then
                proximoNumero = NumeroDoID(idAtual)
            End If

        End If

    Next linha

    For linha = PRIMEIRA_LINHA_HISTORICO To ULTIMA_LINHA_HISTORICO

        idAtual = Trim$(CStr(wsHistorico.Cells(linha, 14).Value))

        If idAtual <> "" Then

            If IDValido(idAtual) Then

                If Not idsUsados.Exists(UCase$(idAtual)) Then
                    idsUsados.Add UCase$(idAtual), True
                End If

                If NumeroDoID(idAtual) > proximoNumero Then
                    proximoNumero = NumeroDoID(idAtual)
                End If

            End If

        End If

    Next linha

    For linha = PRIMEIRA_LINHA_CADASTRO To ULTIMA_LINHA_CADASTRO

        nomeAtleta = Trim$(CStr(wsCadastro.Cells(linha, 2).Value))
        idAtual = Trim$(CStr(wsCadastro.Cells(linha, 1).Value))

        If nomeAtleta <> "" And idAtual = "" Then
            quantidadeGerada = quantidadeGerada + 1
        End If

    Next linha

    If quantidadeGerada = 0 Then
        MsgBox "Não existem atletas aguardando geração de ID." & vbCrLf & _
               "Os IDs já preenchidos não foram alterados.", _
               vbInformation, "Liga Padel"
        Exit Sub
    End If

    resposta = MsgBox( _
        quantidadeGerada & " atleta(s) estão com nome preenchido e sem ID." & vbCrLf & vbCrLf & _
        "A macro criará IDs permanentes no padrão ATL-1, ATL-2, ATL-3..." & vbCrLf & _
        "IDs existentes não serão alterados." & vbCrLf & vbCrLf & _
        "Deseja continuar?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Gerar IDs dos Atletas")

    If resposta <> vbYes Then
        MsgBox "Operação cancelada. Nenhum dado foi alterado.", _
               vbInformation, "Liga Padel"
        Exit Sub
    End If

    calculoAnterior = Application.Calculation
    eventosAnteriores = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    quantidadeGerada = 0

    For linha = PRIMEIRA_LINHA_CADASTRO To ULTIMA_LINHA_CADASTRO

        nomeAtleta = Trim$(CStr(wsCadastro.Cells(linha, 2).Value))
        idAtual = Trim$(CStr(wsCadastro.Cells(linha, 1).Value))

        If nomeAtleta <> "" And idAtual = "" Then

            Do
                proximoNumero = proximoNumero + 1
                novoId = "ATL-" & CStr(proximoNumero)
            Loop While idsUsados.Exists(UCase$(novoId))

            wsCadastro.Cells(linha, 1).Value = novoId
            idsUsados.Add UCase$(novoId), True

            If Trim$(CStr(wsCadastro.Cells(linha, 6).Value)) = "" Then
                wsCadastro.Cells(linha, 6).Value = Date
                wsCadastro.Cells(linha, 6).NumberFormat = "dd/mm/yyyy"
            End If

            quantidadeGerada = quantidadeGerada + 1

        End If

    Next linha

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    ThisWorkbook.Save

    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    MsgBox quantidadeGerada & " ID(s) permanente(s) foram gerados com sucesso." & vbCrLf & vbCrLf & _
           "Os IDs existentes não foram modificados." & vbCrLf & _
           "A planilha foi salva.", _
           vbInformation, "Liga Padel"

    Exit Sub

TratarErro:

    mensagemErro = Err.Description

    On Error Resume Next
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    MsgBox "A geração de IDs foi interrompida." & vbCrLf & _
           "Confira a planilha antes de salvar novamente." & vbCrLf & vbCrLf & _
           "Detalhe: " & mensagemErro, _
           vbCritical, "Liga Padel"

End Sub

Private Function IDValido(ByVal textoId As String) As Boolean

    textoId = UCase$(Trim$(textoId))

    If Left$(textoId, 4) <> "ATL-" Then Exit Function
    If Len(textoId) <= 4 Then Exit Function
    If Not IsNumeric(Mid$(textoId, 5)) Then Exit Function
    If CLng(Mid$(textoId, 5)) < 1 Then Exit Function

    IDValido = True

End Function

Private Function NumeroDoID(ByVal textoId As String) As Long

    If IDValido(textoId) Then
        NumeroDoID = CLng(Mid$(Trim$(textoId), 5))
    Else
        NumeroDoID = 0
    End If

End Function
```

## 14.3 `CicloLiga`

**Nota:** Fechamento de ciclo antigo; contém ótimas validações, mas mapeamento de categorias/grupos está legado.

SHA-256 do código extraído:

`a7e09c17ec061d4208c760ff14a509fb740b46a33cc45de7f64c16ee024b45d3`

```vb
Attribute VB_Name = "CicloLiga"
Option Explicit

Private Const PRIMEIRA_LINHA_HISTORICO As Long = 9
Private Const PRIMEIRA_LINHA_CADASTRO As Long = 3
Private Const ULTIMA_LINHA_CADASTRO As Long = 200
Private Const PRIMEIRA_LINHA_MOVIMENTACAO As Long = 6
Private Const ULTIMA_LINHA_MOVIMENTACAO As Long = 57
Private Const ULTIMA_LINHA_CLASSIFICACAO As Long = 93
Private Const ULTIMA_LINHA_RESULTADOS As Long = 80

Public Sub EncerrarCicloEIniciarProximo()

    Dim wsCad As Worksheet
    Dim wsGrupos As Worksheet
    Dim wsResultados As Worksheet
    Dim wsClass As Worksheet
    Dim wsRank As Worksheet
    Dim wsMov As Worksheet
    Dim wsFim As Worksheet
    Dim wsProxima As Worksheet
    Dim wsHist As Worksheet

    Dim mapaBlocos As Object
    Dim mapaClassificacao As Object
    Dim linhasCadastro As Object
    Dim idsMovimentados As Object
    Dim quantidadePorGrupo As Object

    Dim cicloAtual As Long
    Dim proximoCiclo As Long
    Dim maiorCicloHistorico As Long
    Dim dataEncerramento As Date
    Dim cicloJaArquivado As Boolean

    Dim linha As Long
    Dim linhaCadastro As Long
    Dim linhaClassificacao As Long
    Dim linhaDestino As Long
    Dim primeiraLinhaInserida As Long
    Dim ultimaLinhaInserida As Long
    Dim quantidadeAtletas As Long
    Dim quantidadeAtivos As Long

    Dim idAtleta As String
    Dim atleta As String
    Dim categoriaAtual As String
    Dim grupoAtual As String
    Dim novaCategoria As String
    Dim novoGrupo As String
    Dim chave As String
    Dim posicao As Long

    Dim dadosBloco As Variant
    Dim chavesBlocos As Variant
    Dim backupBlocos() As Variant
    Dim i As Long
    Dim vaga As Long
    Dim linhaGrupo As Long

    Dim backupCadastro As Variant
    Dim backupResultados As Variant
    Dim backupDecisoes As Variant
    Dim backupObservacoes As Variant
    Dim backupRankingAuxiliar As Variant
    Dim backupRankingVisivel As Variant
    Dim backupCiclo As Variant
    Dim backupData As Variant
    Dim backupConfirmacao As Variant
    Dim backupsCriados As Boolean
    Dim mensagemErro As String

    On Error GoTo TratarErro

    Set wsCad = ThisWorkbook.Worksheets("CADASTRO DE ATLETAS")
    Set wsGrupos = ThisWorkbook.Worksheets("GRUPOS")
    Set wsResultados = ThisWorkbook.Worksheets("LANÇAR RESULTADOS")
    Set wsClass = ThisWorkbook.Worksheets("CLASSIFICAÇÃO")
    Set wsRank = ThisWorkbook.Worksheets("RANKING PESSOAL")
    Set wsMov = ThisWorkbook.Worksheets("SUBIDAS E DESCIDAS")
    Set wsFim = ThisWorkbook.Worksheets("FIM DE CICLO")
    Set wsProxima = ThisWorkbook.Worksheets("PRÓXIMA RODADA")
    Set wsHist = ThisWorkbook.Worksheets("HISTÓRICO DOS CICLOS")

    If UCase$(Trim$(CStr(wsFim.Range("B18").Value))) <> "SIM" Then
        MsgBox "Digite SIM em FIM DE CICLO!B18 antes de encerrar o ciclo.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    If CStr(wsFim.Range("A14").Value) <> "PRONTO PARA GERAR A PRÓXIMA RODADA" Then
        MsgBox "O ciclo ainda possui pendências." & vbCrLf & _
               "Revise a aba FIM DE CICLO antes de continuar.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    If Not IsNumeric(wsHist.Range("B5").Value) Then
        MsgBox "Informe um número de ciclo válido em HISTÓRICO DOS CICLOS!B5.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    cicloAtual = CLng(wsHist.Range("B5").Value)

    If cicloAtual < 1 Then
        MsgBox "O número do ciclo deve ser maior ou igual a 1.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    If IsDate(wsHist.Range("B6").Value) Then
        dataEncerramento = CDate(wsHist.Range("B6").Value)
    Else
        dataEncerramento = Date
    End If

    Set mapaBlocos = CriarMapaDosBlocos()
    Set mapaClassificacao = CriarMapaDaClassificacao()
    Set linhasCadastro = CreateObject("Scripting.Dictionary")
    Set idsMovimentados = CreateObject("Scripting.Dictionary")
    Set quantidadePorGrupo = CreateObject("Scripting.Dictionary")

    linhasCadastro.CompareMode = 1
    idsMovimentados.CompareMode = 1
    quantidadePorGrupo.CompareMode = 1

    For linha = PRIMEIRA_LINHA_CADASTRO To ULTIMA_LINHA_CADASTRO
        idAtleta = Trim$(CStr(wsCad.Cells(linha, 1).Value))

        If idAtleta <> "" Then
            If linhasCadastro.Exists(idAtleta) Then
                Err.Raise vbObjectError + 1001, , _
                          "Existe ID duplicado no cadastro: " & idAtleta
            End If

            linhasCadastro.Add idAtleta, linha
        End If
    Next linha

    quantidadeAtivos = Application.WorksheetFunction.CountIf( _
                       wsCad.Range("E" & PRIMEIRA_LINHA_CADASTRO & _
                       ":E" & ULTIMA_LINHA_CADASTRO), "ATIVO")

    For linha = PRIMEIRA_LINHA_MOVIMENTACAO To ULTIMA_LINHA_MOVIMENTACAO

        idAtleta = Trim$(CStr(wsMov.Cells(linha, 1).Value))

        If idAtleta <> "" Then

            If Left$(UCase$(idAtleta), 4) <> "ATL-" Then
                Err.Raise vbObjectError + 1002, , _
                          "ID inválido na movimentação, linha " & linha & ": " & idAtleta
            End If

            If Not linhasCadastro.Exists(idAtleta) Then
                Err.Raise vbObjectError + 1003, , _
                          "O ID " & idAtleta & " não foi encontrado no cadastro."
            End If

            If idsMovimentados.Exists(idAtleta) Then
                Err.Raise vbObjectError + 1004, , _
                          "O atleta " & idAtleta & " aparece mais de uma vez nas movimentações."
            End If

            novaCategoria = Trim$(CStr(wsMov.Cells(linha, 8).Value))
            novoGrupo = Trim$(CStr(wsMov.Cells(linha, 9).Value))

            If novaCategoria = "" Or novoGrupo = "" Then
                Err.Raise vbObjectError + 1005, , _
                          "Há atleta sem nova categoria ou novo grupo na linha " & linha & "."
            End If

            chave = MontarChave(novaCategoria, novoGrupo)

            If Not mapaBlocos.Exists(chave) Then
                Err.Raise vbObjectError + 1006, , _
                          "Não existe espaço configurado para: " & _
                          novaCategoria & " — " & novoGrupo
            End If

            If Not quantidadePorGrupo.Exists(chave) Then
                quantidadePorGrupo.Add chave, 0
            End If

            quantidadePorGrupo(chave) = CLng(quantidadePorGrupo(chave)) + 1

            If CLng(quantidadePorGrupo(chave)) > 4 Then
                Err.Raise vbObjectError + 1007, , _
                          "O grupo " & novaCategoria & " — " & novoGrupo & _
                          " ficou com mais de 4 atletas."
            End If

            idsMovimentados.Add idAtleta, linha

        End If

    Next linha

    If idsMovimentados.Count <> quantidadeAtivos Then
        Err.Raise vbObjectError + 1008, , _
                  "A quantidade de atletas ativos (" & quantidadeAtivos & _
                  ") é diferente da quantidade de atletas movimentados (" & _
                  idsMovimentados.Count & ")."
    End If

    For linha = PRIMEIRA_LINHA_CADASTRO To ULTIMA_LINHA_CADASTRO
        If UCase$(Trim$(CStr(wsCad.Cells(linha, 5).Value))) = "ATIVO" Then

            idAtleta = Trim$(CStr(wsCad.Cells(linha, 1).Value))

            If idAtleta = "" Or Not idsMovimentados.Exists(idAtleta) Then
                Err.Raise vbObjectError + 1009, , _
                          "O atleta ativo da linha " & linha & _
                          " não está representado nas movimentações."
            End If

        End If
    Next linha

    ValidarResultadosConcluidos wsResultados

    maiorCicloHistorico = MaiorCicloRegistrado(wsHist)

    If maiorCicloHistorico > cicloAtual Then
        Err.Raise vbObjectError + 1010, , _
                  "O histórico já contém ciclo superior ao número informado em B5."
    End If

    cicloJaArquivado = (Application.WorksheetFunction.CountIf( _
                        wsHist.Range("A" & PRIMEIRA_LINHA_HISTORICO & _
                        ":A" & wsHist.Rows.Count), cicloAtual) > 0)

    backupCadastro = wsCad.Range("C3:D200").Formula
    backupResultados = wsResultados.Range("F1:G80").Value2
    backupDecisoes = wsMov.Range("G6:G57").Formula
    backupObservacoes = wsMov.Range("J6:J57").Formula
    backupRankingAuxiliar = wsRank.Range("AA2:AM53").Formula
    backupRankingVisivel = wsRank.Range("A5:W36").Formula
    backupCiclo = wsHist.Range("B5").Value
    backupData = wsHist.Range("B6").Value
    backupConfirmacao = wsFim.Range("B18").Value

    chavesBlocos = mapaBlocos.Keys
    ReDim backupBlocos(LBound(chavesBlocos) To UBound(chavesBlocos))

    For i = LBound(chavesBlocos) To UBound(chavesBlocos)
        dadosBloco = mapaBlocos(chavesBlocos(i))
        backupBlocos(i) = wsGrupos.Range( _
                           wsGrupos.Cells(CLng(dadosBloco(4)), CLng(dadosBloco(0))), _
                           wsGrupos.Cells(CLng(dadosBloco(4)) + 3, CLng(dadosBloco(3))) _
                           ).Formula
    Next i

    backupsCriados = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    PrepararColunaIdHistorico wsHist, wsCad

    If Not cicloJaArquivado Then

        linhaDestino = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
        If linhaDestino < PRIMEIRA_LINHA_HISTORICO Then
            linhaDestino = PRIMEIRA_LINHA_HISTORICO
        End If

        primeiraLinhaInserida = linhaDestino

        For linha = PRIMEIRA_LINHA_MOVIMENTACAO To ULTIMA_LINHA_MOVIMENTACAO

            idAtleta = Trim$(CStr(wsMov.Cells(linha, 1).Value))

            If idAtleta <> "" Then

                atleta = Trim$(CStr(wsMov.Cells(linha, 2).Value))
                categoriaAtual = Trim$(CStr(wsMov.Cells(linha, 3).Value))
                grupoAtual = Trim$(CStr(wsMov.Cells(linha, 4).Value))
                posicao = CLng(wsMov.Cells(linha, 5).Value)
                novoGrupo = Trim$(CStr(wsMov.Cells(linha, 9).Value))

                chave = MontarChave(categoriaAtual, grupoAtual)

                If Not mapaClassificacao.Exists(chave) Then
                    Err.Raise vbObjectError + 1011, , _
                              "Não foi possível localizar a classificação de " & _
                              categoriaAtual & " — " & grupoAtual
                End If

                linhaClassificacao = CLng(mapaClassificacao(chave)) + posicao - 1

                If UCase$(Trim$(CStr(wsClass.Cells(linhaClassificacao, 2).Value))) <> _
                   UCase$(atleta) Then

                    Err.Raise vbObjectError + 1012, , _
                              "A classificação não corresponde à movimentação do atleta: " & atleta
                End If

                wsHist.Cells(linhaDestino, 1).Value = cicloAtual
                wsHist.Cells(linhaDestino, 2).Value = dataEncerramento
                wsHist.Cells(linhaDestino, 3).Value = categoriaAtual
                wsHist.Cells(linhaDestino, 4).Value = grupoAtual
                wsHist.Cells(linhaDestino, 5).Value = posicao
                wsHist.Cells(linhaDestino, 6).Value = atleta
                wsHist.Cells(linhaDestino, 7).Value = wsClass.Cells(linhaClassificacao, 3).Value
                wsHist.Cells(linhaDestino, 8).Value = wsClass.Cells(linhaClassificacao, 4).Value
                wsHist.Cells(linhaDestino, 9).Value = wsClass.Cells(linhaClassificacao, 5).Value
                wsHist.Cells(linhaDestino, 10).Value = wsClass.Cells(linhaClassificacao, 6).Value
                wsHist.Cells(linhaDestino, 11).Value = wsClass.Cells(linhaClassificacao, 7).Value
                wsHist.Cells(linhaDestino, 12).Value = wsClass.Cells(linhaClassificacao, 8).Value
                wsHist.Cells(linhaDestino, 13).Value = novoGrupo
                wsHist.Cells(linhaDestino, 14).Value = idAtleta

                linhaDestino = linhaDestino + 1
                quantidadeAtletas = quantidadeAtletas + 1

            End If
        Next linha

        ultimaLinhaInserida = linhaDestino - 1

        If quantidadeAtletas = 0 Then
            Err.Raise vbObjectError + 1013, , _
                      "Nenhum atleta foi encontrado para arquivamento."
        End If

        wsHist.Range("B" & primeiraLinhaInserida & ":B" & ultimaLinhaInserida).NumberFormat = "dd/mm/yyyy"
        wsHist.Range("E" & primeiraLinhaInserida & ":E" & ultimaLinhaInserida).NumberFormat = "0°"

    Else
        quantidadeAtletas = idsMovimentados.Count
    End If

    ' Atualiza categoria e grupo atuais no cadastro.
    For linha = PRIMEIRA_LINHA_MOVIMENTACAO To ULTIMA_LINHA_MOVIMENTACAO

        idAtleta = Trim$(CStr(wsMov.Cells(linha, 1).Value))

        If idAtleta <> "" Then
            linhaCadastro = CLng(linhasCadastro(idAtleta))
            wsCad.Cells(linhaCadastro, 3).Value = Trim$(CStr(wsMov.Cells(linha, 8).Value))
            wsCad.Cells(linhaCadastro, 4).Value = Trim$(CStr(wsMov.Cells(linha, 9).Value))
        End If

    Next linha

    ' Limpa as vagas atuais de todos os grupos.
    For i = LBound(chavesBlocos) To UBound(chavesBlocos)

        dadosBloco = mapaBlocos(chavesBlocos(i))

        wsGrupos.Range( _
            wsGrupos.Cells(CLng(dadosBloco(4)), CLng(dadosBloco(0))), _
            wsGrupos.Cells(CLng(dadosBloco(4)) + 3, CLng(dadosBloco(3))) _
        ).ClearContents

    Next i

    quantidadePorGrupo.RemoveAll

    ' Preenche os novos grupos, respeitando categoria + grupo.
    For linha = PRIMEIRA_LINHA_MOVIMENTACAO To ULTIMA_LINHA_MOVIMENTACAO

        idAtleta = Trim$(CStr(wsMov.Cells(linha, 1).Value))

        If idAtleta <> "" Then

            novaCategoria = Trim$(CStr(wsMov.Cells(linha, 8).Value))
            novoGrupo = Trim$(CStr(wsMov.Cells(linha, 9).Value))
            chave = MontarChave(novaCategoria, novoGrupo)

            If Not quantidadePorGrupo.Exists(chave) Then
                quantidadePorGrupo.Add chave, 0
            End If

            quantidadePorGrupo(chave) = CLng(quantidadePorGrupo(chave)) + 1
            vaga = CLng(quantidadePorGrupo(chave))

            dadosBloco = mapaBlocos(chave)
            linhaGrupo = CLng(dadosBloco(4)) + vaga - 1

            wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(0))).Value = idAtleta

            AplicarFormulaDeCadastro wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(1))), _
                                     wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(0))), 2

            AplicarFormulaDeCadastro wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(2))), _
                                     wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(0))), 5

            AplicarFormulaDeCadastro wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(3))), _
                                     wsGrupos.Cells(linhaGrupo, CLng(dadosBloco(0))), 7

        End If

    Next linha

    ' Limpa somente os placares das rodadas que serão reutilizadas no novo ciclo.
    For linha = 1 To ULTIMA_LINHA_RESULTADOS

        If IsNumeric(wsResultados.Cells(linha, 3).Value) Then
            If CLng(wsResultados.Cells(linha, 3).Value) >= 1 And _
               CLng(wsResultados.Cells(linha, 3).Value) <= 3 Then

                wsResultados.Cells(linha, 6).ClearContents
                wsResultados.Cells(linha, 7).ClearContents

            End If
        End If

    Next linha

    ' As decisões manuais e observações administrativas não passam para o novo ciclo.
    wsMov.Range("G6:G57").ClearContents
    wsMov.Range("J6:J57").ClearContents

    ' O ranking passa a somar o histórico encerrado com os resultados do ciclo em andamento.
    ConfigurarRankingAcumulado wsRank

    maiorCicloHistorico = MaiorCicloRegistrado(wsHist)
    proximoCiclo = Application.WorksheetFunction.Max(cicloAtual, maiorCicloHistorico) + 1

    wsHist.Range("B5").Value = proximoCiclo
    wsHist.Range("B6").ClearContents
    wsFim.Range("B18").Value = "NÃO"

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    ThisWorkbook.Save

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If cicloJaArquivado Then
        MsgBox "O ciclo " & cicloAtual & " já estava registrado no histórico." & vbCrLf & vbCrLf & _
               "Ciclo " & proximoCiclo & " iniciado com sucesso." & vbCrLf & _
               quantidadeAtletas & " atletas foram distribuídos nos novos grupos." & vbCrLf & _
               "Os placares foram limpos e o ranking acumulado foi preservado.", _
               vbInformation, "Liga Padel"
    Else
        MsgBox "Ciclo " & cicloAtual & " encerrado e registrado no histórico." & vbCrLf & vbCrLf & _
               "Ciclo " & proximoCiclo & " iniciado com sucesso." & vbCrLf & _
               quantidadeAtletas & " atletas foram distribuídos nos novos grupos." & vbCrLf & _
               "Os placares foram limpos e o ranking acumulado foi preservado.", _
               vbInformation, "Liga Padel"
    End If

    Exit Sub

TratarErro:

    mensagemErro = Err.Description
    On Error Resume Next

    If backupsCriados Then

        wsCad.Range("C3:D200").Formula = backupCadastro
        wsResultados.Range("F1:G80").Value2 = backupResultados
        wsMov.Range("G6:G57").Formula = backupDecisoes
        wsMov.Range("J6:J57").Formula = backupObservacoes
        wsRank.Range("AA2:AM53").Formula = backupRankingAuxiliar
        wsRank.Range("A5:W36").Formula = backupRankingVisivel
        wsHist.Range("B5").Value = backupCiclo
        wsHist.Range("B6").Value = backupData
        wsFim.Range("B18").Value = backupConfirmacao

        If IsArray(chavesBlocos) Then
            For i = LBound(chavesBlocos) To UBound(chavesBlocos)
                dadosBloco = mapaBlocos(chavesBlocos(i))

                wsGrupos.Range( _
                    wsGrupos.Cells(CLng(dadosBloco(4)), CLng(dadosBloco(0))), _
                    wsGrupos.Cells(CLng(dadosBloco(4)) + 3, CLng(dadosBloco(3))) _
                ).Formula = backupBlocos(i)
            Next i
        End If

    End If

    If primeiraLinhaInserida >= PRIMEIRA_LINHA_HISTORICO And _
       ultimaLinhaInserida >= primeiraLinhaInserida Then

        wsHist.Range("A" & primeiraLinhaInserida & _
                     ":N" & ultimaLinhaInserida).ClearContents
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CalculateFullRebuild

    MsgBox "O encerramento foi cancelado." & vbCrLf & _
           "A planilha tentou restaurar os dados anteriores." & vbCrLf & vbCrLf & _
           "Detalhe: " & mensagemErro, _
           vbCritical, "Liga Padel"

End Sub

Private Sub ValidarResultadosConcluidos(ByVal wsResultados As Worksheet)

    Dim linha As Long
    Dim possuiJogo As Boolean

    For linha = 1 To ULTIMA_LINHA_RESULTADOS

        possuiJogo = _
            Trim$(CStr(wsResultados.Cells(linha, 4).Value)) <> "" And _
            Trim$(CStr(wsResultados.Cells(linha, 5).Value)) <> "" And _
            Trim$(CStr(wsResultados.Cells(linha, 8).Value)) <> "" And _
            Trim$(CStr(wsResultados.Cells(linha, 9).Value)) <> ""

        If possuiJogo Then

            If Trim$(CStr(wsResultados.Cells(linha, 6).Value)) = "" Or _
               Trim$(CStr(wsResultados.Cells(linha, 7).Value)) = "" Then

                Err.Raise vbObjectError + 1020, , _
                          "Existe jogo sem placar na linha " & linha & _
                          " da aba LANÇAR RESULTADOS."
            End If

        End If

    Next linha

End Sub

Private Sub PrepararColunaIdHistorico(ByVal wsHist As Worksheet, _
                                      ByVal wsCad As Worksheet)

    Dim ultimaLinha As Long
    Dim linha As Long
    Dim atleta As String
    Dim resultado As Variant

    If Trim$(CStr(wsHist.Range("N8").Value)) = "" Then
        wsHist.Range("M8").Copy
        wsHist.Range("N8").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsHist.Range("N8").Value = "ID ÚNICO"
        wsHist.Columns("N").ColumnWidth = 14
    End If

    ultimaLinha = wsHist.Cells(wsHist.Rows.Count, "F").End(xlUp).Row

    If ultimaLinha < PRIMEIRA_LINHA_HISTORICO Then Exit Sub

    For linha = PRIMEIRA_LINHA_HISTORICO To ultimaLinha

        If Trim$(CStr(wsHist.Cells(linha, 14).Value)) = "" Then

            atleta = Trim$(CStr(wsHist.Cells(linha, 6).Value))

            If atleta <> "" Then

                resultado = Application.Match( _
                            atleta, _
                            wsCad.Range("B" & PRIMEIRA_LINHA_CADASTRO & _
                            ":B" & ULTIMA_LINHA_CADASTRO), _
                            0)

                If Not IsError(resultado) Then
                    wsHist.Cells(linha, 14).Value = _
                        wsCad.Cells(PRIMEIRA_LINHA_CADASTRO - 1 + CLng(resultado), 1).Value
                End If

            End If

        End If

    Next linha

End Sub

Private Sub AplicarFormulaDeCadastro(ByVal celulaDestino As Range, _
                                     ByVal celulaId As Range, _
                                     ByVal colunaCadastro As Long)

    Dim enderecoId As String

    enderecoId = celulaId.Address(False, False)

    celulaDestino.Formula = _
        "=IF(" & enderecoId & "="""","""",IFERROR(INDEX('CADASTRO DE ATLETAS'!" & _
        ColunaLetra(colunaCadastro) & "$3:" & ColunaLetra(colunaCadastro) & "$200," & _
        "MATCH(" & enderecoId & ",'CADASTRO DE ATLETAS'!$A$3:$A$200,0)),""""))"

End Sub

Private Sub ConfigurarRankingAcumulado(ByVal wsRank As Worksheet)

    Dim linha As Long
    Dim linhaCadastro As Long
    Dim posicao As Long

    For linha = 2 To 53

        linhaCadastro = linha + 1

        wsRank.Cells(linha, 27).Formula = _
            "=IF('CADASTRO DE ATLETAS'!A" & linhaCadastro & "="""","""",'CADASTRO DE ATLETAS'!A" & linhaCadastro & ")"

        wsRank.Cells(linha, 28).Formula = _
            "=IF($AA" & linha & "="""","""",'CADASTRO DE ATLETAS'!B" & linhaCadastro & ")"

        wsRank.Cells(linha, 29).Formula = _
            "=IF($AA" & linha & "="""","""",'CADASTRO DE ATLETAS'!C" & linhaCadastro & ")"

        wsRank.Cells(linha, 30).Formula = _
            "=IF($AA" & linha & "="""","""",'CADASTRO DE ATLETAS'!D" & linhaCadastro & ")"

        wsRank.Cells(linha, 31).Formula = _
            "=IF($AA" & linha & "="""","""",SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA" & linha & _
            ",'HISTÓRICO DOS CICLOS'!$G$9:$G$10000)+SUMIF(CLASSIFICAÇÃO!$B$6:$B$93,$AB" & linha & _
            ",CLASSIFICAÇÃO!$C$6:$C$93))"

        wsRank.Cells(linha, 32).Formula = _
            "=IF($AA" & linha & "="""","""",SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA" & linha & _
            ",'HISTÓRICO DOS CICLOS'!$H$9:$H$10000)+SUMIF(CLASSIFICAÇÃO!$B$6:$B$93,$AB" & linha & _
            ",CLASSIFICAÇÃO!$D$6:$D$93))"

        wsRank.Cells(linha, 33).Formula = _
            "=IF($AA" & linha & "="""","""",SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA" & linha & _
            ",'HISTÓRICO DOS CICLOS'!$I$9:$I$10000)+SUMIF(CLASSIFICAÇÃO!$B$6:$B$93,$AB" & linha & _
            ",CLASSIFICAÇÃO!$E$6:$E$93))"

        wsRank.Cells(linha, 34).Formula = _
            "=IF($AA" & linha & "="""","""",SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA" & linha & _
            ",'HISTÓRICO DOS CICLOS'!$J$9:$J$10000)+SUMIF(CLASSIFICAÇÃO!$B$6:$B$93,$AB" & linha & _
            ",CLASSIFICAÇÃO!$F$6:$F$93))"

        wsRank.Cells(linha, 35).Formula = _
            "=IF($AA" & linha & "="""","""",SUMIF('HISTÓRICO DOS CICLOS'!$N$9:$N$10000,$AA" & linha & _
            ",'HISTÓRICO DOS CICLOS'!$K$9:$K$10000)+SUMIF(CLASSIFICAÇÃO!$B$6:$B$93,$AB" & linha & _
            ",CLASSIFICAÇÃO!$G$6:$G$93))"

        wsRank.Cells(linha, 36).Formula = _
            "=IF($AA" & linha & "="""","""",$AH" & linha & "-$AI" & linha & ")"

        wsRank.Cells(linha, 37).Formula = _
            "=IF(OR($AA" & linha & "="""",$AE" & linha & "=0),0,$AF" & linha & "/$AE" & linha & ")"

        wsRank.Cells(linha, 38).Formula = _
            "=IF($AA" & linha & "="""","""",1+" & _
            "COUNTIF($AF$2:$AF$53,"">""&$AF" & linha & ")+" & _
            "COUNTIFS($AF$2:$AF$53,$AF" & linha & ",$AH$2:$AH$53,"">""&$AH" & linha & ")+" & _
            "COUNTIFS($AF$2:$AF$53,$AF" & linha & ",$AH$2:$AH$53,$AH" & linha & ",$AI$2:$AI$53,""<""&$AI" & linha & ")+" & _
            "COUNTIFS($AF$2:$AF$53,$AF" & linha & ",$AH$2:$AH$53,$AH" & linha & ",$AI$2:$AI$53,$AI" & linha & ",$AJ$2:$AJ$53,"">""&$AJ" & linha & ")+" & _
            "COUNTIFS($AF$2:$AF$53,$AF" & linha & ",$AH$2:$AH$53,$AH" & linha & ",$AI$2:$AI$53,$AI" & linha & ",$AJ$2:$AJ$53,$AJ" & linha & ",$AB$2:$AB$53,""<""&$AB" & linha & "))"

        wsRank.Cells(linha, 39).Formula = "=IF($AA" & linha & "="""","""",$AL" & linha & ")"

    Next linha

    wsRank.Range("A5:K36").ClearContents
    wsRank.Range("M5:W36").ClearContents

    For linha = 5 To 30

        posicao = linha - 4

        PreencherLinhaRanking wsRank, linha, 1, posicao

        posicao = linha + 22

        PreencherLinhaRanking wsRank, linha, 13, posicao

    Next linha

    wsRank.Range("A5:A30").NumberFormat = "0°"
    wsRank.Range("M5:M30").NumberFormat = "0°"
    wsRank.Range("K5:K30").NumberFormat = "0.0%"
    wsRank.Range("W5:W30").NumberFormat = "0.0%"

End Sub

Private Sub PreencherLinhaRanking(ByVal wsRank As Worksheet, _
                                  ByVal linha As Long, _
                                  ByVal colunaInicial As Long, _
                                  ByVal posicao As Long)

    Dim colPosicao As Long
    Dim colAtleta As Long

    colPosicao = colunaInicial
    colAtleta = colunaInicial + 1

    wsRank.Cells(linha, colPosicao).Formula = _
        "=IF(" & ColunaLetra(colAtleta) & linha & "="""",""""," & posicao & ")"

    wsRank.Cells(linha, colAtleta).Formula = _
        "=IFERROR(INDEX($AB$2:$AB$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 2).Formula = _
        "=IFERROR(INDEX($AC$2:$AC$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 3).Formula = _
        "=IFERROR(INDEX($AD$2:$AD$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 4).Formula = _
        "=IFERROR(INDEX($AE$2:$AE$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 5).Formula = _
        "=IFERROR(INDEX($AF$2:$AF$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 6).Formula = _
        "=IFERROR(INDEX($AG$2:$AG$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 7).Formula = _
        "=IFERROR(INDEX($AH$2:$AH$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 8).Formula = _
        "=IFERROR(INDEX($AI$2:$AI$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 9).Formula = _
        "=IFERROR(INDEX($AJ$2:$AJ$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

    wsRank.Cells(linha, colunaInicial + 10).Formula = _
        "=IFERROR(INDEX($AK$2:$AK$53,MATCH(" & posicao & ",$AL$2:$AL$53,0)),"""")"

End Sub

Private Function CriarMapaDosBlocos() As Object

    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1

    ' Array: coluna do ID, coluna do atleta, coluna do status,
    ' coluna da observação, primeira linha das quatro vagas.

    mapa.Add MontarChave("Avançado Masculino", "GRUPO A"), Array(2, 3, 4, 5, 9)
    mapa.Add MontarChave("Avançado Masculino", "GRUPO B"), Array(8, 9, 10, 11, 9)

    mapa.Add MontarChave("Intermediário Masculino", "GRUPO A"), Array(2, 3, 4, 5, 20)
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO B"), Array(8, 9, 10, 11, 20)
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO C"), Array(14, 15, 16, 17, 20)

    mapa.Add MontarChave("Intermediário Masculino", "GRUPO D"), Array(2, 3, 4, 5, 28)
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO E"), Array(8, 9, 10, 11, 28)
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO F"), Array(14, 15, 16, 17, 28)

    mapa.Add MontarChave("Intermediário Masculino", "GRUPO G"), Array(2, 3, 4, 5, 36)

    mapa.Add MontarChave("Iniciante Masculino", "GRUPO A"), Array(2, 3, 4, 5, 47)
    mapa.Add MontarChave("Iniciante Masculino", "GRUPO B"), Array(8, 9, 10, 11, 47)

    mapa.Add MontarChave("Intermediário Feminino", "GRUPO A"), Array(2, 3, 4, 5, 58)
    mapa.Add MontarChave("Intermediário Feminino", "GRUPO B"), Array(8, 9, 10, 11, 58)

    Set CriarMapaDosBlocos = mapa

End Function

Private Function CriarMapaDaClassificacao() As Object

    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = 1

    mapa.Add MontarChave("Avançado Masculino", "GRUPO A"), 6
    mapa.Add MontarChave("Avançado Masculino", "GRUPO B"), 13

    mapa.Add MontarChave("Intermediário Masculino", "GRUPO A"), 20
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO B"), 27
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO C"), 34
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO D"), 41
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO E"), 48
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO F"), 55
    mapa.Add MontarChave("Intermediário Masculino", "GRUPO G"), 62

    mapa.Add MontarChave("Iniciante Masculino", "GRUPO A"), 69
    mapa.Add MontarChave("Iniciante Masculino", "GRUPO B"), 76

    mapa.Add MontarChave("Intermediário Feminino", "GRUPO A"), 83
    mapa.Add MontarChave("Intermediário Feminino", "GRUPO B"), 90

    Set CriarMapaDaClassificacao = mapa

End Function

Private Function MontarChave(ByVal categoria As String, _
                             ByVal grupo As String) As String

    MontarChave = UCase$(Trim$(categoria)) & "|" & UCase$(Trim$(grupo))

End Function

Private Function MaiorCicloRegistrado(ByVal wsHist As Worksheet) As Long

    Dim ultimaLinha As Long
    Dim resultado As Variant

    ultimaLinha = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row

    If ultimaLinha < PRIMEIRA_LINHA_HISTORICO Then
        MaiorCicloRegistrado = 0
        Exit Function
    End If

    resultado = Application.Max( _
                wsHist.Range("A" & PRIMEIRA_LINHA_HISTORICO & _
                ":A" & ultimaLinha))

    If IsNumeric(resultado) Then
        MaiorCicloRegistrado = CLng(resultado)
    Else
        MaiorCicloRegistrado = 0
    End If

End Function

Private Function ColunaLetra(ByVal numeroColuna As Long) As String

    Dim resultado As String
    Dim numero As Long

    numero = numeroColuna

    Do While numero > 0
        resultado = Chr$(((numero - 1) Mod 26) + 65) & resultado
        numero = (numero - 1) \ 26
    Loop

    ColunaLetra = resultado

End Function
```

## 14.4 `HistoricoLiga`

**Nota:** Arquivador antigo; mapa hard-coded de categorias e lookup por nome.

SHA-256 do código extraído:

`2490c423a2fb34d7f6059df60425f74cf41bb465f68bc4fede90b6054b1f8cea`

```vb
Attribute VB_Name = "HistoricoLiga"
Option Explicit

Public Sub ArquivarCicloNoHistorico()

    Const PRIMEIRA_LINHA_HISTORICO As Long = 9

    Dim wsFim As Worksheet
    Dim wsClass As Worksheet
    Dim wsMov As Worksheet
    Dim wsHist As Worksheet

    Dim grupos As Variant
    Dim item As Variant
    Dim categoria As String
    Dim grupo As String
    Dim linhaInicial As Long
    Dim linhaClass As Long
    Dim linhaDestino As Long
    Dim linhaMov As Variant

    Dim numeroCiclo As Long
    Dim dataEncerramento As Date
    Dim atleta As String
    Dim primeiraLinhaInserida As Long
    Dim ultimaLinhaInserida As Long
    Dim quantidade As Long

    On Error GoTo TratarErro

    Set wsFim = ThisWorkbook.Worksheets("FIM DE CICLO")
    Set wsClass = ThisWorkbook.Worksheets("CLASSIFICAÇÃO")
    Set wsMov = ThisWorkbook.Worksheets("SUBIDAS E DESCIDAS")
    Set wsHist = ThisWorkbook.Worksheets("HISTÓRICO DOS CICLOS")

    If UCase$(Trim$(CStr(wsFim.Range("B18").Value))) <> "SIM" Then
        MsgBox "Digite SIM em FIM DE CICLO antes de arquivar.", vbExclamation, "Liga Padel"
        Exit Sub
    End If

    If Not IsNumeric(wsHist.Range("B5").Value) Or wsHist.Range("B5").Value < 1 Then
        MsgBox "Informe um número de ciclo válido em HISTÓRICO DOS CICLOS!B5.", vbExclamation, "Liga Padel"
        Exit Sub
    End If

    numeroCiclo = CLng(wsHist.Range("B5").Value)

    If IsDate(wsHist.Range("B6").Value) Then
        dataEncerramento = CDate(wsHist.Range("B6").Value)
    Else
        dataEncerramento = Date
    End If

    If Application.WorksheetFunction.CountIf( _
        wsHist.Range("A" & PRIMEIRA_LINHA_HISTORICO & ":A" & wsHist.Rows.Count), _
        numeroCiclo) > 0 Then

        MsgBox "O ciclo " & numeroCiclo & " já está registrado no histórico." & vbCrLf & _
               "Nenhum dado foi duplicado.", vbInformation, "Liga Padel"
        Exit Sub
    End If

    grupos = Array( _
        Array("Avançado Masculino", "GRUPO A", 6), _
        Array("Avançado Masculino", "GRUPO B", 13), _
        Array("Intermediário Masculino", "GRUPO A", 20), _
        Array("Intermediário Masculino", "GRUPO B", 27), _
        Array("Intermediário Masculino", "GRUPO C", 34), _
        Array("Intermediário Masculino", "GRUPO D", 41), _
        Array("Intermediário Masculino", "GRUPO E", 48), _
        Array("Intermediário Masculino", "GRUPO F", 55), _
        Array("Intermediário Masculino", "GRUPO G", 62), _
        Array("Iniciante Masculino", "GRUPO A", 69), _
        Array("Iniciante Masculino", "GRUPO B", 76), _
        Array("Intermediário Feminino", "GRUPO A", 83), _
        Array("Intermediário Feminino", "GRUPO B", 90) _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    linhaDestino = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
    If linhaDestino < PRIMEIRA_LINHA_HISTORICO Then linhaDestino = PRIMEIRA_LINHA_HISTORICO

    primeiraLinhaInserida = linhaDestino

    For Each item In grupos

        categoria = CStr(item(0))
        grupo = CStr(item(1))
        linhaInicial = CLng(item(2))

        For linhaClass = linhaInicial To linhaInicial + 3

            atleta = Trim$(CStr(wsClass.Cells(linhaClass, "B").Value))

            If atleta <> "" Then

                wsHist.Cells(linhaDestino, "A").Value = numeroCiclo
                wsHist.Cells(linhaDestino, "B").Value = dataEncerramento
                wsHist.Cells(linhaDestino, "C").Value = categoria
                wsHist.Cells(linhaDestino, "D").Value = grupo
                wsHist.Cells(linhaDestino, "E").Value = wsClass.Cells(linhaClass, "A").Value
                wsHist.Cells(linhaDestino, "F").Value = atleta
                wsHist.Cells(linhaDestino, "G").Value = wsClass.Cells(linhaClass, "C").Value
                wsHist.Cells(linhaDestino, "H").Value = wsClass.Cells(linhaClass, "D").Value
                wsHist.Cells(linhaDestino, "I").Value = wsClass.Cells(linhaClass, "E").Value
                wsHist.Cells(linhaDestino, "J").Value = wsClass.Cells(linhaClass, "F").Value
                wsHist.Cells(linhaDestino, "K").Value = wsClass.Cells(linhaClass, "G").Value
                wsHist.Cells(linhaDestino, "L").Value = wsClass.Cells(linhaClass, "H").Value

                linhaMov = Application.Match(atleta, wsMov.Range("B6:B57"), 0)

                If IsError(linhaMov) Then
                    wsHist.Cells(linhaDestino, "M").Value = ""
                Else
                    wsHist.Cells(linhaDestino, "M").Value = _
                        wsMov.Cells(5 + CLng(linhaMov), "I").Value
                End If

                linhaDestino = linhaDestino + 1
                quantidade = quantidade + 1

            End If
        Next linhaClass
    Next item

    If quantidade = 0 Then
        MsgBox "Nenhum atleta foi encontrado na CLASSIFICAÇÃO." & vbCrLf & _
               "Nada foi registrado.", vbExclamation, "Liga Padel"
        GoTo Finalizar
    End If

    ultimaLinhaInserida = linhaDestino - 1

    wsHist.Range("B" & primeiraLinhaInserida & ":B" & ultimaLinhaInserida).NumberFormat = "dd/mm/yyyy"
    wsHist.Range("E" & primeiraLinhaInserida & ":E" & ultimaLinhaInserida).NumberFormat = "0°"

    MsgBox quantidade & " atletas do ciclo " & numeroCiclo & _
           " foram registrados no histórico." & vbCrLf & _
           "Nenhuma outra aba foi limpa ou alterada.", _
           vbInformation, "Liga Padel"

Finalizar:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

TratarErro:
    On Error Resume Next

    If primeiraLinhaInserida > 0 And linhaDestino > primeiraLinhaInserida Then
        wsHist.Range("A" & primeiraLinhaInserida & ":M" & linhaDestino - 1).ClearContents
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "A macro encontrou um erro e cancelou o arquivamento." & vbCrLf & _
           "Nenhum registro parcial foi mantido." & vbCrLf & vbCrLf & _
           "Detalhe: " & Err.Description, vbCritical, "Liga Padel"

End Sub
```

## 14.5 `ReiniciarLigaOficial`

**Nota:** Reset destrutivo legado.

SHA-256 do código extraído:

`a752aa5972d8c0e8804a8a7b38cdaa6f59c4e938a4110d48c51521224ce997f1`

```vb
Attribute VB_Name = "ReiniciarLigaOficial"
Option Explicit

Public Sub ReiniciarParaLigaOficial()

    Const PRIMEIRA_LINHA_CADASTRO As Long = 3
    Const ULTIMA_LINHA_CADASTRO As Long = 200
    Const PRIMEIRA_LINHA_HISTORICO As Long = 9
    Const ULTIMA_LINHA_HISTORICO As Long = 10000
    Const ULTIMA_LINHA_RESULTADOS As Long = 80

    Dim wsCadastro As Worksheet
    Dim wsGrupos As Worksheet
    Dim wsResultados As Worksheet
    Dim wsMovimentacao As Worksheet
    Dim wsFim As Worksheet
    Dim wsHistorico As Worksheet

    Dim resposta As VbMsgBoxResult
    Dim confirmacaoDigitada As String
    Dim quantidadeCadastro As Long
    Dim quantidadeHistorico As Long
    Dim cicloAtual As Variant

    Dim nomeArquivo As String
    Dim nomeSemExtensao As String
    Dim extensao As String
    Dim posicaoPonto As Long
    Dim caminhoBackup As String

    Dim blocos As Variant
    Dim bloco As Variant
    Dim colunaId As Long
    Dim primeiraLinhaGrupo As Long
    Dim primeiraLinhaCadastroGrupo As Long
    Dim vaga As Long
    Dim linhaCadastro As Long
    Dim linhaGrupo As Long
    Dim linha As Long

    Dim calculoAnterior As XlCalculation
    Dim eventosAnteriores As Boolean
    Dim telaAnterior As Boolean
    Dim statusAnterior As Variant

    Dim backupCriado As Boolean
    Dim mensagemErro As String

    On Error GoTo TratarErro

    Set wsCadastro = ThisWorkbook.Worksheets("CADASTRO DE ATLETAS")
    Set wsGrupos = ThisWorkbook.Worksheets("GRUPOS")
    Set wsResultados = ThisWorkbook.Worksheets("LANÇAR RESULTADOS")
    Set wsMovimentacao = ThisWorkbook.Worksheets("SUBIDAS E DESCIDAS")
    Set wsFim = ThisWorkbook.Worksheets("FIM DE CICLO")
    Set wsHistorico = ThisWorkbook.Worksheets("HISTÓRICO DOS CICLOS")

    If ThisWorkbook.Path = "" Then
        MsgBox "Salve a planilha antes de executar a reinicialização.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    posicaoPonto = InStrRev(ThisWorkbook.Name, ".")

    If posicaoPonto = 0 Then
        MsgBox "Não foi possível identificar o formato do arquivo.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    extensao = LCase$(Mid$(ThisWorkbook.Name, posicaoPonto))

    If extensao <> ".xlsm" And extensao <> ".xlsb" Then
        MsgBox "Esta operação deve ser executada em uma planilha habilitada para macros (.xlsm)." & _
               vbCrLf & "Salve corretamente o arquivo antes de continuar.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    quantidadeCadastro = Application.WorksheetFunction.CountA( _
                         wsCadastro.Range("A" & PRIMEIRA_LINHA_CADASTRO & _
                         ":G" & ULTIMA_LINHA_CADASTRO))

    quantidadeHistorico = Application.WorksheetFunction.CountA( _
                          wsHistorico.Range("A" & PRIMEIRA_LINHA_HISTORICO & _
                          ":N" & ULTIMA_LINHA_HISTORICO))

    cicloAtual = wsHistorico.Range("B5").Value

    resposta = MsgBox( _
        "ATENÇÃO: esta função preparará uma base limpa para a Liga Oficial." & vbCrLf & vbCrLf & _
        "Situação atual encontrada:" & vbCrLf & _
        "• Ciclo indicado: " & CStr(cicloAtual) & vbCrLf & _
        "• Células preenchidas no cadastro: " & quantidadeCadastro & vbCrLf & _
        "• Células preenchidas no histórico: " & quantidadeHistorico & vbCrLf & vbCrLf & _
        "A operação irá apagar os atletas, placares e histórico de teste." & vbCrLf & _
        "As fórmulas, formatações e macros serão preservadas." & vbCrLf & vbCrLf & _
        "Antes de apagar qualquer dado, será criada automaticamente uma cópia de segurança." & vbCrLf & vbCrLf & _
        "Deseja continuar?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Preparar Liga Oficial")

    If resposta <> vbYes Then Exit Sub

    confirmacaoDigitada = InputBox( _
        "Para confirmar a reinicialização, digite exatamente:" & vbCrLf & vbCrLf & _
        "REINICIAR LIGA OFICIAL" & vbCrLf & vbCrLf & _
        "Qualquer outro texto cancelará a operação.", _
        "Confirmação de segurança")

    If UCase$(Trim$(confirmacaoDigitada)) <> "REINICIAR LIGA OFICIAL" Then
        MsgBox "Operação cancelada. Nenhum dado foi alterado.", _
               vbInformation, "Liga Padel"
        Exit Sub
    End If

    nomeArquivo = ThisWorkbook.Name
    nomeSemExtensao = Left$(nomeArquivo, posicaoPonto - 1)

    caminhoBackup = ThisWorkbook.Path & Application.PathSeparator & _
                    nomeSemExtensao & "_BACKUP_ANTES_REINICIO_" & _
                    Format$(Now, "yyyymmdd_hhnnss") & extensao

    ' A cópia é criada antes da primeira alteração destrutiva.
    ThisWorkbook.SaveCopyAs caminhoBackup
    backupCriado = True

    calculoAnterior = Application.Calculation
    eventosAnteriores = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating
    statusAnterior = Application.StatusBar

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando a base oficial da Liga Padel..."

    ' ------------------------------------------------------------------
    ' 1. LIMPA SOMENTE OS DADOS DOS ATLETAS DE TESTE.
    ' Cabeçalhos, formatação e validações permanecem.
    ' ------------------------------------------------------------------
    wsCadastro.Range("A3:G200").ClearContents

    ' ------------------------------------------------------------------
    ' 2. RESTAURA A DISTRIBUIÇÃO INICIAL DOS GRUPOS.
    '
    ' Cada conjunto contém:
    ' coluna do ID na aba GRUPOS,
    ' primeira linha das quatro vagas,
    ' primeira linha correspondente no cadastro.
    '
    ' Capacidade atual: 13 grupos de quatro atletas = 52 atletas.
    ' ------------------------------------------------------------------
    blocos = Array( _
        Array(2, 9, 3), _
        Array(8, 9, 7), _
        Array(2, 20, 11), _
        Array(8, 20, 15), _
        Array(14, 20, 19), _
        Array(2, 28, 23), _
        Array(8, 28, 27), _
        Array(14, 28, 31), _
        Array(2, 36, 35), _
        Array(2, 47, 39), _
        Array(8, 47, 43), _
        Array(2, 58, 47), _
        Array(8, 58, 51) _
    )

    For Each bloco In blocos

        colunaId = CLng(bloco(0))
        primeiraLinhaGrupo = CLng(bloco(1))
        primeiraLinhaCadastroGrupo = CLng(bloco(2))

        For vaga = 0 To 3

            linhaGrupo = primeiraLinhaGrupo + vaga
            linhaCadastro = primeiraLinhaCadastroGrupo + vaga

            wsGrupos.Cells(linhaGrupo, colunaId).Formula = _
                FormulaCadastroVazioSeSemDados("A", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 1).Formula = _
                FormulaCadastroVazioSeSemDados("B", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 2).Formula = _
                FormulaCadastroVazioSeSemDados("E", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 3).Formula = _
                FormulaCadastroVazioSeSemDados("G", linhaCadastro)

        Next vaga
    Next bloco

    ' ------------------------------------------------------------------
    ' 3. LIMPA SOMENTE OS PLACARES.
    ' Os títulos, duplas automáticas e fórmulas permanecem.
    ' ------------------------------------------------------------------
    For linha = 1 To ULTIMA_LINHA_RESULTADOS

        If IsNumeric(wsResultados.Cells(linha, 3).Value) Then

            If CLng(wsResultados.Cells(linha, 3).Value) >= 1 And _
               CLng(wsResultados.Cells(linha, 3).Value) <= 3 Then

                wsResultados.Cells(linha, 6).ClearContents
                wsResultados.Cells(linha, 7).ClearContents

            End If
        End If

    Next linha

    ' ------------------------------------------------------------------
    ' 4. LIMPA SOMENTE OS CAMPOS MANUAIS DA MOVIMENTAÇÃO.
    ' As sugestões e os demais cálculos continuam por fórmula.
    ' ------------------------------------------------------------------
    wsMovimentacao.Range("G6:G57").ClearContents
    wsMovimentacao.Range("J6:J57").ClearContents

    ' ------------------------------------------------------------------
    ' 5. APAGA O HISTÓRICO DE TESTE E VOLTA AO CICLO 1.
    ' O cabeçalho e a estrutura da aba permanecem.
    ' ------------------------------------------------------------------
    wsHistorico.Range("A9:N10000").ClearContents
    wsHistorico.Range("B5").Value = 1
    wsHistorico.Range("B6").ClearContents

    If Trim$(CStr(wsHistorico.Range("N8").Value)) = "" Then
        wsHistorico.Range("M8").Copy
        wsHistorico.Range("N8").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsHistorico.Range("N8").Value = "ID ÚNICO"
        wsHistorico.Columns("N").ColumnWidth = 14
    End If

    ' ------------------------------------------------------------------
    ' 6. RETIRA A CONFIRMAÇÃO DO FECHAMENTO.
    ' A próxima rodada voltará a aparecer somente após o ciclo oficial.
    ' ------------------------------------------------------------------
    wsFim.Range("B18").Value = "NÃO"

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    wsCadastro.Activate
    wsCadastro.Range("A3").Select

    ThisWorkbook.Save

    Application.StatusBar = False
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    MsgBox _
        "A base oficial foi preparada com sucesso." & vbCrLf & vbCrLf & _
        "A planilha voltou ao CICLO 1 e está pronta para cadastrar os atletas reais." & vbCrLf & _
        "Os dados e resultados de teste foram removidos." & vbCrLf & _
        "As fórmulas, formatações e macros foram preservadas." & vbCrLf & vbCrLf & _
        "Cópia de segurança criada em:" & vbCrLf & _
        caminhoBackup, _
        vbInformation, "Liga Padel"

    Exit Sub

TratarErro:

    mensagemErro = Err.Description

    On Error Resume Next
    Application.StatusBar = False
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    If backupCriado Then
        MsgBox _
            "A reinicialização encontrou um erro e foi interrompida." & vbCrLf & vbCrLf & _
            "Uma cópia de segurança foi criada antes das alterações:" & vbCrLf & _
            caminhoBackup & vbCrLf & vbCrLf & _
            "Feche a planilha sem salvar e abra essa cópia de segurança." & vbCrLf & vbCrLf & _
            "Detalhe do erro: " & mensagemErro, _
            vbCritical, "Liga Padel"
    Else
        MsgBox _
            "A reinicialização não foi iniciada." & vbCrLf & _
            "Nenhum dado deveria ter sido apagado." & vbCrLf & vbCrLf & _
            "Detalhe do erro: " & mensagemErro, _
            vbCritical, "Liga Padel"
    End If

End Sub

Private Function FormulaCadastroVazioSeSemDados( _
    ByVal letraColuna As String, _
    ByVal linhaCadastro As Long) As String

    FormulaCadastroVazioSeSemDados = _
        "=IF('CADASTRO DE ATLETAS'!" & letraColuna & linhaCadastro & _
        "="""","""",'CADASTRO DE ATLETAS'!" & letraColuna & linhaCadastro & ")"

End Function
```

## 14.6 `ReiniciarLigaOficialV2`

**Nota:** Reset destrutivo legado com backup mais seguro.

SHA-256 do código extraído:

`3274b60907de7f237004ca4b478a46492db16d21987882f495ae0701e24a5f32`

```vb
Attribute VB_Name = "ReiniciarLigaOficialV2"
Option Explicit

Public Sub ReiniciarParaLigaOficialV2()

    Const PRIMEIRA_LINHA_CADASTRO As Long = 3
    Const ULTIMA_LINHA_CADASTRO As Long = 200
    Const PRIMEIRA_LINHA_HISTORICO As Long = 9
    Const ULTIMA_LINHA_HISTORICO As Long = 10000
    Const ULTIMA_LINHA_RESULTADOS As Long = 80

    Dim wsCadastro As Worksheet
    Dim wsGrupos As Worksheet
    Dim wsResultados As Worksheet
    Dim wsMovimentacao As Worksheet
    Dim wsFim As Worksheet
    Dim wsHistorico As Worksheet

    Dim resposta As VbMsgBoxResult
    Dim confirmacaoDigitada As String
    Dim quantidadeCadastro As Long
    Dim quantidadeHistorico As Long
    Dim cicloAtual As Variant

    Dim nomeArquivo As String
    Dim nomeSemExtensao As String
    Dim extensao As String
    Dim posicaoPonto As Long
    Dim caminhoBackup As String
    Dim nomeBackup As String
    Dim pastaPadrao As String
    Dim selecaoBackup As Variant

    Dim blocos As Variant
    Dim bloco As Variant
    Dim colunaId As Long
    Dim primeiraLinhaGrupo As Long
    Dim primeiraLinhaCadastroGrupo As Long
    Dim vaga As Long
    Dim linhaCadastro As Long
    Dim linhaGrupo As Long
    Dim linha As Long

    Dim calculoAnterior As XlCalculation
    Dim eventosAnteriores As Boolean
    Dim telaAnterior As Boolean
    Dim statusAnterior As Variant

    Dim backupCriado As Boolean
    Dim mensagemErro As String

    On Error GoTo TratarErro

    Set wsCadastro = ThisWorkbook.Worksheets("CADASTRO DE ATLETAS")
    Set wsGrupos = ThisWorkbook.Worksheets("GRUPOS")
    Set wsResultados = ThisWorkbook.Worksheets("LANÇAR RESULTADOS")
    Set wsMovimentacao = ThisWorkbook.Worksheets("SUBIDAS E DESCIDAS")
    Set wsFim = ThisWorkbook.Worksheets("FIM DE CICLO")
    Set wsHistorico = ThisWorkbook.Worksheets("HISTÓRICO DOS CICLOS")

    If ThisWorkbook.Path = "" Then
        MsgBox "Salve a planilha antes de executar a reinicialização.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    posicaoPonto = InStrRev(ThisWorkbook.Name, ".")

    If posicaoPonto = 0 Then
        MsgBox "Não foi possível identificar o formato do arquivo.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    extensao = LCase$(Mid$(ThisWorkbook.Name, posicaoPonto))

    If extensao <> ".xlsm" And extensao <> ".xlsb" Then
        MsgBox "Esta operação deve ser executada em uma planilha habilitada para macros (.xlsm)." & _
               vbCrLf & "Salve corretamente o arquivo antes de continuar.", _
               vbExclamation, "Liga Padel"
        Exit Sub
    End If

    quantidadeCadastro = Application.WorksheetFunction.CountA( _
                         wsCadastro.Range("A" & PRIMEIRA_LINHA_CADASTRO & _
                         ":G" & ULTIMA_LINHA_CADASTRO))

    quantidadeHistorico = Application.WorksheetFunction.CountA( _
                          wsHistorico.Range("A" & PRIMEIRA_LINHA_HISTORICO & _
                          ":N" & ULTIMA_LINHA_HISTORICO))

    cicloAtual = wsHistorico.Range("B5").Value

    resposta = MsgBox( _
        "ATENÇÃO: esta função preparará uma base limpa para a Liga Oficial." & vbCrLf & vbCrLf & _
        "Situação atual encontrada:" & vbCrLf & _
        "• Ciclo indicado: " & CStr(cicloAtual) & vbCrLf & _
        "• Células preenchidas no cadastro: " & quantidadeCadastro & vbCrLf & _
        "• Células preenchidas no histórico: " & quantidadeHistorico & vbCrLf & vbCrLf & _
        "A operação irá apagar os atletas, placares e histórico de teste." & vbCrLf & _
        "As fórmulas, formatações e macros serão preservadas." & vbCrLf & vbCrLf & _
        "Antes de apagar qualquer dado, será criada automaticamente uma cópia de segurança." & vbCrLf & vbCrLf & _
        "Deseja continuar?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Preparar Liga Oficial — V2")

    If resposta <> vbYes Then Exit Sub

    confirmacaoDigitada = InputBox( _
        "Para confirmar a reinicialização, digite exatamente:" & vbCrLf & vbCrLf & _
        "REINICIAR LIGA OFICIAL" & vbCrLf & vbCrLf & _
        "Qualquer outro texto cancelará a operação.", _
        "Confirmação de segurança")

    If UCase$(Trim$(confirmacaoDigitada)) <> "REINICIAR LIGA OFICIAL" Then
        MsgBox "Operação cancelada. Nenhum dado foi alterado.", _
               vbInformation, "Liga Padel"
        Exit Sub
    End If

    nomeArquivo = ThisWorkbook.Name
    nomeSemExtensao = Left$(nomeArquivo, posicaoPonto - 1)

    nomeBackup = nomeSemExtensao & "_BACKUP_ANTES_REINICIO_" & _
                 Format$(Now, "yyyymmdd_hhnnss") & extensao

    ' Quando a planilha está aberta pelo OneDrive, ThisWorkbook.Path pode
    ' retornar um endereço https://d.docs.live.net. SaveCopyAs não consegue
    ' gravar diretamente nesse endereço. Por isso, o usuário escolhe um
    ' caminho local antes de qualquer alteração.
    pastaPadrao = Environ$("USERPROFILE") & "\Downloads"

    If Dir$(pastaPadrao, vbDirectory) = "" Then
        pastaPadrao = Environ$("TEMP")
    End If

    selecaoBackup = Application.GetSaveAsFilename( _
        InitialFileName:=pastaPadrao & Application.PathSeparator & nomeBackup, _
        FileFilter:="Pasta de Trabalho Habilitada para Macro (*.xlsm), *.xlsm", _
        Title:="Escolha onde salvar o backup antes da reinicialização")

    If VarType(selecaoBackup) = vbBoolean Then
        If selecaoBackup = False Then
            MsgBox "Operação cancelada. Nenhum dado foi alterado.", _
                   vbInformation, "Liga Padel"
            Exit Sub
        End If
    End If

    caminhoBackup = CStr(selecaoBackup)

    If LCase$(Right$(caminhoBackup, Len(extensao))) <> extensao Then
        caminhoBackup = caminhoBackup & extensao
    End If

    ' A cópia é criada e conferida antes da primeira alteração destrutiva.
    ThisWorkbook.SaveCopyAs caminhoBackup

    If Dir$(caminhoBackup) = "" Then
        Err.Raise vbObjectError + 3001, , _
                  "O Excel não conseguiu confirmar a criação do backup em: " & caminhoBackup
    End If

    backupCriado = True

    calculoAnterior = Application.Calculation
    eventosAnteriores = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating
    statusAnterior = Application.StatusBar

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando a base oficial da Liga Padel..."

    ' ------------------------------------------------------------------
    ' 1. LIMPA SOMENTE OS DADOS DOS ATLETAS DE TESTE.
    ' Cabeçalhos, formatação e validações permanecem.
    ' ------------------------------------------------------------------
    wsCadastro.Range("A3:G200").ClearContents

    ' ------------------------------------------------------------------
    ' 2. RESTAURA A DISTRIBUIÇÃO INICIAL DOS GRUPOS.
    '
    ' Cada conjunto contém:
    ' coluna do ID na aba GRUPOS,
    ' primeira linha das quatro vagas,
    ' primeira linha correspondente no cadastro.
    '
    ' Capacidade atual: 13 grupos de quatro atletas = 52 atletas.
    ' ------------------------------------------------------------------
    blocos = Array( _
        Array(2, 9, 3), _
        Array(8, 9, 7), _
        Array(2, 20, 11), _
        Array(8, 20, 15), _
        Array(14, 20, 19), _
        Array(2, 28, 23), _
        Array(8, 28, 27), _
        Array(14, 28, 31), _
        Array(2, 36, 35), _
        Array(2, 47, 39), _
        Array(8, 47, 43), _
        Array(2, 58, 47), _
        Array(8, 58, 51) _
    )

    For Each bloco In blocos

        colunaId = CLng(bloco(0))
        primeiraLinhaGrupo = CLng(bloco(1))
        primeiraLinhaCadastroGrupo = CLng(bloco(2))

        For vaga = 0 To 3

            linhaGrupo = primeiraLinhaGrupo + vaga
            linhaCadastro = primeiraLinhaCadastroGrupo + vaga

            wsGrupos.Cells(linhaGrupo, colunaId).Formula = _
                FormulaCadastroVazioSeSemDados("A", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 1).Formula = _
                FormulaCadastroVazioSeSemDados("B", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 2).Formula = _
                FormulaCadastroVazioSeSemDados("E", linhaCadastro)

            wsGrupos.Cells(linhaGrupo, colunaId + 3).Formula = _
                FormulaCadastroVazioSeSemDados("G", linhaCadastro)

        Next vaga
    Next bloco

    ' ------------------------------------------------------------------
    ' 3. LIMPA SOMENTE OS PLACARES.
    ' Os títulos, duplas automáticas e fórmulas permanecem.
    ' ------------------------------------------------------------------
    For linha = 1 To ULTIMA_LINHA_RESULTADOS

        If IsNumeric(wsResultados.Cells(linha, 3).Value) Then

            If CLng(wsResultados.Cells(linha, 3).Value) >= 1 And _
               CLng(wsResultados.Cells(linha, 3).Value) <= 3 Then

                wsResultados.Cells(linha, 6).ClearContents
                wsResultados.Cells(linha, 7).ClearContents

            End If
        End If

    Next linha

    ' ------------------------------------------------------------------
    ' 4. LIMPA SOMENTE OS CAMPOS MANUAIS DA MOVIMENTAÇÃO.
    ' As sugestões e os demais cálculos continuam por fórmula.
    ' ------------------------------------------------------------------
    wsMovimentacao.Range("G6:G57").ClearContents
    wsMovimentacao.Range("J6:J57").ClearContents

    ' ------------------------------------------------------------------
    ' 5. APAGA O HISTÓRICO DE TESTE E VOLTA AO CICLO 1.
    ' O cabeçalho e a estrutura da aba permanecem.
    ' ------------------------------------------------------------------
    wsHistorico.Range("A9:N10000").ClearContents
    wsHistorico.Range("B5").Value = 1
    wsHistorico.Range("B6").ClearContents

    If Trim$(CStr(wsHistorico.Range("N8").Value)) = "" Then
        wsHistorico.Range("M8").Copy
        wsHistorico.Range("N8").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsHistorico.Range("N8").Value = "ID ÚNICO"
        wsHistorico.Columns("N").ColumnWidth = 14
    End If

    ' ------------------------------------------------------------------
    ' 6. RETIRA A CONFIRMAÇÃO DO FECHAMENTO.
    ' A próxima rodada voltará a aparecer somente após o ciclo oficial.
    ' ------------------------------------------------------------------
    wsFim.Range("B18").Value = "NÃO"

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    wsCadastro.Activate
    wsCadastro.Range("A3").Select

    ThisWorkbook.Save

    Application.StatusBar = False
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    MsgBox _
        "A base oficial foi preparada com sucesso." & vbCrLf & vbCrLf & _
        "A planilha voltou ao CICLO 1 e está pronta para cadastrar os atletas reais." & vbCrLf & _
        "Os dados e resultados de teste foram removidos." & vbCrLf & _
        "As fórmulas, formatações e macros foram preservadas." & vbCrLf & vbCrLf & _
        "Cópia de segurança criada em:" & vbCrLf & _
        caminhoBackup, _
        vbInformation, "Liga Padel"

    Exit Sub

TratarErro:

    mensagemErro = Err.Description

    On Error Resume Next
    Application.StatusBar = False
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior

    If backupCriado Then
        MsgBox _
            "A reinicialização encontrou um erro e foi interrompida." & vbCrLf & vbCrLf & _
            "Uma cópia de segurança foi criada antes das alterações:" & vbCrLf & _
            caminhoBackup & vbCrLf & vbCrLf & _
            "Feche a planilha sem salvar e abra essa cópia de segurança." & vbCrLf & vbCrLf & _
            "Detalhe do erro: " & mensagemErro, _
            vbCritical, "Liga Padel"
    Else
        MsgBox _
            "A reinicialização não foi iniciada." & vbCrLf & _
            "Nenhum dado deveria ter sido apagado." & vbCrLf & vbCrLf & _
            "Detalhe do erro: " & mensagemErro, _
            vbCritical, "Liga Padel"
    End If

End Sub

Private Function FormulaCadastroVazioSeSemDados( _
    ByVal letraColuna As String, _
    ByVal linhaCadastro As Long) As String

    FormulaCadastroVazioSeSemDados = _
        "=IF('CADASTRO DE ATLETAS'!" & letraColuna & linhaCadastro & _
        "="""","""",'CADASTRO DE ATLETAS'!" & letraColuna & linhaCadastro & ")"

End Function
```

---

# 15. Instrução recomendada para o Claude Code

Ao utilizar este arquivo:

> Consulte o código VBA para entender intenção, validações e casos de segurança, mas não converta as macros literalmente para o backend. A planilha evoluiu depois que alguns módulos foram escritos, portanto estruturas hard-coded de categorias, grupos e linhas podem estar obsoletas. Considere `ESPECIFICACAO_TECNICA_LIGA_PADEL.md` a referência principal das regras atuais; use este documento para complementar comportamento de automações e para extrair requisitos de segurança/transação. A regra de ORDEM DE ENTRADA do `Worksheet_Change` é atual e validada. Em qualquer conflito entre macro legada e regra final documentada/testada, sinalize a divergência antes de implementar.

---

# 16. Precisão e limitações desta extração

O código acima foi extraído diretamente do `vbaProject.bin` embutido no `.xlsm`.

Foram encontrados vários módulos de objeto sem lógica além de metadados/atributos. Eles não foram reproduzidos porque não contêm regras de negócio.

Foram reproduzidos integralmente os módulos com código operacional relevante identificados:

```text
CicloLiga
HistoricoLiga
IDSeguroManual
ReiniciarLigaOficial
ReiniciarLigaOficialV2
Planilha12
```

`Planilha12` foi confirmado como codename da aba `CADASTRO DE ATLETAS` pela estrutura XML do workbook.

Não foi inferida nenhuma regra que não estivesse presente no código ou confirmada no desenvolvimento.
