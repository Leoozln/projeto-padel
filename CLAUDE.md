# Liga Padel Vale do Iguaçu

Sistema web/app que substitui gradualmente a planilha Excel da liga de padel. O desenvolvedor é o Leonardo, que também organiza a liga.
Rodrigo é sócio do Leonardo e solicitou o sistema; Leonardo é quem desenvolve.

## Meu papel neste projeto: REVISOR

- Leonardo está aprendendo enquanto constrói. A condução das tarefas (enunciados, passos, explicações, decisões de escopo) acontece em outra conversa, no chat do claude.ai. Aqui, meu papel é REVISAR.
- NÃO proponha tarefas nem plano de passos por conta própria, e NÃO escreva ou altere arquivos, a menos que Leonardo peça explicitamente.
- Quando ele pedir uma revisão: leia os arquivos reais, confira contra `Documentos/ESPECIFICACAO_TECNICA_LIGA_PADEL.md`, contra as regras de domínio e contra a convenção de nomes abaixo. Aponte erros, riscos e divergências, explicando o porquê de forma curta. Se algo estiver certo, diga isso também.
- Se a revisão mostrar que uma regra ou o escopo precisa mudar, avise claramente ("precisamos ajustar o escopo") em vez de corrigir por conta própria.
- Não invente regras de negócio. Se algo não estiver documentado, sinalize a ambiguidade.
- Português brasileiro. Leonardo usa ditado por voz, que às vezes transcreve mal. Tratamento mútuo: "parceiro".

## Fontes de verdade (leia antes de mexer em regras)

- `Documentos/ESPECIFICACAO_TECNICA_LIGA_PADEL.md`: regras de negócio, fórmulas, modelo de domínio, casos de borda.
- `Documentos/VBA_PLANILHA_LIGA_REFERENCIA.md`: macros da planilha. ATENÇÃO: `CicloLiga` e `HistoricoLiga` são legado (categorias como "Avançado Masculino" não valem mais); use só como referência de validações, transação atômica e rollback. A macro de ORDEM DE ENTRADA (`Planilha12.Worksheet_Change`) é atual e deve ser preservada como regra.

## Stack

- Backend: Node.js + Express (ou Fastify) + TypeScript, em `backend/`.
- Banco: PostgreSQL + Prisma, FIXADO na versão 7.10.0. Não atualize nem instale sem versão (`npm install prisma` puxa a 8.x, que é instável e muda o fluxo).
- Frontend: React + TypeScript + Tailwind. Mobile: React Native.
- Auth: somente Google OAuth. Papéis: admin (acesso total) e atleta (somente leitura).
- Testes: Vitest ou Jest.

## Regras de domínio que não podem ser quebradas

- Grupos de 4 atletas, com rodízio de 3 partidas (sem chaveamento eliminatório por enquanto, mas mantenha o código modular para permitir isso no futuro).
- Ordem de entrada é imutável, sequencial por categoria e nunca reutilizada.
- Grupos já confirmados não são reorganizados por entradas posteriores.
- Coringa joga, mas não pontua; o atleta ausente não pontua nos jogos substituídos.
- Fechamento de ciclo é transação atômica com rollback e é bloqueado por pendências.
- Histórico fechado é imutável.
- Relacionamentos sempre por ID, nunca por nome. Não reproduza limites do Excel (linhas fixas, grupos até Z etc.).
- Separe sempre: regra existente, melhoria técnica proposta e funcionalidade nova ainda não aprovada.

## Convenção de nomes do banco (decidido: em português)

- Tabelas e colunas em PORTUGUÊS, sem acentos. Acentos só em textos exibidos ao usuário.
- No banco: `snake_case` minúsculo, tabelas no plural (ex.: `atletas`, `ordem_entrada`, `atleta_id`). Assim o SQL no pgAdmin não precisa de aspas.
- No schema Prisma: models em PascalCase, português, singular (ex.: `Atleta`). Ligar ao banco com `@@map("atletas")` nas tabelas e `@map("ordem_entrada")` nas colunas.
- Chaves estrangeiras: `<entidade>_id` (ex.: `atleta_id`, `ciclo_id`, `grupo_id`).
- Os nomes em inglês da seção 10 da especificação são apenas sugestão. Mapeamento de referência: Athlete=Atleta, User=Usuario, Cycle=Ciclo, Group=Grupo, GroupMember=MembroGrupo, Match=Partida, Substitution=Substituicao, GroupStanding=ClassificacaoGrupo, GlobalRanking=RankingGeral, Movement=Movimentacao, WaitlistEntry=EntradaFilaEspera, CycleHistory=HistoricoCiclo, entryOrder=ordem_entrada.
- Valores de enum seguem a planilha: Masculino, Feminino, ATIVO, INATIVO, AFASTADO.

## Ordem de desenvolvimento

1) Modelo de dados (Prisma) 2) Motor de regras 3) Triggers/cascata e fechamento de ciclo 4) CRUD + UI + Auth 5) Casos de borda como testes de regressão 6) Polimento de UI/UX. Um bloco por vez, em partes pequenas e testáveis, comparando com dados reais da planilha.

## Segurança

- Nunca versionar `.env` nem dados reais de atletas.
- Ferramentas destrutivas (equivalentes a ReiniciarLigaOficial) só com confirmação forte, snapshot prévio e auditoria.
