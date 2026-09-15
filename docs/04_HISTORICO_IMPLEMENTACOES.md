# HISTÓRICO DE IMPLEMENTAÇÕES
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 04_HISTORICO_IMPLEMENTACOES.md  
**Versão:** 1.0  
**Status:** Oficial  
**Data de criação:** 15/09/2026  

---

# 1. OBJETIVO

Este documento registra o histórico real de desenvolvimento do
Auditor de Planilhas Excel.

Ele deverá permitir identificar rapidamente:

- fase atual;
- fases concluídas;
- funcionalidades implementadas;
- testes executados;
- commits realizados;
- decisões técnicas tomadas;
- problemas encontrados;
- limitações identificadas;
- pendências;
- próximo passo autorizado.

Este documento NÃO é um planejamento.

O planejamento oficial está em:

`03_PLANO_DE_DESENVOLVIMENTO.md`

Aqui devem ser registrados somente fatos relacionados ao desenvolvimento
efetivamente realizado.

---

# 2. DOCUMENTOS OFICIAIS DO PROJETO

A documentação oficial é composta por:

1. `01_ESPECIFICACAO_FUNCIONAL.md`
2. `02_ARQUITETURA.md`
3. `03_PLANO_DE_DESENVOLVIMENTO.md`
4. `04_HISTORICO_IMPLEMENTACOES.md`

Funções:

`01_ESPECIFICACAO_FUNCIONAL.md`
Define os requisitos e regras de negócio.

`02_ARQUITETURA.md`
Define a arquitetura técnica.

`03_PLANO_DE_DESENVOLVIMENTO.md`
Define fases, ordem e critérios de aceite.

`04_HISTORICO_IMPLEMENTACOES.md`
Registra o que realmente aconteceu durante o desenvolvimento.

---

# 3. REGRA DE ATUALIZAÇÃO

Este documento deverá ser atualizado:

- ao concluir uma fase;
- quando ocorrer bloqueio relevante;
- quando uma decisão técnica importante for tomada;
- quando uma limitação real for identificada;
- quando uma fase precisar ser interrompida.

Não é necessário registrar cada pequena alteração de código.

O objetivo é manter um histórico útil e enxuto.

---

# 4. REGRA DE COMMITS

Cada fase possui limite máximo de:

3 commits.

Preferência:

1 ou 2 commits por fase.

Se uma fase atingir 3 commits e ainda não estiver concluída, registrar
o motivo neste documento antes de qualquer decisão de continuidade.

Não ultrapassar o limite silenciosamente.

---

# 5. STATUS POSSÍVEIS

Utilizar:

⬜ NÃO INICIADA

🟡 EM ANDAMENTO

🟢 CONCLUÍDA

🔴 BLOQUEADA

⚠️ CONCLUÍDA COM RESSALVA

---

# 6. VISÃO GERAL

| Fase | Descrição | Status | Commits |
|---|---|---|---:|
| F1 | Fundação e Banco de Auditoria | 🟢 CONCLUÍDA | 2 |
| F2 | Motor Excel e Comparação | 🟢 CONCLUÍDA | 4 |
| F3 | Auditor Local Incremental | 🟢 CONCLUÍDA | 2 |
| F4 | Microsoft Graph / SharePoint | 🔴 BLOQUEADA | 3 |
| F5 | Interface e Relatório | ⬜ NÃO INICIADA | 0 |
| F6 | Robustez e Preparação para Produção | ⬜ NÃO INICIADA | 0 |

---

# 7. PROGRESSO GERAL

Fases concluídas:

3 de 6

Progresso funcional inicial:

Fases 1, 2 e 3 concluídas.

Fase atual:

F4 — Microsoft Graph / SharePoint bloqueada na validação controlada real.

Próxima ação prevista:

Disponibilizar ambiente e credenciais de leitura para concluir a F4.

---

# 8. ESTADO INICIAL

Na criação deste documento:

- documentação funcional definida;
- arquitetura inicial definida;
- plano de desenvolvimento definido;
- histórico de implementação criado;
- implementação da V1 ainda não iniciada.

Nenhuma funcionalidade deverá ser marcada como implementada antes de
existir código e teste correspondente.

---

# 9. HISTÓRICO DA FASE 1

## F1 — Fundação e Banco de Auditoria

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 2 (implementação e encerramento documental)

### Implementado

- fundação executável em Python com ponto de entrada `main.py`;
- configuração por variáveis de ambiente, criação de diretórios e logging;
- banco SQLite criado automaticamente e de forma idempotente;
- seis tabelas oficiais, chaves, relacionamentos, índices e timestamps;
- constraints para identidade técnica, checkpoint único, comparações e
  alterações não duplicadas, estados e tipos controlados;
- encerramento explícito da conexão por gerenciador de contexto;
- testes automatizados de configuração, inicialização, persistência,
  integridade referencial, unicidade, reexecução e encerramento.

### Arquivos criados

- `main.py`, `requirements.txt`, `pytest.ini`, `.gitignore`, `.env.example` e
  `README.md`;
- `app/__init__.py`, `app/config.py`, `app/database.py`, `app/models.py`,
  `app/exceptions.py` e `app/logging_config.py`;
- `tests/test_config.py`, `tests/test_database.py` e `tests/test_main.py`;
- marcadores dos diretórios `data/database`, `data/temp`, `data/reports` e
  `logs`.

### Arquivos alterados

- `docs/04_HISTORICO_IMPLEMENTACOES.md`.

### Banco de dados

Implementado em SQLite, por padrão em `data/database/auditoria.db`, com as
estruturas `planilha`, `checkpoint`, `versao_processada`, `alteracao`,
`execucao_auditoria` e `erro_processamento`.

A identidade de `planilha` utiliza a composição `site_id`, `drive_id` e
`drive_item_id`; o nome permanece descritivo. Essa composição preserva o
contexto técnico até a validação real prevista para a Fase 4.

### Testes executados

Comando:

`pytest -q`

Resultado:

`8 passed in 0.08s`

Comando:

`AUDIT_DATABASE_PATH=/tmp/auditoria-f1.db AUDIT_LOG_PATH=/tmp/auditoria-f1.log python main.py`

Resultado:

aplicação finalizada com código 0; banco e log não vazios foram criados nos
caminhos configurados.

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

### Critérios de aceite

[x] projeto Python executável;

[x] nenhuma dependência de Node.js;

[x] banco SQLite criado automaticamente;

[x] todas as tabelas oficiais existentes;

[x] relacionamentos básicos e integridade referencial funcionais;

[x] proteção essencial contra duplicidade implementada;

[x] testes da camada de persistência aprovados;

[x] histórico atualizado.

### Commits

`4e80575` — Implementa fundação e banco da trilha de auditoria.

O segundo commit encerra a fase com esta atualização factual do histórico; seu
identificador é informado no relatório da execução, pois um commit não pode
registrar o próprio hash em seu conteúdo.

### Decisões técnicas

- utilização exclusiva da biblioteca padrão `sqlite3` na persistência, evitando
  dependência e abstração prematura;
- habilitação de `PRAGMA foreign_keys` em toda conexão;
- criação idempotente por `CREATE TABLE/INDEX IF NOT EXISTS`;
- credenciais não são carregadas nem necessárias na Fase 1.

### Problemas encontrados

A primeira coleta de testes falhou porque o executável `pytest` do ambiente não
incluiu a raiz do projeto no caminho de importação. Foi adicionado `pytest.ini`
com `pythonpath = .`; a execução posterior passou integralmente.

Nenhum bloqueio permanece.

### Pendências

Nenhuma pendência da Fase 1.

### Próximo passo

F2 — Motor Excel e Comparação, somente após autorização do responsável pelo
projeto.

---

# 10. HISTÓRICO DA FASE 2

## F2 — Motor Excel e Comparação

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 4 (incluindo ajuste corretivo solicitado após a
integração da fase)

### Implementado

- leitor `.xlsx` somente leitura com `openpyxl`, `data_only=False`, fechamento
  garantido da pasta de trabalho e snapshots por aba/endereço;
- descarte exclusivo de valores `None`, preservando zero e `False`;
- comparador puro com resultados imutáveis para ADD, DEL e MOD;
- representação determinística de abas adicionadas/removidas pelas alterações de
  suas células e ordenação estável por aba, linha e coluna;
- distinção explícita entre booleanos e números;
- geração temporária de quatro versões `.xlsx` controladas, sem binários
  versionados, cobrindo fórmulas, múltiplas abas, versão sem diferenças e os
  tipos de alteração obrigatórios;
- descarte explícito dessas versões ao encerrar a fixture e regra de
  `.gitignore` limitada à árvore de testes, sem bloquear planilhas reais em
  outros diretórios;
- testes automatizados do reader e comparator.

### Arquivos criados

- `app/excel/__init__.py`;
- `app/excel/reader.py`;
- `app/excel/comparator.py`;
- `tests/test_reader.py`;
- `tests/test_comparator.py`;
- `tests/conftest.py`.

### Testes executados

Comando:

`pytest -q`

Resultado final após a disponibilização de `openpyxl 3.1.5`:

`17 passed`.

### Critérios de aceite

[x] Reader implementado;

[x] fórmulas preservadas por configuração `data_only=False`;

[x] ADD, DEL e MOD implementados;

[x] zero e False tratados explicitamente;

[x] múltiplas abas e snapshots iguais cobertos por testes;

[x] resultado determinístico implementado e coberto por teste;

[x] testes automatizados executados com as dependências instaladas.

### Commits

`dfafb72` — Substitui fixtures binárias por geração em teste (integra os três
commits originalmente planejados para a Fase 2).

O quarto commit é um ajuste corretivo solicitado após a integração: restringe
o ignore aos testes, garante o descarte explícito dos temporários, alinha a
documentação e conclui a validação. Seu identificador é informado no relatório
da execução, pois um commit não pode registrar o próprio hash em seu conteúdo.

### Problemas encontrados

O bloqueio anterior foi resolvido com a disponibilização de `openpyxl 3.1.5`.
Na primeira execução integral, o teste do reader revelou que a asserção não
incluía `C3 = "Pendente"`, embora esse valor estivesse corretamente presente na
versão 0.85 para formar o cenário DEL seguinte. A expectativa foi corrigida e a
suíte integral passou.

### Pendências

Nenhuma pendência da Fase 2.

### Próximo passo

Aguardar autorização expressa para a F3. Não iniciar a fase seguinte.

---

# 11. HISTÓRICO DA FASE 3

## F3 — Auditor Local Incremental

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 2 (implementação e encerramento documental)

### Objetivo

Comprovar todo o fluxo de auditoria incremental utilizando versões
locais simuladas.

### Implementado

- contrato de fonte independente de tecnologia, com identidades compostas de
  site, drive e DriveItem e modelo normalizado de versão;
- fonte local somente leitura, baseada em ordem explícita e arquivos históricos
  gerados temporariamente nos testes;
- serviço de auditoria com cadastro/atualização da planilha, seleção a partir do
  checkpoint e preservação da versão-base;
- comparações consecutivas mantendo somente os snapshots necessários em memória;
- persistência atômica, por comparação, da versão processada, alterações e
  checkpoint;
- registro de versões sem diferenças, execuções sem novidades e falhas;
- rollback da comparação com falha, preservação das comparações confirmadas e
  retomada posterior exata do checkpoint;
- testes de primeira auditoria, idempotência, incremento, falha e retomada.

### Testes executados

Comando:

`pytest -q`

Resultado:

`22 passed in 1.00s`

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

### Critérios de aceite

[x] fonte local funciona;

[x] primeira auditoria e consolidação do histórico funcionam;

[x] checkpoint e preservação da versão-base funcionam;

[x] reexecução não duplica registros;

[x] auditoria incremental processa somente novas comparações;

[x] falha não avança o checkpoint incorretamente e permite retomada;

[x] versão sem alterações é registrada;

[x] histórico de execução e erro é criado;

[x] testes passam.

### Commits

`53b29ec` — Implementa auditoria local incremental.

O segundo commit encerra a fase com esta atualização factual do histórico; seu
identificador é informado no relatório da execução, pois um commit não pode
registrar o próprio hash em seu conteúdo.

### Problemas encontrados

Na primeira execução da suíte ampliada, uma asserção comparou diretamente uma
`sqlite3.Row` com uma tupla. A asserção foi ajustada para comparar a conversão
explícita, e a suíte integral passou. Nenhum bloqueio permanece.

### Pendências

Nenhuma pendência da Fase 3.

### Próximo passo

F4 — Microsoft Graph / SharePoint, somente após autorização expressa do
responsável pelo projeto.

---

# 12. HISTÓRICO DA FASE 4

## F4 — Microsoft Graph / SharePoint

**Status:** 🔴 BLOQUEADA

**Data de início:** 15/09/2026

**Data de conclusão:** —

**Quantidade de commits:** 3 (limite da fase atingido)

### Objetivo

Integrar o núcleo validado com o SharePoint Online através de mecanismos
de leitura suportados.

### Implementado

- configuração por ambiente para tenant, aplicativo, segredo, site e drive,
  validada sem expor o valor do segredo;
- autenticação OAuth 2.0 `client_credentials` contra a plataforma de identidade
  Microsoft, solicitando o escopo `.default` do Microsoft Graph;
- fonte SharePoint com chamadas Graph exclusivamente `GET` para listar arquivos
  `.xlsx`, paginar resultados, enumerar versões e baixar conteúdo histórico;
- preservação de site, drive, DriveItem ID, nome, caminho, identificador da
  versão, data/hora, autor e tamanho disponíveis;
- comentário mantido nulo quando o contrato `DriveItemVersion` usado não fornece
  esse dado, evitando atribuição indevida de metadados;
- arquivos históricos em diretório temporário descartável;
- integração automatizada da fonte com o `AuditService`, sem reescrever reader,
  comparator ou persistência;
- falhas na listagem de versões agora registradas como falhas de execução pelo
  serviço de auditoria.

### Testes executados

Comando:

`pytest -q`

Resultado:

`28 passed in 1.04s`

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

Tentativa de instalar as dependências:

`python -m pip install -r requirements.txt`

Resultado:

falhou porque o índice configurado no ambiente retornou `403 Forbidden` ao
consultar `msal`. A dependência foi eliminada: o fluxo OAuth suportado foi
implementado com a biblioteca padrão e o `requirements.txt` permaneceu enxuto.

### Critérios de aceite

[x] configuração e autenticação seguras implementadas;

[x] cliente limitado a operações de leitura no Microsoft Graph;

[x] DriveItem ID, versões, metadados disponíveis e conteúdo histórico cobertos
por testes automatizados;

[x] `AuditService` funciona com `SharePointSource` em teste isolado;

[x] motor Excel não foi reescrito;

[ ] autenticação e leitura comprovadas em SharePoint controlado;

[ ] versões reais, especialmente secundárias, enumeradas e recuperadas;

[ ] teste controlado concluído e ausência de escrita comprovada no ambiente.

### Commits

`df7943f` — Implementa fonte SharePoint somente leitura.

O segundo commit registra este bloqueio e o estado factual da fase; seu
identificador é informado no relatório da execução.

### BLOQUEIO

**Problema:** os critérios obrigatórios de prova real não podem ser executados
sem tenant, aplicativo autorizado, segredo, site, drive e uma planilha
controlada acessível.

**Causa:** o ambiente da sessão não contém credenciais SharePoint nem o cenário
corporativo controlado.

**Impacto:** não é possível afirmar que o ambiente enumera ou permite baixar as
versões secundárias necessárias, nem concluir a F4.

**Alternativas:**

1. fornecer ao ambiente as cinco variáveis documentadas em `.env.example`, com
   permissões mínimas de leitura, e indicar uma planilha controlada;
2. executar o teste controlado externamente e fornecer evidências técnicas dos
   endpoints, metadados e versões recuperadas.

**Recomendação:** disponibilizar credenciais de aplicação com menor privilégio e
uma planilha controlada; então validar uma única planilha antes de qualquer uso
operacional.

### Pendências

- comprovar acesso ao site e à biblioteca no ambiente real;
- confirmar a ordem efetiva retornada pela coleção de versões;
- confirmar enumeração e download das versões secundárias;
- registrar metadados reais e executar a auditoria controlada.

### Nova tentativa de validação — 15/09/2026

O estado do repositório e o histórico foram novamente inspecionados. As cinco
variáveis `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`,
`SHAREPOINT_CLIENT_SECRET`, `SHAREPOINT_SITE_ID` e `SHAREPOINT_DRIVE_ID`
continuam ausentes no ambiente. Nenhuma chamada ao SharePoint foi tentada sem
credenciais, e nenhuma operação de escrita foi realizada.

Testes executados nesta tentativa:

- `pytest -q` — `28 passed in 0.96s`;
- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros antes da atualização documental.

O bloqueio permanece inalterado. O terceiro e último commit permitido para a
F4 registra esta nova verificação; seu identificador é informado no relatório
da execução, pois um commit não pode registrar o próprio hash em seu conteúdo.
Novas alterações da F4 não devem ser realizadas sem decisão explícita sobre o
limite de commits da fase.

### Próximo passo

Retomar somente a validação controlada da F4 após remoção do bloqueio e decisão
do responsável sobre o limite de commits já atingido. Não iniciar a F5.

---

# 13. HISTÓRICO DA FASE 5

## F5 — Interface e Relatório

**Status:** ⬜ NÃO INICIADA

**Data de início:** —

**Data de conclusão:** —

**Quantidade de commits:** 0

### Objetivo

Disponibilizar interface simples para operação e geração do relatório
Excel consolidado.

### Implementado

Ainda não iniciado.

### Testes executados

Nenhum.

### Commits

Nenhum.

### Problemas encontrados

Nenhum.

### Pendências

Aguardar conclusão e aprovação da Fase 4.

### Próximo passo

Não autorizado.

---

# 14. HISTÓRICO DA FASE 6

## F6 — Robustez e Preparação para Produção

**Status:** ⬜ NÃO INICIADA

**Data de início:** —

**Data de conclusão:** —

**Quantidade de commits:** 0

### Objetivo

Validar robustez, desempenho, integridade, logs, empacotamento e
preparação da V1 para homologação.

### Implementado

Ainda não iniciado.

### Testes executados

Nenhum.

### Commits

Nenhum.

### Problemas encontrados

Nenhum.

### Pendências

Aguardar conclusão e aprovação da Fase 5.

### Próximo passo

Não autorizado.

---

# 15. REGISTRO DE DECISÕES TÉCNICAS

Esta seção registra somente decisões relevantes tomadas durante o
desenvolvimento.

---

## DEC-001 — Aplicação sem Node.js

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A aplicação será baseada em Python e não possuirá dependência
obrigatória de Node.js, npm ou frameworks frontend baseados nesse
ecossistema.

### Motivo

Manter execução simples e reduzir dependências desnecessárias.

---

## DEC-002 — SQLite como banco inicial

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A V1 utilizará SQLite como banco inicial.

### Motivo

Simplicidade operacional e facilidade de implantação.

### Observação

A arquitetura deverá permitir migração futura para SQL Server.

---

## DEC-003 — Banco como fonte oficial

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

O banco de auditoria será a fonte oficial da trilha consolidada.

O relatório Excel será uma representação gerada a partir do banco.

### Consequência

Alterações ou exclusão do relatório não deverão destruir a trilha
armazenada.

---

## DEC-004 — Auditoria incremental

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

Cada planilha possuirá checkpoint individual.

Exemplo:

Primeira execução:

0.84 → ... → 0.99

Checkpoint:

0.99

Execução posterior:

0.99 → ... → 1.20

Novo checkpoint:

1.20

O histórico consolidado não deverá ser reprocessado durante execução
incremental normal.

---

## DEC-005 — Identidade independente do nome

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

O nome da planilha não será utilizado como identidade técnica.

A aplicação utilizará identificadores fornecidos pelo
SharePoint/Microsoft Graph, preservando DriveItem ID e demais
identificadores necessários.

---

## DEC-006 — SharePoint somente leitura

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A aplicação não poderá realizar operações de escrita no SharePoint.

Isso inclui:

- alteração;
- exclusão;
- criação de versões;
- check-in;
- check-out;
- alteração de metadados;
- alteração de comentários.

---

## DEC-007 — Limite de commits

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

Cada fase possuirá no máximo 3 commits.

Preferência:

1 ou 2 commits.

### Motivo

Evitar desenvolvimento excessivamente fragmentado e manter o projeto
curto e controlável.

---

# 16. REGISTRO DE BLOQUEIOS

## BLOQ-001 — Validação controlada do SharePoint indisponível

**Data:** 15/09/2026

**Fase:** F4 — Microsoft Graph / SharePoint

**Status:** ATIVO

### Problema

Não é possível executar a prova real obrigatória de leitura de uma planilha
controlada, incluindo enumeração e recuperação de versões secundárias.

### Causa

O ambiente não possui as cinco variáveis de configuração SharePoint nem acesso
ao cenário corporativo controlado.

### Impacto

A F4 não atende aos critérios de aceite que exigem validação real e permanece
bloqueada. A F5 não está autorizada.

### Alternativas

1. Disponibilizar credenciais de aplicação com menor privilégio e uma planilha
   controlada no ambiente.
2. Executar o teste controlado externamente e fornecer evidências técnicas dos
   endpoints, metadados e versões recuperadas.

### Recomendação

Validar uma única planilha controlada com permissões mínimas de leitura. Como a
F4 atingiu três commits, obter também decisão explícita do responsável antes de
qualquer novo commit nessa fase.

### Decisão

Aguardando responsável.

---

# 17. REGISTRO DE LIMITAÇÕES CONFIRMADAS

Nenhuma limitação técnica da implementação foi confirmada até o momento.

As limitações descritas nos documentos anteriores que ainda dependem
de validação técnica não deverão ser registradas aqui como fatos
confirmados.

Quando uma limitação for comprovada:

## LIM-XXX — Título

**Data:**

**Fase:**

### Comportamento esperado

[...]

### Comportamento observado

[...]

### Impacto

[...]

### Tratamento adotado

[...]

---

# 18. MODELO DE ATUALIZAÇÃO DE FASE

Ao concluir uma fase, utilizar aproximadamente:

## FASE X — NOME

**Status:** 🟢 CONCLUÍDA

**Início:** DD/MM/AAAA

**Conclusão:** DD/MM/AAAA

**Commits:** X

### Implementado

- [...]
- [...]
- [...]

### Arquivos principais

- [...]
- [...]
- [...]

### Testes

Executados:

`pytest ...`

Resultado:

XX passed

### Critérios de aceite

[x] requisito 1

[x] requisito 2

[x] requisito 3

### Commits

`abcdef1` — descrição

`abcdef2` — descrição

### Decisões

[...]

### Problemas encontrados

[...]

### Pendências

[...]

### Próximo passo

Fase X+1 aguardando autorização.

---

# 19. MODELO PARA FASE BLOQUEADA

Quando uma fase não puder ser concluída:

**Status:** 🔴 BLOQUEADA

Registrar obrigatoriamente:

- último ponto concluído;
- teste que apresentou problema;
- erro observado;
- impacto;
- alternativas;
- recomendação;
- estado do repositório.

Não marcar como concluída.

Não avançar para a próxima fase.

---

# 20. REGRA PARA TESTES

Não registrar:

"Testes OK"

sem informar quais testes foram executados.

Preferir:

Comando:

`pytest`

Resultado:

`24 passed`

ou equivalente.

Testes manuais relevantes também poderão ser registrados.

---

# 21. REGRA PARA COMMITS

Registrar o identificador real do commit.

Exemplo:

`a12bc34` — Implementa motor de comparação Excel

Não inventar hashes de commits.

Se ainda não houver commit:

registrar:

"Commit ainda não realizado."

---

# 22. REGRA PARA ALTERAÇÕES DOCUMENTAIS

Pequenos ajustes documentais realizados como consequência direta da
implementação poderão ser incluídos no commit correspondente.

Mudanças relevantes de:

- requisito;
- arquitetura;
- escopo;
- segurança;

deverão ser explicitamente registradas.

---

# 23. REGRA PARA NOVAS FUNCIONALIDADES

Sugestões surgidas durante o desenvolvimento não deverão ser
automaticamente implementadas.

Registrar como:

MELHORIA FUTURA

quando não forem necessárias para os critérios de aceite da V1.

Isso evita crescimento descontrolado do projeto.

---

# 24. MELHORIAS FUTURAS

Nenhuma melhoria adicional aprovada neste momento.

Possibilidades já identificadas, mas fora do escopo automático da V1:

- SQL Server;
- auditoria automática agendada;
- processamento em lote;
- Power BI;
- dashboards;
- notificações;
- múltiplos ambientes SharePoint.

A presença nesta seção não significa autorização para implementação.

---

# 25. ESTADO ATUAL OFICIAL

**Data:** 15/09/2026

**Projeto:** Auditor de Planilhas Excel — SharePoint Online

**Versão planejada:** V1

**Fase atual:** F4 — Microsoft Graph / SharePoint bloqueada

**Implementação:** fonte Graph somente leitura e integração automatizada com o
núcleo concluídas; validação real ainda não executada

**Fases concluídas:** 3/6

**Commits da Fase 4:** 3 (limite da fase atingido)

**Bloqueios:** 1

**Próxima ação:**

Disponibilizar ambiente SharePoint controlado e credenciais mínimas de leitura
para retomar e concluir a F4. Não iniciar a F5.

---

# 26. INSTRUÇÃO AO CODEX

Antes de iniciar qualquer implementação:

1. consultar este documento;
2. consultar `01_ESPECIFICACAO_FUNCIONAL.md`;
3. consultar `02_ARQUITETURA.md`;
4. consultar `03_PLANO_DE_DESENVOLVIMENTO.md`;
5. verificar o estado real do repositório;
6. identificar a fase autorizada.

Após executar a fase:

1. atualizar este documento;
2. registrar apenas fatos reais;
3. registrar testes efetivamente executados;
4. registrar hashes reais dos commits;
5. informar bloqueios e limitações;
6. parar antes de iniciar a fase seguinte.

---

FIM DO DOCUMENTO
