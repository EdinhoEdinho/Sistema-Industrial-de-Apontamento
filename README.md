# Sistema-Industrial-de-Apontamento
Apontamento industrial com rastreabilidade e validações de processo.

## 1. Contexto de negocio

Este sistema foi criado para digitalizar um processo industrial que, anteriormente, era executado de forma manual em uma industria de adesivos, com preenchimento em papel.

### Situacao anterior (manual)

- Registros fisicos por etapa.
- Dificuldade para localizar historico de um quadro/lote.
- Risco de perda, rasura ou preenchimento incompleto.
- Baixa visibilidade para auditoria e qualidade.
- Pouca rastreabilidade entre etapas e responsaveis.

### Situacao atual (digital)

- Coleta de dados por etapa com interface unica.
- Registro de funcionario, data/hora, status e observacao.
- Validacoes automaticas para reduzir erro operacional.
- Regras de bloqueio para evitar duplicidade e retrabalho indevido.
- Rastreabilidade por `id_quadro` do inicio ao fim do fluxo.

---

## 2. Objetivo do sistema

Garantir controle de processo, qualidade e rastreabilidade do ciclo produtivo do quadro, cobrindo:

1. Esticagem
2. Pos-esticagem
3. Emulsao
4. Revelacao
5. Revelacao final (com decisao de reutilizacao)

---

## 3. Visao geral do processo

## 3.1 Esticagem

- Leitura do `ID Quadro (QR)`.
- Definicao de origem, disposicao, mesh e cola.
- Coleta de medicoes X/Y.
- Resultado de status (`OK`/`NG`) com base em limites de parametro.

## 3.2 Pos-esticagem

- Registro de medicoes X/Y e espessura.
- Validacao de campos obrigatorios.
- Registro de status e observacao.

## 3.3 Emulsao

- Leitura do QR da emulsao.
- Validacao de espessuras (pos-esticagem x emulsao).
- Registro de funcionario, data, status e observacao.

## 3.4 Revelacao

- Leitura/validacao de fotolito.
- Montagem de status inicial da revelacao.
- Registro com operador e observacao.

## 3.5 Revelacao final

- Consolidacao final do quadro.
- Definicao de `status`.
- Se `NG`, obrigatoriedade de observacao e decisao de reutilizacao (`Sim`/`Nao`).

## 3.6 Analise das etapas industriais

Esta secao descreve o processo pela otica industrial (chao de fabrica), alem da otica de sistema.

### Etapa 1 - Esticagem (formacao da base dimensional)

- Objetivo industrial:
  - garantir tensao adequada da tela para estabilidade geometrica do quadro.
- Entradas principais:
  - `id_quadro`, origem, mesh, cola, medicoes X/Y, operador.
- Risco operacional:
  - tensao fora de faixa gera deformacao, perda de repetibilidade e defeito nas etapas seguintes.
- Controle no sistema:
  - comparacao das medicoes com parametros min/max.
  - bloqueio de duplicidade quando ja existe quadro aprovado sem reset.
- Saida da etapa:
  - quadro com status `OK`/`NG` e rastreabilidade de quem mediu, quando e com quais valores.

### Etapa 2 - Pos-esticagem (confirmacao de estabilidade)

- Objetivo industrial:
  - confirmar comportamento da tela apos assentamento/tempo de processo.
- Entradas principais:
  - medicoes X/Y da pos-esticagem, espessura media, operador, observacao.
- Risco operacional:
  - aprovar quadro instavel aumenta variacao dimensional e desperdicio em emulsao/revelacao.
- Controle no sistema:
  - exige esticagem previa.
  - exige todas as medicoes e impede sobrescrever `OK` com novo `OK`.
- Saida da etapa:
  - liberacao tecnica para emulsao ou retorno para ajuste/retrabalho.

### Etapa 3 - Emulsao (formacao funcional da camada)

- Objetivo industrial:
  - aplicar e validar espessura de emulsao para desempenho de transferencia e resolucao.
- Entradas principais:
  - QR da emulsao, espessura pos-esticagem, espessura de emulsao calculada, operador.
- Risco operacional:
  - camada fora do range compromete definicao da imagem, durabilidade e repetibilidade.
- Controle no sistema:
  - validacao de espessura com base em `range_emulsao`.
  - observacao obrigatoria para `NG`.
  - bloqueio de novo registro quando ja existe `OK` sem evento de retorno.
- Saida da etapa:
  - quadro com camada validada e historico tecnico do material aplicado.

### Etapa 4 - Revelacao (validacao inicial de imagem/processo)

- Objetivo industrial:
  - confirmar que o fotolito e os parametros de revelacao geraram resultado aderente.
- Entradas principais:
  - QR do fotolito, dados de modelo/apelido, operador, status/observacao.
- Risco operacional:
  - erro de fotolito ou revelacao inadequada pode invalidar todo o quadro.
- Controle no sistema:
  - validacao de fotolito e consistencia dos dados.
  - bloqueio quando ja existe revelacao `OK`.
- Saida da etapa:
  - aprovacao inicial para fechamento ou indicacao de nao conformidade.

### Etapa 5 - Revelacao final (decisao de liberacao e reaproveitamento)

- Objetivo industrial:
  - consolidar decisao final de uso do quadro na producao.
- Entradas principais:
  - status final, observacao, decisao de reutilizacao (`Sim`/`Nao`), operador.
- Risco operacional:
  - sem regra clara de reutilizacao, quadro reprovado pode voltar ao fluxo indevidamente.
- Controle no sistema:
  - exige etapa anterior concluida.
  - para `NG`, obriga observacao e decisao de reutilizacao.
  - regra de bloqueio quando `reutilizar = Nao`.
- Saida da etapa:
  - decisao formal de liberacao, retrabalho ou descarte/restricao de reutilizacao.

### Leitura integrada do processo

- O fluxo implementa uma logica de "gate de qualidade" por etapa.
- Cada fase so deve avancar quando os controles da fase anterior estao consistentes.
- O principal ganho industrial nao e apenas digitalizar formularios, mas reduzir variacao de processo e aumentar confiabilidade das decisoes de liberacao.

---

## 4. Arquitetura tecnica

## 4.1 Stack

- Interface: `Tkinter` (Python desktop app)
- Banco: `SQL Server`
- Conexao: `pyodbc` + `ODBC Driver 17 for SQL Server`

## 4.2 Conexao de banco

- Conexao via string ODBC.
- Operacoes de leitura e gravacao por cursores SQL.
- Persistencia principal com `INSERT` nas tabelas de processo.

## 4.3 Tabelas usadas (principais)

- `dbo.esticagem`
- `dbo.pos_esticagem`
- `dbo.emulsao`
- `revelacao`
- `revelacao_final`
- `parametro_esticagem`
- `parametro_pos_esticagem`
- `range_emulsao`
- `funcionario_nome`
- `mp_sistema`
- `origem_BD`

---

## 5. Regras de validacao e condicoes

## 5.1 Regras gerais

- Matrícula de funcionario:
  - QR deve terminar com `010101`.
  - Extracao de matricula numerica (remove zeros a esquerda).
  - Nome carregado da tabela `funcionario_nome`.
- Campos obrigatorios por etapa.
- Se status = `NG`, observacao obrigatoria em varias etapas.

## 5.2 Validacao de ID Quadro

- `ID Quadro` valido deve terminar com `CV050201`.
- Caso invalido: bloqueia carregamento e solicita nova leitura.

## 5.3 Validacao numerica de medicoes

- Entradas aceitam decimal com `,` ou `.`.
- Normalizacao e comparacao com limites min/max carregados de parametros.
- Fora de faixa: campo em vermelho + status `NG`.

## 5.4 Validacao de Mesh

- QR Mesh:
  - minimo 6 caracteres para referencia.
  - sufixo obrigatorio `MPI`.
  - consulta em `mp_sistema` com `Tipo = Mesh`.
- Regra adicional:
  - se origem = `Importado`, mesh deve estar na lista permitida.
  - se mesh contem `115`, disposicao `Bias` e bloqueada.

## 5.5 Validacao de Cola

- QR Cola:
  - referencia numerica no inicio.
  - sufixo `MPI`.
  - consulta em `mp_sistema` com `Tipo = Cola`.

## 5.6 Regras de bloqueio por etapa

### Esticagem

- Se ja existe `OK` em esticagem para o mesmo quadro:
  - so libera novo salvamento se ultima pos-esticagem for `NG`.
- Se ultima pos-esticagem = `OK`: bloqueia novo salvamento.

### Pos-esticagem

- Exige que exista esticagem para o quadro.
- Se ultimo status em pos-esticagem ja for `OK`, nao permite sobrescrever com novo `OK`.
- Todas medicoes obrigatorias e numericas.

### Emulsao

- Exige esticagem previa.
- Se ultima emulsao = `OK` e ultima revelacao nao = `NG`, bloqueia novo registro.
- Para `NG`, observacao obrigatoria.
- Regras de carga:
  - se revelacao final anterior for `NG + Sim`, emulsao so carrega se foi refeita depois.

### Revelacao

- Exige quadro valido.
- Nao permite novo salvamento se ultimo status de revelacao ja for `OK`.

### Revelacao final

- Exige existencia de registro em `revelacao`.
- Campos obrigatorios: quadro, fotolito, funcionario e status.
- Se `NG`:
  - observacao obrigatoria;
  - `reutilizar quadro` obrigatorio (`Sim`/`Nao`).
- Se `OK`: reutilizacao pode ficar `N/D`.

## 5.7 Regra de reutilizacao do quadro

- Funcao de verificacao consulta `revelacao_final`.
- Se ultimo registro marcar `reutilizar = Nao`, o sistema bloqueia continuidade para aquele quadro.

---

## 6. Rastreabilidade implementada

Cada apontamento gravado registra, no minimo:

- `id_quadro`
- operador (`reg_func`)
- data/hora
- status
- observacao
- dados tecnicos da etapa (medicoes, espessuras, insumos lidos por QR)

Com isso, e possivel:

- reconstruir historico completo do quadro;
- identificar quem executou cada etapa;
- auditar desvios (`NG`) com contexto;
- controlar reprocessos e reaproveitamento.

---

## 7. Operacao em modo offline

O sistema foi ajustado para abrir mesmo sem conexao com SQL Server.

### Comportamento offline

- Interface abre normalmente.
- Usuario recebe aviso de `Modo Offline`.
- Listas vindas do banco (ex.: origem) podem vir vazias.
- Operacoes que dependem de banco podem falhar ao salvar (com aviso/erro), pois nao ha persistencia.

### Uso recomendado

- Offline serve para manter disponibilidade da interface.
- Para producao oficial e rastreabilidade efetiva, o modo online com banco deve estar ativo.

---

## 8. Fluxo resumido de uso (operacao)

1. Ler `ID Quadro`.
2. Registrar esticagem e salvar.
3. Registrar pos-esticagem e salvar.
4. Registrar emulsao e salvar.
5. Registrar revelacao e salvar.
6. Registrar revelacao final e definir reutilizacao quando necessario.

---

## 9. Ganhos obtidos com a digitalizacao

- Eliminacao de papeis fisicos como fonte primaria.
- Padronizacao de apontamento e validacao.
- Menor risco de retrabalho por dado incompleto.
- Maior velocidade de consulta historica.
- Base de dados para analise de qualidade e melhoria continua.

---

## 10. Recomendacoes tecnicas (evolucao)

- Externalizar credenciais de banco para variaveis de ambiente.
- Implementar logs estruturados de auditoria.
- Criar testes para regras criticas (duplicidade, bloqueios, `NG`).
- Separar camadas UI/negocio/repositorio para facilitar manutencao.
- Preparar fila de sincronizacao para uso offline com posterior envio.

