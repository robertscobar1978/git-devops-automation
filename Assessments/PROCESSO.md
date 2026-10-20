# Assessment SQL Server – Solo Network

Padrão para gerar o relatório HTML dos assessments de SQL Server dos clientes SOLO, sempre no layout do
`Assessment_Calcenter_PENINSULA_SQL2017_v2.html` e com o papel timbrado Solo.

| Arquivo | Para que serve |
|---|---|
| `SQL_Server_Assessment_Profissional_v11.sql` | Script executado no cliente (somente leitura). Gera os resultsets que vão para o `.rpt`. |
| `Gerar-RelatorioAssessment_v11.ps1` | Alternativa quando dá para rodar direto da máquina: executa o script e grava o HTML pronto. |
| `Papel Solo.docx` | Papel timbrado Solo (origem das imagens do cabeçalho e do rodapé). |
| `Assessment_Calcenter_PENINSULA_SQL2017_v2.html` | Modelo de saída: estrutura e estilo que todo relatório deve seguir. |
| `gerador/gerar_html_assessment.py` | Converte o `.rpt` no HTML do modelo. |
| `gerador/template/` | Cabeçalho (CSS, scripts das guias, imagens do papel Solo) e rodapé Solo, extraídos do modelo. |

## Fluxo

1. No cliente, abrir o `SQL_Server_Assessment_Profissional_v11.sql` no SSMS.
2. Deixar `@Modo = 1`, `@GerarHTML = 1` e `@SomenteHTML = 0` (padrão do script).
3. **Query > Results To > Results to File** (Ctrl+Shift+F), executar e salvar o `.rpt`.
4. Enviar o `.rpt` para o Claude (anexo na conversa ou commit em `Assessments/rpt/`), informando o nome do cliente.
5. O Claude roda o gerador e devolve o HTML:

```bash
python3 Assessments/gerador/gerar_html_assessment.py CLIENTE.rpt --cliente "Nome do Cliente"
# saída padrão: Assessment_<cliente>_<instância>.html ao lado do .rpt (ou --saida arquivo.html)
```

O gerador só usa a biblioteca padrão do Python 3. Ao terminar, ele informa quantos resultsets leu,
quantas seções montou e se algum resultset não foi reconhecido.

### Observações sobre o `.rpt`

- Aceita UTF-16 (padrão do SSMS), UTF-8 e ANSI, com colunas alinhadas (padrão) ou separadas por tab.
- O SSMS corta cada coluna em 256 caracteres no modo texto (Ferramentas > Opções > Resultados da
  Consulta > SQL Server > Resultados para Texto > *Número máximo de caracteres*). Para textos longos
  (mensagens de job, SQL das top queries) vale subir para 8192 antes de executar.
- Valores com quebra de linha (ex.: mensagem do job) são reagrupados na mesma célula.
- A coluna `Relatorio_HTML` do último resultset é ignorada (vem truncada no `.rpt`). Se o `.rpt` só
  tiver esse resultset (`@SomenteHTML = 1`) e ele estiver completo, o gerador grava esse HTML.
- Funciona também com `.rpt` de versões anteriores do script (v8 a v10): as seções que não existem
  na versão simplesmente não aparecem e a capa mostra a versão detectada.

## Mapeamento `.rpt` → HTML

O `.rpt` não traz o nome das seções, só cabeçalhos de colunas. Cada resultset é reconhecido pelas
primeiras colunas (catálogo `SECOES` no gerador, espelho do `#Secoes` do script) e vira um bloco
`<details>` na guia correspondente, com o mesmo título, descrição e numeração de grupo do script.

| Resultset (primeiras colunas) | Grupo · Seção | Guia |
|---|---|---|
| `ID, Categoria (IDENTIFICAÇÃO…)` | Dashboard executivo | KPIs da capa |
| `ID, Categoria (01 - …)` | Checklist (`#AssessmentSummary`) | Verificações de cada guia, Atenção e Checklist |
| `MachineName, ServerName` | 01 · Propriedades da instância | Visão Geral |
| `LogicalCPUCount` | 02 · Hardware | Visão Geral |
| `host_platform` | 02 · Sistema operacional | Visão Geral |
| `servicename` | 02 · Serviços do SQL Server | Visão Geral |
| `NodeName` / `VerboseLogging` | 03 · Cluster (FCI) | Alta Disponibilidade |
| `AGName, …` / `cluster_name` | 04 · Always On (AGs, réplicas, databases, quórum) | Alta Disponibilidade |
| `Information` | 03/04/20/25 · seção informativa (FCI, Always On, fragmentação, log shipping) | conforme o grupo |
| `name, product` | 05 · Linked Servers | Conexões |
| `name, endpoint_id` | 06 · Endpoints | Conexões |
| `database_id, DatabaseName` | 07 · Databases | Databases |
| `DatabaseName, file_id, LogicalName` | 08 · Arquivos de dados e log | Armazenamento |
| `volume_mount_point` | 08 · Volumes | Armazenamento |
| `fixed_drive_path` | 08 · Unidades fixas | Armazenamento |
| `name, physical_name` | 09 · TempDB - arquivos | Databases |
| `DatabaseName, reserved_page_count` | 09 · Version store | Databases |
| `DatabaseName, recovery_model_desc` | 10 · Backups e último CHECKDB | Backups |
| `JobName` | 11 · Jobs do SQL Agent | Jobs |
| `name, event_source` | 11 · Alertas | Jobs |
| `LoginName` | 12 · Logins | Segurança |
| `name, ConfiguredValue` / `configuration_id, name` | 13 · sp_configure / pendentes | Performance |
| `SQLMemoryUsedMB` / `PhysicalMemoryMB, Available…` / `ObjectName, instance_name` / `SQLServerProcessCPU` | 14 · Memória, PLE e histórico de CPU | Performance |
| `DatabaseName, file_id, physical_name` / `IORank` / `LogDate, ProcessInfo` | 15 · Latência e volume de I/O, avisos de I/O | Performance |
| `DatabaseName, file_id, page_id` | 16 · Suspect pages | Backups |
| `TraceFlag` | 16 · Trace flags | Performance |
| `BufferPoolRank` | 17 · Buffer pool | Performance |
| `DatabaseName, LogSizeMB` / `DatabaseName, file_id, VLFCount` | 17 · Uso do log e VLFs | Databases |
| `wait_type` | 18 · Top 20 wait stats | Performance |
| `PolLogin` / `SrvRoleName` / `PermPrincipal` | 19 · Política de senha, server roles, permissões | Segurança |
| `ObjDatabase` | 20 · Objetos por database | Databases |
| `FragResDatabase` / `FragDatabase` / `MIDatabase` | 20 · Fragmentação e missing indexes | Performance |
| `TQDatabase` / `TCDatabase` | 21 · Top queries | Performance |
| `SessSessionId` | 22 · Sessões | Conexões |
| `BkDatabase` / `LogBkDatabase` | 23 · Histórico de backups | Backups |
| `SchJobName` | 24 · Agendamentos dos jobs | Jobs |
| `LSPDatabase` / `LSSDatabase` / `LSMPapel` / `LSJFuncao` / `LSEData` | 25 · Log shipping | Alta Disponibilidade |

As linhas do checklist vão para a guia pelo prefixo da categoria (`01`/`02` Visão Geral, `07`/`09`/`17`
Databases, `13`/`14`/`15`/`19`/`20` Performance, `10`/`16`/`22` Backups, `12`/`18` Segurança, `11` Jobs,
`08` Armazenamento, `05`/`06`/`21` Conexões, `03`/`04`/`23` Alta Disponibilidade), igual ao script.

### Blocos calculados

| Bloco | Origem no `.rpt` |
|---|---|
| Capa: servidor, versão, edição, collation, autenticação | 01 · Propriedades da instância (+ nome da versão do dashboard) |
| Capa: SO e virtualização | 02 · Sistema operacional e Hardware |
| Data da coleta | Mensagem `Data:` do script; senão `CollectionTime` do PLE ou a última amostra de CPU |
| Resumo executivo (gauge e contadores) | Status do checklist: % OK sobre OK + ATENÇÃO + CRÍTICO |
| KPI CPUs / memória / uptime | 02 · Hardware + `max server memory` do sp_configure |
| KPI databases e tamanho alocado | 07 · Databases + 08 · Arquivos (SizeMB por tipo) |
| KPI Always On, jobs, linked servers | Dashboard + 04, 11 e 25 |
| Capacidade: volumes / maiores databases | 08 · Volumes (TotalGB, AvailableGB) / 07 · Databases (dados + log, top 10) |
| Gráfico de CPU | 14 · Histórico de CPU (ring buffer) |
| Observações da coleta | Origem do arquivo, erros `Msg …` encontrados no `.rpt` e resultsets não reconhecidos |

### Formatação (igual ao script)

- Datas `yyyy-mm-dd hh:mm:ss.mmm` → `dd/mm/aaaa hh:mm:ss`; decimais → `1.234,56`; inteiros sem separador.
- Colunas bit (`is_*`, `JobEnabled`, `Pol*`, …) → Sim/Não; números alinhados à direita.
- Status conhecidos viram selos coloridos: ONLINE/Running/Sucesso (verde), ATENÇÃO/RESTORING (laranja),
  CRÍTICO/OFFLINE/Falha (vermelho). Vazio/NULL → `—`.
- Tabelas com 1 linha e mais de 6 colunas viram Propriedade | Valor. Máximo de 500 linhas por tabela;
  nos jobs, os do Report Server vão para o fim.

## Quando o script mudar

Se uma versão nova do `.sql` criar uma seção, o gerador avisa `resultset não mapeado` e mostra a tabela
na Visão Geral. Para encaixá-la no lugar certo, acrescentar uma linha em `SECOES` (primeiras colunas,
grupo, título, descrição, colunas largas) e, se tiver colunas bit, incluí-las em `COLUNAS_BIT`.

Teste de regressão (reconstrói um `.rpt` a partir do HTML modelo e gera de novo):

```bash
python3 Assessments/gerador/testes/html_para_rpt.py Assessments/Assessment_Calcenter_PENINSULA_SQL2017_v2.html /tmp/teste.rpt
python3 Assessments/gerador/gerar_html_assessment.py /tmp/teste.rpt --cliente Calcenter --saida /tmp/teste.html
```
